/* =========================================================
   ApprovalDecisionService.ts
   Shared approve/reject logic used by:
   - TeamRequestsPage (via useTeamRequestsLogic)
   - ExportPDFPage (via Eform.tsx submitExportDecision)
   
   Validates stock, builds skuMismatches[], and calls Power Automate.
========================================================= */

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { HttpClient, SPHttpClient } from "@microsoft/sp-http";
import {
  fetchProductFromAws,
  normalizeFlowToRow,
  spUpdateItem,
} from "./CreateRequestPage/CreateRequestService";

export type FlowAction = "Approve" | "Reject" | "Bypass";

export interface SkuMismatch {
  itemId: number;
  sku: string;
  tagPricePerUnit: number;
  requestedQty: number;
  totalTagPrice: number;
  availableQty: number;
  actualTotalTagPrice: number;
  differenceQty: number;
  // Actual live warehouse stock (Warehouse_SOH_Qty) from the API — used for the
  // StockAvailability column. Distinct from availableQty (which drives ApprovalCount).
  liveStockQty: number;
}

export interface RequestItemRow {
  Id: number;
  SKUCode?: string;
  Quantity?: number;
  UnitPrice?: number;
  RowStatus?: string;
}

export interface DecisionInput {
  context: WebPartContext;
  approvalFlowUrl: string;
  apiKey: string;
  productFlowUrl?: string;

  requestId: number;            // SP Id
  transactionId: string;        // Display TX ID (Title)
  actorEmail: string;
  position: string;
  action: FlowAction;
  comments: string;

  // Optional — passed when caller already has them
  totalQuantity?: number;
  totalTagPrice?: number;
  purpose?: string;
  status?: string;
}

export interface DecisionResult {
  success: boolean;
  message: string;
  newStatus?: string;
}

/* =========================================================
   Fetch active line items for stock validation
========================================================= */
async function getActiveRequestItems(
  context: WebPartContext,
  requestId: number
): Promise<RequestItemRow[]> {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const url =
    `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items` +
    `?$select=Id,SKUCode,Quantity,UnitPrice,RowStatus` +
    `&$filter=RequestIDId eq ${requestId} and RowStatus eq 'Active'` +
    `&$orderby=LineNumber asc`;

  const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!res.ok) {
    throw new Error(`Failed to load request items (${res.status})`);
  }
  const json = await res.json();
  return (json?.value || []) as RequestItemRow[];
}

/* =========================================================
   Validate stock and build skuMismatches[]
========================================================= */
export async function validateStockForApproval(
  context: WebPartContext,
  requestId: number,
  productFlowUrl?: string
): Promise<{
  skuMismatches: SkuMismatch[];
  quantityShort: boolean;
  allItemResults: SkuMismatch[];
  hasZeroStockItem: boolean;
}> {
  if (!productFlowUrl) {
    return { skuMismatches: [], quantityShort: false, allItemResults: [], hasZeroStockItem: false };
  }

  const items = await getActiveRequestItems(context, requestId);
  const skuMismatches: SkuMismatch[] = [];
  // Every checked item (mismatch or not) — used to persist ApprovalCount/RaiseCount
  // for every SKU on final approval, not just mismatched ones.
  const allItemResults: SkuMismatch[] = [];
  // True when any SKU's live stock came back as zero/unavailable — blocks approval entirely.
  let hasZeroStockItem = false;

  for (const it of items) {
    const reqQty = Number(it.Quantity || 0);
    const unitPrice = Number(it.UnitPrice || 0);
    const sku = String(it.SKUCode || "-").trim() || "-";
    const itemId = Number(it.Id || 0);
    const totalTagPrice = reqQty * unitPrice;

    try {
      const json = await fetchProductFromAws(context.httpClient, sku, productFlowUrl);
      const flow = normalizeFlowToRow(json);
      const liveStock = Number(flow?.Warehouse_SOH_Qty ?? 0);

      if (!flow || liveStock <= 0) {
        hasZeroStockItem = true;
        const entry: SkuMismatch = {
          itemId, sku,
          tagPricePerUnit: unitPrice,
          requestedQty: reqQty,
          totalTagPrice,
          availableQty: 0,
          actualTotalTagPrice: 0,
          differenceQty: 0 - reqQty,
          liveStockQty: 0,
        };
        skuMismatches.push(entry);
        allItemResults.push(entry);
      } else if (reqQty > liveStock) {
        const entry: SkuMismatch = {
          itemId, sku,
          tagPricePerUnit: unitPrice,
          requestedQty: reqQty,
          totalTagPrice,
          availableQty: liveStock,
          actualTotalTagPrice: liveStock * unitPrice,
          differenceQty: liveStock - reqQty,
          liveStockQty: liveStock,
        };
        skuMismatches.push(entry);
        allItemResults.push(entry);
      } else {
        // Stock OK — no mismatch entry, but still record it: ApprovalCount = the
        // requested qty itself (fully available); StockAvailability = actual live stock.
        allItemResults.push({
          itemId, sku,
          tagPricePerUnit: unitPrice,
          requestedQty: reqQty,
          totalTagPrice,
          availableQty: reqQty,
          actualTotalTagPrice: totalTagPrice,
          differenceQty: 0,
          liveStockQty: liveStock,
        });
      }
    } catch {
      // Unverifiable
      const entry: SkuMismatch = {
        itemId, sku,
        tagPricePerUnit: unitPrice,
        requestedQty: reqQty,
        totalTagPrice,
        availableQty: -1,
        actualTotalTagPrice: -1,
        differenceQty: -1,
        liveStockQty: -1,
      };
      skuMismatches.push(entry);
      allItemResults.push(entry);
    }
  }

  return { skuMismatches, quantityShort: skuMismatches.length > 0, allItemResults, hasZeroStockItem };
}

/* =========================================================
   Submit decision to Power Automate
========================================================= */
export async function submitApprovalDecision(
  input: DecisionInput
): Promise<DecisionResult> {
  if (!input.approvalFlowUrl) throw new Error("Approval flow URL is missing");
  if (!input.apiKey) throw new Error("Security key is missing");

  // Stock validation only on Approve
  let quantityShort = false;
  let skuMismatches: SkuMismatch[] = [];
  // Every checked item (mismatch or not) — used to persist RaiseCount/ApprovalCount
  // for every SKU on final approval, not just mismatched ones.
  let allItemResults: SkuMismatch[] = [];
  // Final approver = the one whose action should persist mismatch data to
  // SharePoint. Single-approver records: any approver is final. Multi-stage
  // records: only the approver at the highest order is final.
  let isFinalApprover = false;

  if (input.action === "Approve") {
    const check = await validateStockForApproval(
      input.context,
      input.requestId,
      input.productFlowUrl
    );
    quantityShort = check.quantityShort;
    skuMismatches = check.skuMismatches;
    allItemResults = check.allItemResults;

    // Hard block: any SKU with zero live stock cannot be approved at all.
    if (check.hasZeroStockItem) {
      throw new Error("Stock availability for the particular record is zero, so you can't approve this record.");
    }
  }

  const siteOrigin = String(input.context.pageContext.web.absoluteUrl || "").replace(/\/$/, "");
  const detailViewUrl = `${siteOrigin}/SitePages/RequestDetail.aspx?requestId=${input.requestId}`;

  // Mirror TeamService: update item quantities and header totals before calling PA
  let finalTotalQuantity = input.totalQuantity ?? 0;
  let finalTotalTagPrice = input.totalTagPrice ?? 0;

  // Determine if the current user is the final approver — same logic as TeamService.
  // Runs for EVERY approval (not only on mismatch) so ApprovalCount / StockAvailability
  // is still persisted when quantities fully match available stock.
  if (input.action === "Approve") {
    try {
      const chainRes = await input.context.spHttpClient.get(
        `${siteOrigin}/_api/web/lists/getbytitle('InventoryRequests')/items(${input.requestId})` +
          `?$select=ApprovalChainJSON,HasMultipleApprovals`,
        SPHttpClient.configurations.v1
      );
      if (chainRes.ok) {
        const chainData = await chainRes.json() as {
          ApprovalChainJSON?: string;
          HasMultipleApprovals?: boolean;
        };
        const hasMultiple = chainData?.HasMultipleApprovals === true;
        let levels: Array<{ order: number; approverEmail: string; canApproveRequest: boolean }> = [];
        try {
          const parsed = JSON.parse(String(chainData?.ApprovalChainJSON || "")) as {
            allApprovalLevels?: Array<{
              order?: number;
              approverEmail?: string;
              canApproveRequest?: boolean;
            }>;
          };
          if (Array.isArray(parsed?.allApprovalLevels)) {
            levels = parsed.allApprovalLevels
              .map((l) => ({
                order: Number(l?.order ?? 0),
                approverEmail: String(l?.approverEmail || "").trim().toLowerCase(),
                canApproveRequest: l?.canApproveRequest === true,
              }))
              .filter((l) => l.order > 0 && !!l.approverEmail);
          }
        } catch { /* ignore chain parse errors */ }

        const actorEmail = String(input.actorEmail || "").trim().toLowerCase();
        const activeLevels = levels.filter((l) => l.canApproveRequest);
        if (!hasMultiple) {
          // Single-approver record — always the final approver
          isFinalApprover = true;
        } else if (activeLevels.length > 0) {
          // Multi-approver: final = the approver at the highest order
          const maxOrder = Math.max(...activeLevels.map((l) => l.order));
          isFinalApprover = activeLevels.some(
            (l) => l.approverEmail === actorEmail && l.order === maxOrder
          );
        }
        console.log("[EForm Approval] final-approver check:", {
          requestId: input.requestId,
          actorEmail,
          hasMultiple,
          levels,
          isFinalApprover,
        });
      } else {
        console.warn(
          "[EForm Approval] Could not read ApprovalChainJSON/HasMultipleApprovals " +
            `(status ${chainRes.status}). isFinalApprover stays false — column write will be skipped.`
        );
      }
    } catch (err) {
      console.warn("[EForm Approval] final-approver detection failed (non-blocking):", err);
    }
  }

  // The final approver adjusts short-stock quantities. IMPORTANT: nothing is written
  // to SharePoint before the Power Automate call — if PA errors, no list change occurs.
  // Here we only COMPUTE the adjusted totals (in memory) so the PA payload carries the
  // final figures; the actual SharePoint writes happen only AFTER PA succeeds (below).
  const doQuantityAdjust =
    isFinalApprover && input.action === "Approve" && skuMismatches.length > 0;
  // Totals are SUMMED from the approved quantity of every line item, so they can
  // never go negative.
  if (doQuantityAdjust && allItemResults.length > 0) {
    let qtySum = 0;
    let priceSum = 0;
    for (const m of allItemResults) {
      const requested = Number.isFinite(Number(m.requestedQty)) ? Number(m.requestedQty) : 0;
      const availRaw = Number(m.availableQty);
      // Unverifiable stock (-1) keeps the requested quantity.
      const adjQty =
        !Number.isFinite(availRaw) || availRaw === -1 ? requested : Math.max(0, availRaw);
      const unitPrice = Number.isFinite(Number(m.tagPricePerUnit)) ? Number(m.tagPricePerUnit) : 0;
      qtySum += adjQty;
      priceSum += adjQty * unitPrice;
    }
    finalTotalQuantity = Math.max(0, qtySum);
    finalTotalTagPrice = Math.max(0, priceSum);
  }

  // Strip the internal-only liveStockQty field — the Power Automate trigger schema
  // rejects unknown properties (additionalProperties: false).
  const payloadSkuMismatches = skuMismatches.map((m) => ({
    itemId: m.itemId,
    sku: m.sku,
    tagPricePerUnit: m.tagPricePerUnit,
    requestedQty: m.requestedQty,
    totalTagPrice: m.totalTagPrice,
    availableQty: m.availableQty,
    actualTotalTagPrice: m.actualTotalTagPrice,
    differenceQty: m.differenceQty,
  }));

  const payload = {
    requestId: input.requestId,
    actorEmail: input.actorEmail,
    position: input.position,
    action: input.action,
    comments: input.comments || "",
    quantityShort,
    skuMismatches: payloadSkuMismatches,
    transactionId: input.transactionId,
    detailViewUrl,
    totalQuantity: finalTotalQuantity,
    totalTagPrice: finalTotalTagPrice,
    purpose: input.purpose ?? "",
    status: input.status ?? "",
  };

  const res = await input.context.httpClient.post(
    input.approvalFlowUrl,
    HttpClient.configurations.v1,
    {
      headers: {
        "Content-Type": "application/json",
        "x-api-key": input.apiKey,
      },
      body: JSON.stringify(payload),
    }
  );

  if (!res.ok) {
    const text = await res.text();
    console.error(
      `[EForm Approval] Power Automate call failed (status ${res.status}). ` +
        "Column write is skipped because it runs only after a successful approval call. " +
        "Check approvalflowurl and the API key in WorkflowConfiguration.",
      text
    );
    throw new Error(text || `Flow call failed (${res.status})`);
  }

  // PA succeeded — persist the adjusted header totals.
  // NOTE: the line-item Quantity/LineTotal are intentionally NOT overwritten — the
  // originally requested quantity stays intact and the approved figure is recorded
  // in ApprovalCount instead.
  if (doQuantityAdjust) {
    await input.context.spHttpClient.post(
      `${siteOrigin}/_api/web/lists/getbytitle('InventoryRequests')/items(${input.requestId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: JSON.stringify({
          // Hard guard: never store a negative total in SharePoint.
          TotalQuantity: Math.max(0, Number(finalTotalQuantity) || 0),
          TotalAmount: Math.max(0, Number(finalTotalTagPrice) || 0),
        }),
      }
    );
  }

  // Persist RaiseCount / ApprovalCount / StockAvailability for every checked item —
  // mismatched or not (best-effort, non-fatal). Same final-approver rule as the
  // Quantity/Total update above — intermediate approvers in a multi-stage chain
  // must not write.
  console.log("[EForm Approval] column-write gate:", {
    isFinalApprover,
    action: input.action,
    allItemResults: allItemResults.length,
    willWrite: isFinalApprover && input.action === "Approve" && allItemResults.length > 0,
  });
  if (isFinalApprover && input.action === "Approve" && allItemResults.length > 0) {
    try {
      const CHUNK = 8;
      for (let i = 0; i < allItemResults.length; i += CHUNK) {
        const chunk = allItemResults.slice(i, i + CHUNK);
        await Promise.all(
          chunk.map((m) => {
            // Coerce to finite numbers — an undefined/NaN value would be dropped or
            // sent as null by JSON.stringify, leaving the column blank.
            const requested = Number.isFinite(Number(m.requestedQty)) ? Number(m.requestedQty) : 0;
            const availRaw = Number(m.availableQty);
            // Items whose live stock could not be verified (-1) still record
            // RaiseCount/ApprovalCount from the requested quantity; only
            // StockAvailability is omitted, since the live figure is unknown.
            const approved =
              !Number.isFinite(availRaw) || availRaw === -1 ? requested : availRaw;
            const unitPrice = Number.isFinite(Number(m.tagPricePerUnit))
              ? Number(m.tagPricePerUnit)
              : 0;
            const fields: Record<string, number> = {
              RaiseCount: requested,
              ApprovalCount: approved,
              // LineTotal reflects the APPROVED quantity (Quantity itself is left as
              // originally requested), so it matches the header TotalAmount.
              LineTotal: Math.max(0, approved * unitPrice),
            };
            const liveRaw = Number(m.liveStockQty);
            if (Number.isFinite(liveRaw) && liveRaw >= 0) fields.StockAvailability = liveRaw;
            return spUpdateItem(input.context, "InventoryRequestItems", m.itemId, fields);
          })
        );
      }
      console.log(
        `[EForm Approval] Wrote RaiseCount/ApprovalCount/StockAvailability for ${allItemResults.length} item(s).`
      );
    } catch (err) {
      console.error(
        "[EForm Approval] RaiseCount/ApprovalCount/StockAvailability write FAILED. " +
          "Verify the column internal names are exactly RaiseCount, ApprovalCount, StockAvailability " +
          "on the InventoryRequestItems list:",
        err
      );
    }
  }

  let resp: { status?: boolean | string; success?: boolean; message?: string; newStatus?: string } = {};
  try {
    resp = await res.json();
  } catch {
    resp = {};
  }

  const ok =
    typeof resp?.status === "boolean"
      ? resp.status
      : String(resp?.status || "").toLowerCase() === "true"
        ? true
        : Boolean(resp?.success);

  const message =
    typeof resp?.message === "string" && resp.message.trim()
      ? resp.message.trim()
      : ok ? "Action completed successfully." : "Action failed.";

  return { success: ok, message, newStatus: resp?.newStatus };
}