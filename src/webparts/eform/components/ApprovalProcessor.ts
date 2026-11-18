/* =========================================================
   ApprovalProcessor.ts
   ONE shared approval engine used by both decision screens:
   - TeamRequestsPage        (TeamService.submitDecision)
   - ExportPDF / detail view (ApprovalDecisionService.submitApprovalDecision)

   Previously these two paths were near-duplicate 300–400 line functions
   that had to be kept byte-identical by hand (and historically drifted).
   They are now thin wrappers around processApproval().

   Behaviour is preserved from the (more complete) TeamService path:
   - out-of-stock does NOT block approval → line approved at qty 0 + shortage email
   - Quantity is never modified; approved qty is recorded in ApprovalCount
   - header TotalQuantity / TotalAmount are SUMMED from approved quantities
   - RaiseCount / ApprovalCount / StockAvailability / LineTotal columns
   - ApprovalChainJSON intermediate multi-stage hand-off

   Performance:
   - stock check: deduped by SKU, run with a concurrency limit (no sequential loop,
     no unbounded Promise.all)
   - item writes: one SharePoint $batch per 100 items (retry on 429/503)
   - header written once; items read once
========================================================= */

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { HttpClient, SPHttpClient } from "@microsoft/sp-http";
import {
  fetchProductFromAws,
  normalizeFlowToRow,
} from "./pages/CreateRequestPage/CreateRequestService";
import {
  computeHeaderTotals,
  computeItemFields,
  mapWithConcurrency,
  SkuMismatch,
} from "./ApprovalCalc";

// Re-exported so existing importers keep their paths.
export {
  approvedQtyFor,
  computeHeaderTotals,
  computeItemFields,
  mapWithConcurrency,
} from "./ApprovalCalc";
export type { SkuMismatch } from "./ApprovalCalc";

/* ===== Types ===== */

export type FlowAction = "Approve" | "Reject" | "Bypass";

export interface ApprovalLevel {
  order: number;
  approverEmail: string; // kept in the caller's casing for writing; compared case-insensitively
  canApproveRequest: boolean;
}

export interface RequestItemRow {
  Id: number;
  SKUCode?: string;
  Quantity?: number;
  UnitPrice?: number;
  RowStatus?: string;
}

export interface ProcessApprovalParams {
  context: WebPartContext;
  siteUrl: string;
  approvalFlowUrl: string;
  flowApiKey: string;
  productFlowUrl?: string;

  requestId: number; // SharePoint item Id
  transactionId: string;
  actorEmail: string;
  position: string;
  action: FlowAction;
  comments: string;
  purpose?: string;
  status?: string;

  // Approval chain — supplied by each caller (Team from the loaded row,
  // ExportPDF by reading ApprovalChainJSON from SharePoint).
  approvalLevels: ApprovalLevel[];
  hasMultipleApprovals: boolean;
  workflowStage: number;

  // Fallback header totals when no stock adjustment is made.
  requestedTotalQty: number;
  requestedTotalAmount: number;
}

export interface ProcessApprovalResult {
  success: boolean;
  message: string;
  newStatus?: string;
  finalTotalQty: number;
  finalTotalAmount: number;
  // Intermediate multi-stage hand-off (so the UI wrapper can update its row state).
  isIntermediate: boolean;
  nextApproverEmail?: string;
  nextApproverId?: number;
  nextStage?: number;
}

/* =========================================================
   Small async utilities
========================================================= */

const sleep = (ms: number): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, ms));

const toUserMessage = (value: unknown, fallback: string): string => {
  if (typeof value === "string" && value.trim()) return value.trim();
  return fallback;
};

/* =========================================================
   SharePoint reads / writes
========================================================= */

async function getActiveRequestItems(
  context: WebPartContext,
  siteUrl: string,
  requestId: number
): Promise<RequestItemRow[]> {
  const url =
    `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items` +
    `?$top=5000` +
    `&$select=Id,SKUCode,Quantity,UnitPrice,RequestIDId,RowStatus` +
    `&$filter=RequestIDId eq ${requestId} and RowStatus eq 'Active'` +
    `&$orderby=LineNumber asc`;
  const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!res.ok) throw new Error(`Failed to load request items (${res.status})`);
  const json = await res.json();
  return (json?.value || []) as RequestItemRow[];
}

async function spMergeHeader(
  context: WebPartContext,
  siteUrl: string,
  requestId: number,
  fields: Record<string, unknown>
): Promise<void> {
  await context.spHttpClient.post(
    `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${requestId})`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "odata-version": "",
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE",
      },
      body: JSON.stringify(fields),
    }
  );
}

const genGuid = (): string =>
  "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });

/** Update many InventoryRequestItems rows via SharePoint $batch — one request per
 *  100 items (SP limit), with retry+backoff on 429/503. Best-effort: throws only
 *  if a whole batch request fails after retries; the caller logs and continues. */
async function spBatchMerge(
  context: WebPartContext,
  siteUrl: string,
  listTitle: string,
  updates: Array<{ itemId: number; fields: Record<string, number> }>,
  chunkSize = 100
): Promise<void> {
  const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;
  for (let i = 0; i < updates.length; i += chunkSize) {
    const chunk = updates.slice(i, i + chunkSize);
    await sendBatchWithRetry(context, siteUrl, listUrl, chunk, 0);
  }
}

async function sendBatchWithRetry(
  context: WebPartContext,
  siteUrl: string,
  listUrl: string,
  chunk: Array<{ itemId: number; fields: Record<string, number> }>,
  attempt: number
): Promise<void> {
  const MAX_ATTEMPTS = 4;
  const batchGuid = `batch_${genGuid()}`;
  const changesetGuid = `changeset_${genGuid()}`;

  const lines: string[] = [];
  lines.push(`--${batchGuid}`);
  lines.push(`Content-Type: multipart/mixed; boundary="${changesetGuid}"`);
  lines.push("");
  chunk.forEach((u) => {
    lines.push(`--${changesetGuid}`);
    lines.push("Content-Type: application/http");
    lines.push("Content-Transfer-Encoding: binary");
    lines.push("");
    lines.push(`MERGE ${listUrl}(${u.itemId}) HTTP/1.1`);
    lines.push("Content-Type: application/json;odata=nometadata");
    lines.push("IF-MATCH: *");
    lines.push("");
    lines.push(JSON.stringify(u.fields));
  });
  lines.push(`--${changesetGuid}--`);
  lines.push(`--${batchGuid}--`);
  lines.push("");
  const body = lines.join("\r\n");

  const res = await context.spHttpClient.post(
    `${siteUrl}/_api/$batch`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        "Content-Type": `multipart/mixed; boundary="${batchGuid}"`,
        Accept: "application/json",
        "odata-version": "",
      },
      body,
    }
  );

  // Whole-batch throttling → wait (Retry-After if given, else exp. backoff) and retry.
  if ((res.status === 429 || res.status === 503) && attempt < MAX_ATTEMPTS) {
    const retryAfter = Number(res.headers.get("Retry-After"));
    const delay =
      Number.isFinite(retryAfter) && retryAfter > 0
        ? retryAfter * 1000
        : Math.min(8000, 500 * Math.pow(2, attempt));
    await sleep(delay);
    return sendBatchWithRetry(context, siteUrl, listUrl, chunk, attempt + 1);
  }

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`$batch failed (${res.status}): ${text.slice(0, 500)}`);
  }

  // A 200 can still contain per-operation failures inside the multipart body.
  const text = await res.text().catch(() => "");
  if (/HTTP\/1\.1 (4\d\d|5\d\d)/.test(text)) {
    // eslint-disable-next-line no-console
    console.warn("[ApprovalProcessor] some $batch operations reported errors:", text.slice(0, 1000));
  }
}

/* =========================================================
   Stock validation (deduped by SKU, concurrency-limited)
========================================================= */

export async function validateStock(
  context: WebPartContext,
  siteUrl: string,
  requestId: number,
  productFlowUrl: string | undefined,
  concurrency = 5
): Promise<{ skuMismatches: SkuMismatch[]; allItemResults: SkuMismatch[]; quantityShort: boolean }> {
  const items = await getActiveRequestItems(context, siteUrl, requestId);
  const skuMismatches: SkuMismatch[] = [];
  const allItemResults: SkuMismatch[] = [];

  const skuOf = (it: RequestItemRow): string => String(it.SKUCode || "-").trim() || "-";

  // Zero-qty lines are always a shortage; no stock call is made, so liveStockQty
  // stays -1 (the write then skips StockAvailability instead of wiping it).
  for (const it of items) {
    if (Number(it.Quantity || 0) === 0) {
      const entry: SkuMismatch = {
        itemId: Number(it.Id || 0),
        sku: skuOf(it),
        tagPricePerUnit: Number(it.UnitPrice || 0),
        requestedQty: 0,
        totalTagPrice: 0,
        availableQty: 0,
        actualTotalTagPrice: 0,
        differenceQty: 0,
        liveStockQty: -1,
      };
      skuMismatches.push(entry);
      allItemResults.push(entry);
    }
  }

  const positive = items.filter((it) => Number(it.Quantity || 0) > 0);

  if (!productFlowUrl) {
    // No WMS URL — cannot verify stock. Record every line (approved = requested,
    // liveStockQty -1) so the totals cover every item instead of only zero-qty lines.
    for (const it of positive) {
      const reqQty = Number(it.Quantity || 0);
      const unitPrice = Number(it.UnitPrice || 0);
      const totalTagPrice = reqQty * unitPrice;
      allItemResults.push({
        itemId: Number(it.Id || 0),
        sku: skuOf(it),
        tagPricePerUnit: unitPrice,
        requestedQty: reqQty,
        totalTagPrice,
        availableQty: reqQty,
        actualTotalTagPrice: totalTagPrice,
        differenceQty: 0,
        liveStockQty: -1,
      });
    }
    return { skuMismatches, allItemResults, quantityShort: skuMismatches.length > 0 };
  }

  // Fetch each DISTINCT SKU once, with a concurrency cap (no sequential loop).
  const distinctSkus = Array.from(new Set(positive.map(skuOf)));
  type StockLookup = { ok: boolean; hasFlow: boolean; liveStock: number };
  const looked = await mapWithConcurrency<string, { sku: string; r: StockLookup }>(
    distinctSkus,
    concurrency,
    async (sku) => {
      try {
        const json = await fetchProductFromAws(context.httpClient, sku, productFlowUrl);
        const flow = normalizeFlowToRow(json);
        return { sku, r: { ok: true, hasFlow: !!flow, liveStock: Number(flow?.Warehouse_SOH_Qty ?? 0) } };
      } catch {
        return { sku, r: { ok: false, hasFlow: false, liveStock: -1 } };
      }
    }
  );
  const stockMap = new Map<string, StockLookup>(looked.map((x) => [x.sku, x.r]));

  for (const it of positive) {
    const reqQty = Number(it.Quantity || 0);
    const unitPrice = Number(it.UnitPrice || 0);
    const sku = skuOf(it);
    const itemId = Number(it.Id || 0);
    const totalTagPrice = reqQty * unitPrice;
    const look = stockMap.get(sku);

    if (!look || !look.ok) {
      // Unverifiable (stock call failed) — -1 sentinels, requested qty is kept.
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
      continue;
    }

    const liveStock = look.liveStock;
    if (!look.hasFlow || liveStock <= 0) {
      // Out of stock (zero/negative) — approved at qty 0, NOT blocked.
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
      // Partial stock.
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
      // Sufficient stock — no mismatch, but recorded so totals cover every line.
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
  }

  return { skuMismatches, allItemResults, quantityShort: skuMismatches.length > 0 };
}

/* =========================================================
   The shared approval engine
========================================================= */

export async function processApproval(
  p: ProcessApprovalParams
): Promise<ProcessApprovalResult> {
  const { context, siteUrl, action } = p;
  const actorEmail = String(p.actorEmail || "").trim().toLowerCase();

  let quantityShort = false;
  let skuMismatches: SkuMismatch[] = [];
  let allItemResults: SkuMismatch[] = [];
  let isFinalApprover = false;
  let finalTotalQty = p.requestedTotalQty;
  let finalTotalAmount = p.requestedTotalAmount;

  if (action === "Approve") {
    // ── Stock validation (deduped + concurrency-limited) ──
    // eslint-disable-next-line no-console
    console.time("[Approval] stock fetch");
    try {
      const check = await validateStock(context, siteUrl, p.requestId, p.productFlowUrl);
      quantityShort = check.quantityShort;
      skuMismatches = check.skuMismatches;
      allItemResults = check.allItemResults;
    } finally {
      // eslint-disable-next-line no-console
      console.timeEnd("[Approval] stock fetch");
    }

    // ── Final-approver detection ──
    const activeLevels = p.approvalLevels.filter((l) => l.canApproveRequest);
    if (!p.hasMultipleApprovals) {
      isFinalApprover = true;
    } else if (activeLevels.length > 0) {
      const maxOrder = Math.max(...activeLevels.map((l) => l.order));
      isFinalApprover = activeLevels.some(
        (l) => String(l.approverEmail || "").trim().toLowerCase() === actorEmail && l.order === maxOrder
      );
    }

    // ── Adjusted header totals, computed in memory (no writes yet) ──
    if (isFinalApprover && quantityShort && skuMismatches.length > 0 && allItemResults.length > 0) {
      const totals = computeHeaderTotals(allItemResults);
      finalTotalQty = totals.totalQty;
      finalTotalAmount = totals.totalAmount;
    }
  }

  // ── Power Automate payload ──
  const siteOrigin = String(context.pageContext.web.absoluteUrl || "").replace(/\/$/, "");
  const detailViewUrl = `${siteOrigin}/SitePages/RequestDetail.aspx?requestId=${p.requestId}`;
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
    requestId: p.requestId,
    actorEmail: p.actorEmail,
    position: p.position,
    action,
    comments: p.comments || "",
    quantityShort,
    skuMismatches: payloadSkuMismatches,
    transactionId: p.transactionId,
    detailViewUrl,
    totalQuantity: finalTotalQty,
    totalTagPrice: finalTotalAmount,
    purpose: p.purpose ?? "",
    status: p.status ?? "",
  };

  const willWriteCounts = isFinalApprover && action === "Approve" && allItemResults.length > 0;

  // Persist RaiseCount/ApprovalCount/StockAvailability (+LineTotal on success) for
  // every checked item, via $batch. Best-effort — never throws out to the caller.
  const writeCounts = async (includeLineTotal: boolean): Promise<void> => {
    if (!willWriteCounts) return;
    // eslint-disable-next-line no-console
    console.time("[Approval] item writes");
    try {
      const updates = allItemResults.map((m) => ({
        itemId: m.itemId,
        fields: computeItemFields(m, includeLineTotal),
      }));
      await spBatchMerge(context, siteUrl, "InventoryRequestItems", updates, 100);
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error(
        "[ApprovalProcessor] RaiseCount/ApprovalCount/StockAvailability write FAILED. " +
          "Verify the InventoryRequestItems column internal names:",
        err
      );
    } finally {
      // eslint-disable-next-line no-console
      console.timeEnd("[Approval] item writes");
    }
  };

  // ── Call the approval flow. It is the approval engine (not just email), so its
  // response drives ok/newStatus and the gated writes — hence it is awaited. But a
  // flow error is logged and surfaced as a failed result, never thrown, so the
  // approval attempt "never fails" with an unhandled exception and still records counts. ──
  let resp: { status?: boolean | string; success?: boolean; message?: string; newStatus?: string } = {};
  let flowErrored = false;
  try {
    const res = await context.httpClient.post(p.approvalFlowUrl, HttpClient.configurations.v1, {
      headers: { "Content-Type": "application/json", "x-api-key": p.flowApiKey },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      flowErrored = true;
      const text = await res.text().catch(() => "");
      // eslint-disable-next-line no-console
      console.error(`[ApprovalProcessor] Power Automate call failed (status ${res.status}).`, text);
    } else {
      try {
        resp = await res.json();
      } catch {
        resp = {};
      }
    }
  } catch (err) {
    flowErrored = true;
    // eslint-disable-next-line no-console
    console.error("[ApprovalProcessor] Power Automate call threw:", err);
  }

  const ok =
    !flowErrored &&
    (typeof resp?.status === "boolean"
      ? resp.status
      : String(resp?.status || "").toLowerCase() === "true"
        ? true
        : Boolean(resp?.success));
  const message = toUserMessage(resp?.message, ok ? "Action completed successfully." : "Action failed.");
  const newStatus = resp?.newStatus ? String(resp.newStatus) : undefined;

  if (!ok) {
    // Flow error or business failure: record the counts (no LineTotal) only.
    await writeCounts(false);
    return {
      success: false,
      message,
      newStatus,
      finalTotalQty,
      finalTotalAmount,
      isIntermediate: false,
    };
  }

  // ── Approved: persist adjusted header totals (only when short-stock adjusted). ──
  if (isFinalApprover && action === "Approve" && quantityShort && skuMismatches.length > 0) {
    // eslint-disable-next-line no-console
    console.time("[Approval] header write");
    try {
      await spMergeHeader(context, siteUrl, p.requestId, {
        TotalQuantity: Math.max(0, Number(finalTotalQty) || 0),
        TotalAmount: Math.max(0, Number(finalTotalAmount) || 0),
      });
    } finally {
      // eslint-disable-next-line no-console
      console.timeEnd("[Approval] header write");
    }
  }

  await writeCounts(true);

  // ── Intermediate multi-stage hand-off (advance to the next approver). ──
  if (action === "Approve" && p.hasMultipleApprovals) {
    const approveable = p.approvalLevels
      .filter((l) => l.canApproveRequest)
      .sort((a, b) => a.order - b.order);
    const currentLevelIdx = approveable.findIndex(
      (l) => String(l.approverEmail || "").trim().toLowerCase() === actorEmail
    );
    const isIntermediate =
      currentLevelIdx >= 0 && currentLevelIdx < approveable.length - 1;

    if (isIntermediate) {
      try {
        const nextLevel = approveable[currentLevelIdx + 1];
        const nextStage = p.workflowStage + 1;

        // Resolve the next approver's SharePoint user id (non-blocking).
        let nextApproverId: number | undefined;
        try {
          const ensureRes = await context.spHttpClient.post(
            `${siteOrigin}/_api/web/ensureUser('${nextLevel.approverEmail.replace(/'/g, "''")}')`,
            SPHttpClient.configurations.v1,
            {
              headers: { "Content-Type": "application/json;odata=verbose" },
              body: JSON.stringify({}),
            }
          );
          if (ensureRes.ok) {
            const ensureJson = (await ensureRes.json()) as { Id?: number };
            nextApproverId = ensureJson.Id;
          }
        } catch {
          /* non-blocking */
        }

        await spMergeHeader(context, siteUrl, p.requestId, {
          WorkflowStatus: "In Approval",
          WorkflowStage: nextStage,
          CurrentApproverEmail: nextLevel.approverEmail,
          ...(nextApproverId ? { CurrentApproverId: nextApproverId } : {}),
          LastUpdatedOn: new Date().toISOString(),
        });

        return {
          success: true,
          message,
          newStatus: "In Approval",
          finalTotalQty,
          finalTotalAmount,
          isIntermediate: true,
          nextApproverEmail: nextLevel.approverEmail,
          nextApproverId,
          nextStage,
        };
      } catch (err) {
        // Non-fatal: the approval already succeeded in Power Automate.
        // eslint-disable-next-line no-console
        console.error("[ApprovalProcessor] intermediate stage hand-off failed:", err);
      }
    }
  }

  return {
    success: ok,
    message,
    newStatus,
    finalTotalQty,
    finalTotalAmount,
    isIntermediate: false,
  };
}
