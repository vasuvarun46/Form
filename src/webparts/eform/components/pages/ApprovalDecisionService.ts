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
): Promise<{ skuMismatches: SkuMismatch[]; quantityShort: boolean }> {
  if (!productFlowUrl) {
    return { skuMismatches: [], quantityShort: false };
  }

  const items = await getActiveRequestItems(context, requestId);
  const skuMismatches: SkuMismatch[] = [];

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
        skuMismatches.push({
          itemId, sku,
          tagPricePerUnit: unitPrice,
          requestedQty: reqQty,
          totalTagPrice,
          availableQty: 0,
          actualTotalTagPrice: 0,
          differenceQty: 0 - reqQty,
        });
      } else if (reqQty > liveStock) {
        skuMismatches.push({
          itemId, sku,
          tagPricePerUnit: unitPrice,
          requestedQty: reqQty,
          totalTagPrice,
          availableQty: liveStock,
          actualTotalTagPrice: liveStock * unitPrice,
          differenceQty: liveStock - reqQty,
        });
      }
      // else: stock OK — no mismatch entry
    } catch {
      // Unverifiable
      skuMismatches.push({
        itemId, sku,
        tagPricePerUnit: unitPrice,
        requestedQty: reqQty,
        totalTagPrice,
        availableQty: -1,
        actualTotalTagPrice: -1,
        differenceQty: -1,
      });
    }
  }

  return { skuMismatches, quantityShort: skuMismatches.length > 0 };
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

  if (input.action === "Approve") {
    const check = await validateStockForApproval(
      input.context,
      input.requestId,
      input.productFlowUrl
    );
    quantityShort = check.quantityShort;
    skuMismatches = check.skuMismatches;
  }

  const siteOrigin = String(input.context.pageContext.web.absoluteUrl || "").replace(/\/$/, "");
  const detailViewUrl = `${siteOrigin}/SitePages/RequestDetail.aspx?requestId=${input.requestId}`;

  const payload = {
    requestId: input.requestId,
    actorEmail: input.actorEmail,
    position: input.position,
    action: input.action,
    comments: input.comments || "",
    quantityShort,
    skuMismatches,
    transactionId: input.transactionId,
    detailViewUrl,
    totalQuantity: input.totalQuantity ?? 0,
    totalTagPrice: input.totalTagPrice ?? 0,
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
    throw new Error(text || `Flow call failed (${res.status})`);
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