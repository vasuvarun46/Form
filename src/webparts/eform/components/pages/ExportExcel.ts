import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface ExportRequestRow {
  id: string;
  spId?: number;
  status: string;
  requestor: string;
  approver?: string;
  approverNames?: string;
  multipleApproverData?: unknown;
  purpose?: string;
  purposeDetail?: string;
  PurposeDetails?: string;
  costCenter?: string;
  department?: string;
  submitted?: string;
  approved?: string;
  approvedDates?: string;
  expected?: string;
  items?: ExportLineItem[];
  qty?: number;
  price?: number;
}

export interface ExportLineItem {
  sku: string;
  productCode?: string;
  color?: string;
  size?: string;
  qty: number;
  unitPrice: number;
  lineTotal?: number;
}

type InventoryRequestItemForExport = {
  RequestIDId?: number;
  SKUCode?: string;
  ProductDescription?: string;
  Color?: string;
  Size?: string;
  Quantity?: number;
  UnitPrice?: number;
  LineTotal?: number;
  LineNumber?: number;
  RowStatus?: string;
};

export async function fetchExportItemsByRequestIds(
  context: WebPartContext,
  requestIds: number[]
): Promise<Record<number, ExportLineItem[]>> {
  const out: Record<number, ExportLineItem[]> = {};
  const uniqueIds = Array.from(
    new Set(
      requestIds
        .map((id) => Number(id))
        .filter((id) => Number.isFinite(id) && id > 0)
    )
  );
  if (uniqueIds.length === 0) return out;

  const siteUrl = context.pageContext.web.absoluteUrl;
  const chunkSize = 25;

  for (let i = 0; i < uniqueIds.length; i += chunkSize) {
    const chunk = uniqueIds.slice(i, i + chunkSize);
    const requestFilter = chunk.map((id) => `RequestIDId eq ${id}`).join(" or ");

    let nextUrl: string | undefined =
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items` +
      `?$top=5000` +
      `&$select=RequestIDId,SKUCode,ProductDescription,Color,Size,Quantity,UnitPrice,LineTotal,LineNumber,Id,RowStatus` +
      `&$filter=(${requestFilter})` +
      `&$orderby=RequestIDId asc,LineNumber asc,Id asc`;

    while (nextUrl) {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        nextUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "odata-version": "",
          },
        }
      );

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Failed to load request line items (${res.status}): ${text}`);
      }

      const json = (await res.json()) as {
        value?: InventoryRequestItemForExport[];
        "@odata.nextLink"?: string;
        "odata.nextLink"?: string;
      };
      const batch = Array.isArray(json.value) ? json.value : [];

      for (const item of batch) {
        const requestId = Number(item?.RequestIDId || 0);
        if (!requestId) continue;
        const rowStatus = String(item?.RowStatus || "").trim().toLowerCase();
        if (rowStatus === "inactive") continue;

        const qty = Number(item?.Quantity || 0);
        const unitPrice = Number(item?.UnitPrice || 0);
        const lineTotal = Number(item?.LineTotal || 0) || unitPrice * qty;

        if (!out[requestId]) out[requestId] = [];
        out[requestId].push({
          sku: String(item?.SKUCode || ""),
          productCode: String(item?.ProductDescription || ""),
          color: String(item?.Color || ""),
          size: String(item?.Size || ""),
          qty,
          unitPrice,
          lineTotal,
        });
      }

      nextUrl =
        json["@odata.nextLink"] ||
        json["odata.nextLink"] ||
        undefined;
    }
  }

  return out;
}

/**
 * Maps a raw WorkflowStatus value to the label shown across the app's list
 * views (e.g. "In Approval"/"Pending" display as "Submitted").
 */
function toDisplayStatus(status: string): string {
  const key = String(status || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");

  const isInApproval =
    key === "pending" || key === "in approval" || key === "in-approval";

  if (key === "submitted" || isInApproval) return "Submitted";
  if (key === "approved") return "Approved";
  if (key === "rejected") return "Rejected";
  if (key === "draft") return "Draft";
  return status;
}

/**
 * Returns the names of approvers who have actually acted on the request
 * (i.e. their stage has a recorded action date), comma-separated. An
 * approver assigned to a later stage who hasn't approved/rejected yet is
 * left out — e.g. if only C1 has approved, only C1's name is returned.
 */
function getApprovedApproverNames(raw: unknown): string {
  if (!raw) return "";
  let parsed: unknown = raw;
  if (typeof raw === "string") {
    const s = raw.trim();
    if (!s) return "";
    try {
      parsed = JSON.parse(s);
    } catch {
      return "";
    }
  }
  if (!parsed || typeof parsed !== "object") return "";

  const stages = Array.isArray((parsed as { stages?: unknown[] }).stages)
    ? ((parsed as { stages: Array<Record<string, unknown>> }).stages)
    : [];

  const names = stages
    .filter(
      (s) => !!String(s?.Date ?? s?.date ?? s?.ActionDate ?? s?.actionDate ?? "").trim()
    )
    .map((s) =>
      String(s?.approverName ?? s?.ApproverName ?? s?.UserName ?? "").trim()
    )
    .filter(Boolean);

  return Array.from(new Set(names)).join(", ");
}

/**
 * Approvals land on ApprovedDates one at a time, in order (C1's date is
 * appended when C1 approves, C2's when C2 approves, etc.) — this already
 * works reliably across the app. ApproverNames lists the assigned approvers
 * in the same stage order. So "how many have approved so far" (the count of
 * dates) tells us how many *names*, from the front of that list, are safe
 * to reveal — e.g. once C1 approves (1 date recorded), show just C1; once
 * C2 also approves (2 dates), show "C1, C2".
 */
function getApprovedCountFromDates(approvedDates?: string): number {
  return String(approvedDates || "")
    .split(",")
    .map((d) => d.trim())
    .filter(Boolean).length;
}

function splitAssignedNames(names?: string): string[] {
  return String(names || "")
    .split(",")
    .map((n) => n.trim())
    .filter(Boolean);
}

/**
 * Resolves the Approver Name cell for a request row. Tries, in order:
 * 1. The per-stage "who has acted so far" breakdown parsed from the stage
 *    JSON (most precise, but depends on the flow's JSON schema).
 * 2. The first N assigned-approver names, where N is the number of
 *    approvals recorded in ApprovedDates so far — this is what makes C1's
 *    name appear as soon as C1 approves (even while the request is still
 *    "In Approval"/shown as "Submitted"), without waiting for the whole
 *    chain to finish.
 * 3. If the request has already reached a terminal status (Approved/
 *    Rejected) but neither of the above resolved anything, fall back to
 *    the full assigned-approver text — safe at that point since the
 *    workflow can't be terminal without its approver(s) having acted.
 * Anything still pending with zero recorded approvals stays blank, so an
 * unactioned approver's name never leaks through.
 */
function getExcelApproverNames(row: ExportRequestRow): string {
  const statusKey = String(row.status || "").trim().toLowerCase();
  const isTerminal = statusKey === "approved" || statusKey === "rejected";
  const assignedNames = splitAssignedNames(row.approverNames || row.approver);
  const approvedCount = getApprovedCountFromDates(row.approvedDates);

  // ApprovedDates is the authoritative signal — PA appends a date only when
  // an approver actually acts. Gate everything behind it so no name appears
  // until at least one approval is recorded.
  if (approvedCount > 0) {
    // Stage JSON is most precise: only stages with a recorded action date are included.
    const fromStages = getApprovedApproverNames(row.multipleApproverData);
    if (fromStages) return fromStages;

    // Fallback: slice the assigned-name list by the number of recorded approvals.
    if (assignedNames.length > 0) {
      return assignedNames.slice(0, Math.min(approvedCount, assignedNames.length)).join(", ");
    }
  }

  // Terminal fallback: Approved/Rejected but no dates recorded (edge case).
  if (isTerminal && assignedNames.length > 0) {
    return assignedNames.join(", ");
  }

  return "";
}

/**
 * Exports requests to Excel. Expands one row per SKU line item.
 * Transaction columns are repeated on every item row for the same request.
 * Uses aoa_to_sheet so both header groups can share the "Total Tag Price" label.
 */
export async function exportToExcel(rows: ExportRequestRow[]): Promise<void> {
  if (!rows || rows.length === 0) return;

  // Sort request rows by spId descending (newest first)
  const sortedRows = [...rows].sort((a, b) => (b.spId || 0) - (a.spId || 0));

  const XLSX = await import(/* webpackChunkName: "xlsx" */ "xlsx");

  const headers: string[] = [
    "Transaction ID",
    "Status",
    "Requestor Name",
    "Approver Name",
    "Purpose",
    "Purpose Detail",
    "Cost Center",
    "Submitted Date",
    "Approved Date",
    "Expected Delivery Date",
    "SKU Code",
    "Product Description",
    "Quantity",
    "Tag Price per Unit",
    "Tag Price Total",
  ];

  const data: (string | number)[][] = [headers];

  for (const r of sortedRows) {
    const txnBase: (string | number)[] = [
      r.id,
      toDisplayStatus(r.status),
      r.requestor,
      getExcelApproverNames(r),
      r.purpose || "",
      r.purposeDetail || r.PurposeDetails || "",
      r.costCenter || "",
      r.submitted || "",
      r.approvedDates || r.approved || "",
      r.expected || "",
    ];

    const items = r.items && r.items.length > 0 ? r.items : null;

    if (items) {
      for (const item of items) {
        data.push([
          ...txnBase,
          item.sku || "",
          item.productCode || "",
          Number(item.qty || 0),
          Number(item.unitPrice || 0),
          Number(item.lineTotal || 0),
        ]);
      }
    } else {
      data.push([...txnBase, "", "", "", "", ""]);
    }
  }

  const worksheet = XLSX.utils.aoa_to_sheet(data);

  worksheet["!cols"] = [
    { wch: 18 },  // Transaction ID
    { wch: 12 },  // Status
    { wch: 22 },  // Requestor Name
    { wch: 22 },  // Approver Name
    { wch: 20 },  // Purpose
    { wch: 28 },  // Purpose Detail
    { wch: 16 },  // Cost Center
    { wch: 14 },  // Submitted Date
    { wch: 22 },  // Approved Date
    { wch: 20 },  // Expected Delivery Date
    { wch: 18 },  // SKU Code
    { wch: 28 },  // Product Description
    { wch: 12 },  // Quantity
    { wch: 18 },  // Tag Price per Unit
    { wch: 18 },  // Tag Price
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Requests");
  XLSX.writeFile(workbook, `Requests_${Date.now()}.xlsx`);
}
