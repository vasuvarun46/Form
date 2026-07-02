import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface ExportRequestRow {
  id: string;
  spId?: number;
  status: string;
  requestor: string;
  approver?: string;
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
      r.status,
      r.requestor,
      r.approver || "",
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
