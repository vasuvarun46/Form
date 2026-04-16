import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface ExportRequestRow {
  id: string;
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
  expected?: string;
  // Line-item fields (new)
  items?: ExportLineItem[];
  // Legacy summary fields (backward compat)
  qty?: number;
  price?: number;
}

// type ExportRequestRowWithDepartmentFallback = ExportRequestRow & {
//   departmentTitle?: string;
//   Department?: { Title?: string };
// };

export interface ExportLineItem {
  sku: string;
  productCode?: string;
  productDescription?: string;
  qty: number;
  unitPrice: number;
}

type InventoryRequestItemForExport = {
  RequestIDId?: number;
  SKUCode?: string;
  ProductDescription?: string;
  Quantity?: number;
  UnitPrice?: number;
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
      `&$select=RequestIDId,SKUCode,ProductDescription,Quantity,UnitPrice,LineNumber,Id,RowStatus` +
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

        if (!out[requestId]) out[requestId] = [];
        out[requestId].push({
          sku: String(item?.SKUCode || "-"),
          productDescription: String(item?.ProductDescription || "-"),
          qty: Number(item?.Quantity || 0),
          unitPrice: Number(item?.UnitPrice || 0),
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
 * Export requests to Excel with 15 columns per FR FS-II-040.
 * If items[] is provided, expands to one row per SKU.
 * Otherwise falls back to summary row.
 */
export async function exportToExcel(rows: ExportRequestRow[]): Promise<void> {
  if (!rows || rows.length === 0) return;

  const XLSX = await import(/* webpackChunkName: "xlsx" */ "xlsx");

  const formatted: Record<string, string | number>[] = [];

  for (const r of rows) {
    // ✅ One row per request — no line-item expansion
    formatted.push({
      "Transaction ID":           r.id,
      "Status":                   r.status,
      "Requestor Name":           r.requestor,
      "Approver Name":            r.approver || " ",
      "Purpose":                  r.purpose || " ",
      "Purpose Detail":           r.purposeDetail || r.PurposeDetails || "-",
      "Cost Center":              r.costCenter || " ",
      "Submitted Date":           r.submitted || " ",
      "Approved Date":            r.approved || " ",
      "Expected Delivery Date":   r.expected || " ",
      "Total Quantity":           Number(r.qty || 0),
      "Total Tag Price":          Number(r.price || 0),
    });
  }

  const worksheet = XLSX.utils.json_to_sheet(formatted);

  // ✅ Column widths matching 12 columns
  worksheet["!cols"] = [
    { wch: 18 },  // Transaction ID
    { wch: 12 },  // Status
    { wch: 22 },  // Requestor Name
    { wch: 22 },  // Approver Name
    { wch: 20 },  // Purpose
    { wch: 28 },  // Purpose Detail
    { wch: 16 },  // Cost Center
    { wch: 14 },  // Submitted Date
    { wch: 14 },  // Approved Date
    { wch: 20 },  // Expected Delivery Date
    { wch: 14 },  // Total Quantity
    { wch: 16 },  // Total Tag Price
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Requests");
  XLSX.writeFile(workbook, `Requests_${Date.now()}.xlsx`);
}
