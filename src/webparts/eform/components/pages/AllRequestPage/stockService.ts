import { HttpClient } from "@microsoft/sp-http";

/* ===== Types (copied safely) ===== */

export interface AwsStockFlow {
  Barcode?: string;
  Warehouse_SOH_Qty?: number;
  ProductCode?: string;
  ColourDesc?: string;
  SizeCode?: string;
  Original_sp?: number;
}

/* ===== AWS fetch ===== */

export async function fetchStockFromAws(
  httpClient: HttpClient,
  sku: string,
  flowUrl: string
): Promise<{ Warehouse_SOH_Qty: number } | undefined> {
  if (!sku) return undefined;
  if (!flowUrl) {
    console.error("Stock check skipped: flowUrl not configured.");
    return undefined;
  }

  try {
    const response = await httpClient.post(
      flowUrl,
      HttpClient.configurations.v1,
      {
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sku }),
      }
    );

    if (!response.ok) throw new Error(response.statusText);

    const json = await response.json();

    const qty = Number(
      json?.Warehouse_SOH_Qty ??
      json?.body?.Warehouse_SOH_Qty ??
      0
    );

    return { Warehouse_SOH_Qty: qty };
  } catch (e) {
    console.error("Stock check failed:", e);
    return undefined;
  }
}




/* ===== Normalizer ===== */

export function normalizeAwsFlow(json: unknown): AwsStockFlow | undefined {
  const o = json as Record<string, unknown>;
  if (!o) return undefined;

 return {
  Barcode: String(o.Barcode || ""),
  Warehouse_SOH_Qty: Number(o.Warehouse_SOH_Qty || 0),
  ProductCode: String(o.ProductCode || ""),
  ColourDesc: String(o.ColourDesc || ""),
  SizeCode: String(o.SizeCode || ""),
  Original_sp: Number(o.Original_sp || 0),
};
}
