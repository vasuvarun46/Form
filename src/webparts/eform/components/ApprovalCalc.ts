/* =========================================================
   ApprovalCalc.ts
   Pure, framework-free approval maths + a concurrency helper.
   No SPFx imports, so it is unit-testable in isolation.
   Consumed by ApprovalProcessor.ts.
========================================================= */

export interface SkuMismatch {
  itemId: number;
  sku: string;
  tagPricePerUnit: number;
  requestedQty: number;
  totalTagPrice: number;
  availableQty: number;
  actualTotalTagPrice: number;
  differenceQty: number;
  // Actual live warehouse stock (Warehouse_SOH_Qty). Distinct from availableQty
  // (which drives ApprovalCount). -1 = unverifiable (no/failed stock call).
  liveStockQty: number;
}

/** Approved quantity for a line: unverifiable stock (-1) keeps the requested qty;
 *  otherwise the available qty, never negative. */
export function approvedQtyFor(requestedQty: number, availableQty: number): number {
  const requested = Number.isFinite(Number(requestedQty)) ? Number(requestedQty) : 0;
  const avail = Number(availableQty);
  return !Number.isFinite(avail) || avail === -1 ? requested : Math.max(0, avail);
}

/** Header totals, SUMMED from each line's approved quantity (never derived by
 *  subtracting from a cached total, which could go negative). Clamped at 0. */
export function computeHeaderTotals(
  items: Array<Pick<SkuMismatch, "requestedQty" | "availableQty" | "tagPricePerUnit">>
): { totalQty: number; totalAmount: number } {
  let qty = 0;
  let amount = 0;
  for (const m of items) {
    const adj = approvedQtyFor(m.requestedQty, m.availableQty);
    const unit = Number.isFinite(Number(m.tagPricePerUnit)) ? Number(m.tagPricePerUnit) : 0;
    qty += adj;
    amount += adj * unit;
  }
  return { totalQty: Math.max(0, qty), totalAmount: Math.max(0, amount) };
}

/** The columns written to one InventoryRequestItems row. LineTotal only on a
 *  genuine approval; StockAvailability only when the live figure is known (>= 0). */
export function computeItemFields(
  m: Pick<SkuMismatch, "requestedQty" | "availableQty" | "tagPricePerUnit" | "liveStockQty">,
  includeLineTotal: boolean
): Record<string, number> {
  const requested = Number.isFinite(Number(m.requestedQty)) ? Number(m.requestedQty) : 0;
  const approved = approvedQtyFor(m.requestedQty, m.availableQty);
  const unit = Number.isFinite(Number(m.tagPricePerUnit)) ? Number(m.tagPricePerUnit) : 0;

  const fields: Record<string, number> = {
    RaiseCount: requested,
    ApprovalCount: approved,
  };
  if (includeLineTotal) {
    fields.LineTotal = Math.max(0, approved * unit);
  }
  const live = Number(m.liveStockQty);
  if (Number.isFinite(live) && live >= 0) {
    fields.StockAvailability = live;
  }
  return fields;
}

/** Run `fn` over `items` with at most `limit` in flight at once. Results keep
 *  input order. No unbounded Promise.all. */
export async function mapWithConcurrency<T, R>(
  items: T[],
  limit: number,
  fn: (item: T, index: number) => Promise<R>
): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  const workerCount = Math.max(1, Math.min(limit, items.length));

  const worker = async (): Promise<void> => {
    for (;;) {
      const i = next;
      next += 1;
      if (i >= items.length) return;
      results[i] = await fn(items[i], i);
    }
  };

  await Promise.all(Array.from({ length: workerCount }, () => worker()));
  return results;
}
