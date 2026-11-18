import {
  approvedQtyFor,
  computeHeaderTotals,
  computeItemFields,
  mapWithConcurrency,
  SkuMismatch,
} from "../webparts/eform/components/ApprovalCalc";

const line = (over: Partial<SkuMismatch>): SkuMismatch => ({
  itemId: 1,
  sku: "SKU1",
  tagPricePerUnit: 10,
  requestedQty: 5,
  totalTagPrice: 50,
  availableQty: 5,
  actualTotalTagPrice: 50,
  differenceQty: 0,
  liveStockQty: 5,
  ...over,
});

describe("approvedQtyFor", () => {
  test("in stock → approved equals requested", () => {
    expect(approvedQtyFor(5, 5)).toBe(5);
  });
  test("partial stock → approved equals available", () => {
    expect(approvedQtyFor(10, 3)).toBe(3);
  });
  test("out of stock → approved is 0", () => {
    expect(approvedQtyFor(10, 0)).toBe(0);
  });
  test("unverifiable stock (-1) → approved keeps requested", () => {
    expect(approvedQtyFor(7, -1)).toBe(7);
  });
  test("negative (non -1) stock is clamped to 0", () => {
    expect(approvedQtyFor(7, -4)).toBe(0);
  });
});

describe("computeItemFields", () => {
  test("RaiseCount is always the requested qty", () => {
    const f = computeItemFields(line({ requestedQty: 9, availableQty: 2 }), true);
    expect(f.RaiseCount).toBe(9);
  });

  test("ApprovalCount + LineTotal for partial stock when included", () => {
    const f = computeItemFields(
      line({ requestedQty: 10, availableQty: 4, tagPricePerUnit: 25 }),
      true
    );
    expect(f.ApprovalCount).toBe(4);
    expect(f.LineTotal).toBe(100); // 4 × 25
  });

  test("out of stock approves at qty 0 with LineTotal 0", () => {
    const f = computeItemFields(
      line({ requestedQty: 6, availableQty: 0, liveStockQty: 0, tagPricePerUnit: 25 }),
      true
    );
    expect(f.ApprovalCount).toBe(0);
    expect(f.LineTotal).toBe(0);
    expect(f.StockAvailability).toBe(0); // verified zero IS written
  });

  test("LineTotal omitted when includeLineTotal = false", () => {
    const f = computeItemFields(line({}), false);
    expect(f.LineTotal).toBeUndefined();
  });

  test("StockAvailability omitted when live stock is unverifiable (-1)", () => {
    const f = computeItemFields(
      line({ availableQty: -1, liveStockQty: -1, requestedQty: 5 }),
      true
    );
    expect("StockAvailability" in f).toBe(false);
    expect(f.ApprovalCount).toBe(5); // unverifiable keeps requested
  });
});

describe("computeHeaderTotals", () => {
  test("sums approved quantities and amounts across lines", () => {
    const totals = computeHeaderTotals([
      line({ requestedQty: 5, availableQty: 5, tagPricePerUnit: 10 }), // 5 × 10
      line({ requestedQty: 10, availableQty: 4, tagPricePerUnit: 20 }), // 4 × 20
    ]);
    expect(totals.totalQty).toBe(9); // 5 + 4
    expect(totals.totalAmount).toBe(130); // 50 + 80
  });

  test("out-of-stock line contributes 0 to both totals, never negative", () => {
    const totals = computeHeaderTotals([
      line({ requestedQty: 3, availableQty: 3, tagPricePerUnit: 10 }), // 30
      line({ requestedQty: 8, availableQty: 0, tagPricePerUnit: 10 }), // 0
    ]);
    expect(totals.totalQty).toBe(3);
    expect(totals.totalAmount).toBe(30);
  });

  test("unverifiable line keeps the requested qty in the totals", () => {
    const totals = computeHeaderTotals([
      line({ requestedQty: 4, availableQty: -1, tagPricePerUnit: 15 }), // keeps 4
    ]);
    expect(totals.totalQty).toBe(4);
    expect(totals.totalAmount).toBe(60);
  });

  test("empty set totals to 0", () => {
    expect(computeHeaderTotals([])).toEqual({ totalQty: 0, totalAmount: 0 });
  });
});

describe("mapWithConcurrency", () => {
  test("preserves input order and processes every item", async () => {
    const out = await mapWithConcurrency([1, 2, 3, 4, 5], 2, async (n) => n * 2);
    expect(out).toEqual([2, 4, 6, 8, 10]);
  });

  test("never exceeds the concurrency limit of in-flight tasks", async () => {
    let inFlight = 0;
    let peak = 0;
    await mapWithConcurrency(Array.from({ length: 12 }, (_, i) => i), 5, async () => {
      inFlight += 1;
      peak = Math.max(peak, inFlight);
      await new Promise((resolve) => setTimeout(resolve, 5));
      inFlight -= 1;
      return 0;
    });
    expect(peak).toBeLessThanOrEqual(5);
  });
});
