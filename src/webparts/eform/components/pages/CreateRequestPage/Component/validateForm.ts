// ═══════════════════════════════════════════════════════════════════════════
// FIX: P2-001 (CRITICAL) + P2-010 (MEDIUM)
// FILE: src/webparts/eform/components/pages/CreateRequestPage/Component/validateForm.ts
// ═══════════════════════════════════════════════════════════════════════════
// REPLACE the entire file with this version.
// Changes marked with // 🔧 FIX


import { IItemRow, ItemRowErrors, IUsageState, UsageFieldErrors, ValidationResult } from "../createRequestTypes";

export default function validateForm(
  usage: IUsageState,
  items: IItemRow[]
): ValidationResult {

  const fe: UsageFieldErrors = {};
  const ie: ItemRowErrors = {};

  const requiredText = (val: string): boolean =>
    String(val || "").trim().length > 0;

  /* =========================================================
     🔴 HEADER VALIDATION
     ========================================================= */

  if (!requiredText(usage.purpose))
    fe.purpose = "Purpose is required.";

  if (!requiredText(usage.costCenter))
    fe.costCenter = "Cost Center is required.";

  if (!requiredText(usage.expectedDelivery))
    fe.expectedDelivery = "Expected Delivery Date is required.";

  // ✅ Safer Past Date Validation
  if (requiredText(usage.expectedDelivery)) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const selectedDate = new Date(usage.expectedDelivery);
    selectedDate.setHours(0, 0, 0, 0);

    if (selectedDate < today) {
      fe.expectedDelivery = "Expected Delivery Date cannot be in the past.";
    }
  }

  if (!requiredText(usage.placeOfDelivery))
    fe.placeOfDelivery = "Place of Delivery is required.";

  if (!requiredText(usage.purposeDetail))
    fe.purposeDetail = "Purpose Detail is required.";

  const hasFieldErrors = Object.keys(fe).length > 0;

  /* =========================================================
     🔴 FIELD ERROR
     ========================================================= */

  if (hasFieldErrors) {
    return {
      ok: false,
      submitMessage: "Please complete all required fields in Usage Information.",
      fieldErrors: fe,
      itemErrors: {},
    };
  }

  /* =========================================================
     🔴 NO ITEMS CHECK
     ========================================================= */

  if (!items.length) {
    return {
      ok: false,
      submitMessage: "Please add at least 1 product before saving draft or submitting.",
      fieldErrors: {},
      itemErrors: {},
    };
  }

  /* =========================================================
     🔴 ITEM VALIDATION
     ========================================================= */

  items.forEach((it, idx) => {
    const qty = Number(it.qty || 0);
    const price = Number(it.unitPrice || 0);

    if (!String(it.sku || "").trim())
      ie[idx] = "SKU Code is missing.";

    else if (!String(it.productCode || "").trim())
      ie[idx] = "Product Code is missing.";

    else if (qty <= 0)
      ie[idx] = "Quantity must be greater than 0.";

    else if (qty > 999)
      ie[idx] = "Quantity must be between 1 and 999.";

    else if (price <= 0)
      ie[idx] = "Tag Price / Unit must be greater than 0.";
  });

  const hasItemErrors = Object.keys(ie).length > 0;

  /* =========================================================
     🔴 ITEM ERROR
     ========================================================= */

  if (hasItemErrors) {
    return {
      ok: false,
      submitMessage: "Please fix the highlighted product row errors before submitting.",
      fieldErrors: {},
      itemErrors: ie,
    };
  }

  /* =========================================================
     🔴 TOTAL AMOUNT LIMIT CHECK (≤ 500,000)
     ========================================================= */

  const totalAmount = items.reduce((sum, it) => {
    return sum + (Number(it.qty || 0) * Number(it.unitPrice || 0));
  }, 0);

  if (totalAmount > 500000) {
    return {
      ok: false,
      showAsPopup: true,
      submitMessage: `Total request amount (${totalAmount.toLocaleString()} THB) exceeds the 500,000 THB limit.`,
      fieldErrors: {},
      itemErrors: {},
    };
  }

  /* =========================================================
     ✅ ALL GOOD
     ========================================================= */

  return {
    ok: true,
    submitMessage: "",
    fieldErrors: {},
    itemErrors: {},
  };
}
