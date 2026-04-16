const SUBMITTED_TXN_ID_PATTERN = /^II\d{6}\/\d{6}$/i;
const LEGACY_TXN_ID_PATTERN = /^II\d{6}\/\d{4}$/i;

const toDraftDigits = (value?: number): string => {
  const safe = Number.isFinite(value) ? Math.max(0, Math.trunc(Number(value))) : 0;
  return String(safe).padStart(4, "0").slice(-4);
};

export const resolveDisplayTransactionId = (params: {
  workflowStatus?: string;
  title?: string;
  requestNumber?: number;
  spId?: number;
}): string => {
  const status = String(params.workflowStatus || "").trim().toLowerCase();
  const title = String(params.title || "").trim();

  if (status === "draft") {
    if (/^\d{4}$/.test(title)) {
      return title;
    }

    return toDraftDigits(params.requestNumber ?? params.spId);
  }
 if (status === "pending") {
    return "";
  }
  if (SUBMITTED_TXN_ID_PATTERN.test(title) || LEGACY_TXN_ID_PATTERN.test(title)) {
    return title;
  }

  if (title && !/^draft$/i.test(title) && !/^temp$/i.test(title)) {
    return title;
  }

  if (params.requestNumber !== undefined && params.requestNumber !== null) {
    return String(params.requestNumber);
  }

  return String(params.spId || "");
};
