export const normalizeWorkflowKey = (t?: string): string => {
  const s = String(t || "").replace(/\s+/g, " ").trim();
  const removedParen = s.replace(/\s*-\s*\(\s*.*?\s*\)\s*$/i, "").trim();
  const removedDashSuffix = removedParen.replace(/\s*-\s*[^-()]+$/i, "").trim();
  return removedDashSuffix || removedParen || s;
};

export const normalizeTitleKey = (t?: string): string =>
  String(t || "")
    .replace(/\s+/g, " ")
    .trim();
