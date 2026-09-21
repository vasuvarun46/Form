export type FilterNormalizer = (value: string) => string;

export interface FilterSets {
  status: Set<string>;
  purpose: Set<string>;
  costCenter: Set<string>;
}

export interface FilterableRequestRow {
  status: string;
  purpose: string;
  costCenter: string;
}

const collapseSpaces = (value: string): string =>
  String(value || "").replace(/\s+/g, " ").trim();

export const normalizeFilterText: FilterNormalizer = (value: string): string =>
  collapseSpaces(value).toLowerCase();

export const canonicalizeStatusFilterKey: FilterNormalizer = (
  value: string
): string => normalizeFilterText(value);

export const dedupeLabels = (
  values: string[],
  keySelector: FilterNormalizer = normalizeFilterText
): string[] => {
  const out: string[] = [];
  const seen = new Set<string>();

  for (const value of values) {
    const label = collapseSpaces(value);
    if (!label) continue;

    const key = keySelector(label);
    if (seen.has(key)) continue;

    seen.add(key);
    out.push(label);
  }

  return out;
};

export const buildFilterSet = (
  values: string[],
  normalizer: FilterNormalizer = normalizeFilterText
): Set<string> => {
  const out = new Set<string>();

  for (const value of values) {
    const normalized = normalizer(value);
    if (!normalized) continue;
    out.add(normalized);
  }

  return out;
};

export const compileFilterSets = (args: {
  statusFilter: string[];
  purposeFilter: string[];
  costCenterFilter: string[];
}): FilterSets => ({
  status: buildFilterSet(args.statusFilter, canonicalizeStatusFilterKey),
  purpose: buildFilterSet(args.purposeFilter, normalizeFilterText),
  costCenter: buildFilterSet(args.costCenterFilter, normalizeFilterText),
});

export const matchesMultiSelect = (
  selectedSet: Set<string>,
  rowValue: string,
  normalizer: FilterNormalizer = normalizeFilterText
): boolean => {
  if (selectedSet.size === 0) return true;
  return selectedSet.has(normalizer(rowValue));
};

export const matchesCompiledFilters = (
  row: FilterableRequestRow,
  filters: FilterSets
): boolean =>
  matchesMultiSelect(filters.status, row.status, canonicalizeStatusFilterKey) &&
  matchesMultiSelect(filters.purpose, row.purpose, normalizeFilterText) &&
  matchesMultiSelect(filters.costCenter, row.costCenter, normalizeFilterText);
