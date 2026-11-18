import * as React from "react";
import ItemCategoryHeader from "../../Common/PageHeader/ItemCategoryHeader";
import { IconChevronDown } from "../../ui/icons";
import ProductFilterCard from "./ProductFilterCard";
import ProductMasterTable, { PRODUCT_COLUMNS } from "./ProductMasterTable";
import Loading from "../../ui/Loading";
import {
  CatalogApiError,
  exportAllMatching,
  searchCatalog,
  SEARCH_PAGE_SIZE,
} from "./itemCategoryApi";
import {
  EMPTY_PRODUCT_FILTERS,
  PRODUCT_FILTER_FIELDS,
  ProductFilterKey,
  ProductFilters,
  ProductMasterRow,
  ProductSortKey,
} from "./itemCategoryTypes";

type SortDir = "asc" | "desc";

const PAGE_SIZE_OPTIONS = [20, 50, 100];

// How long to wait after the user stops typing before calling the API.
const SEARCH_DEBOUNCE_MS = 400;

// Built once. Calling localeCompare(..., "th") per comparison builds a new
// collator on every pair, which is costly when sorting runs on every keystroke.
const TH_COLLATOR = new Intl.Collator("th");

const pageBg = "bg-[#f3f3f3]";
const card = "bg-white rounded-xl border border-[#e5e5e5] shadow-sm";

const pagerButton =
  "h-7 px-2.5 rounded-md border text-[12px] font-medium bg-white border-gray-300 " +
  "text-[#2172DA] hover:bg-[#2172DA] hover:text-white disabled:cursor-not-allowed " +
  "disabled:border-[#D9D9D9] disabled:bg-[#D9D9D9] disabled:text-gray-500";

export default function ItemCategoryPage(): JSX.Element {
  const [filters, setFilters] = React.useState<ProductFilters>({
    ...EMPTY_PRODUCT_FILTERS,
  });
  const [sortKey, setSortKey] = React.useState<ProductSortKey>("sku");
  const [sortDir, setSortDir] = React.useState<SortDir>("asc");
  const [pageSize, setPageSize] = React.useState<number>(20);
  const [pageIndex, setPageIndex] = React.useState<number>(0);
  const [exportError, setExportError] = React.useState<string>("");

  // ── Live search state ──
  // `rows` is the buffer of results fetched so far for the current filters.
  const [rows, setRows] = React.useState<ProductMasterRow[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true); // first load / re-search
  const [searching, setSearching] = React.useState<boolean>(false); // paging more in
  const [loadError, setLoadError] = React.useState<string>("");

  // Keyset paging: the cursor for the page after the buffer, and whether one exists.
  const cursorRef = React.useRef<string | null>(null);
  const [hasNext, setHasNext] = React.useState<boolean>(false);

  // Monotonic id so only the latest request's result may touch the grid. A newer
  // search (or Reset) bumps this; stale responses that return late are dropped.
  const reqIdRef = React.useRef<number>(0);
  const debounceRef = React.useRef<number | undefined>(undefined);

  // ── Run a fresh search (page 1) for the given filters ──
  const runSearch = React.useCallback(
    (nextFilters: ProductFilters, initial: boolean): void => {
      const reqId = (reqIdRef.current += 1);
      cursorRef.current = null;
      if (initial) setLoading(true);
      else setSearching(true);
      setLoadError("");

      searchCatalog(nextFilters, null, SEARCH_PAGE_SIZE)
        .then((result) => {
          if (reqId !== reqIdRef.current) return; // a newer search superseded us
          setRows(result.rows);
          cursorRef.current = result.nextCursor;
          setHasNext(result.hasNext);
          setPageIndex(0);
          setLoading(false);
          setSearching(false);
        })
        .catch((e: unknown) => {
          if (reqId !== reqIdRef.current) return;
          setLoadError(
            e instanceof Error ? e.message : "Failed to load products."
          );
          setRows([]);
          setHasNext(false);
          setLoading(false);
          setSearching(false);
        });
    },
    []
  );

  // ── Fetch the next cursor page and append it to the buffer ──
  const loadMore = React.useCallback((): Promise<void> => {
    const cursor = cursorRef.current;
    if (!cursor) return Promise.resolve();
    const reqId = reqIdRef.current; // tie to the current search; a new one cancels this
    setSearching(true);

    return searchCatalog(filters, cursor, SEARCH_PAGE_SIZE)
      .then((result) => {
        if (reqId !== reqIdRef.current) return;
        setRows((prev) => {
          const seen = new Set(prev.map((r) => r.id));
          return prev.concat(result.rows.filter((r) => !seen.has(r.id)));
        });
        cursorRef.current = result.nextCursor;
        setHasNext(result.hasNext);
        setSearching(false);
      })
      .catch((e: unknown) => {
        if (reqId !== reqIdRef.current) return;
        // A stale cursor (filters changed underneath) → start over from page 1.
        if (e instanceof CatalogApiError && e.code === "invalid_cursor") {
          runSearch(filters, false);
          return;
        }
        setLoadError(
          e instanceof Error ? e.message : "Failed to load more products."
        );
        setSearching(false);
      });
  }, [filters, runSearch]);

  // ── Initial load (no filters) ──
  React.useEffect(() => {
    runSearch({ ...EMPTY_PRODUCT_FILTERS }, true);
    return () => {
      // Invalidate any in-flight request on unmount.
      reqIdRef.current += 1;
      if (debounceRef.current !== undefined) {
        window.clearTimeout(debounceRef.current);
      }
    };
    // runSearch is stable (empty deps); run once on mount.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Debounced search whenever the filters change ──
  const scheduleSearch = React.useCallback(
    (nextFilters: ProductFilters): void => {
      if (debounceRef.current !== undefined) {
        window.clearTimeout(debounceRef.current);
      }
      debounceRef.current = window.setTimeout(() => {
        runSearch(nextFilters, false);
      }, SEARCH_DEBOUNCE_MS);
    },
    [runSearch]
  );

  // ── Sorting (over the loaded buffer) ──
  const sortedRows = React.useMemo<ProductMasterRow[]>(() => {
    const dir = sortDir === "asc" ? 1 : -1;
    const isNumeric = PRODUCT_COLUMNS.some(
      (c) => c.key === sortKey && c.numeric === true
    );

    return [...rows].sort((a, b) => {
      if (isNumeric) {
        return (Number(a[sortKey]) - Number(b[sortKey])) * dir;
      }
      return (
        TH_COLLATOR.compare(String(a[sortKey] ?? ""), String(b[sortKey] ?? "")) *
        dir
      );
    });
  }, [rows, sortKey, sortDir]);

  // ── Paging (over the buffer; Next fetches more when we reach the end) ──
  const totalPages = Math.max(1, Math.ceil(sortedRows.length / pageSize));
  const safePageIndex = Math.min(pageIndex, totalPages - 1);
  const firstRowNumber =
    sortedRows.length === 0 ? 0 : safePageIndex * pageSize + 1;
  const pageRows = sortedRows.slice(
    safePageIndex * pageSize,
    safePageIndex * pageSize + pageSize
  );
  const lastRowNumber =
    firstRowNumber === 0 ? 0 : firstRowNumber + pageRows.length - 1;

  const atBufferEnd = safePageIndex >= totalPages - 1;
  // There is a page after this one if the buffer has one, or the API has more.
  const canGoNext = !atBufferEnd || hasNext;

  const activeFilterCount = PRODUCT_FILTER_FIELDS.filter(
    (f) => (filters[f.key] || "").trim().length > 0
  ).length;

  const statusText =
    activeFilterCount === 0
      ? "View live product data. Enter any details to search the complete product catalogue"
      : `${sortedRows.length}${hasNext ? "+" : ""} loaded for ` +
        `${activeFilterCount} ${activeFilterCount === 1 ? "filter" : "filters"}`;

  const summaryText =
    sortedRows.length === 0
      ? "Showing 0 entries"
      : `Showing ${firstRowNumber} to ${lastRowNumber} of ${sortedRows.length}` +
        `${hasNext ? "+" : ""} entries`;

  // ── Handlers ──
  const handleFilterChange = (key: ProductFilterKey, value: string): void => {
    setFilters((prev) => {
      const next = { ...prev, [key]: value };
      scheduleSearch(next);
      return next;
    });
  };

  const handleClearAll = (): void => {
    if (debounceRef.current !== undefined) {
      window.clearTimeout(debounceRef.current);
    }
    const cleared = { ...EMPTY_PRODUCT_FILTERS };
    setFilters(cleared);
    runSearch(cleared, false);
  };

  const handleSort = (key: ProductSortKey): void => {
    if (key === sortKey) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortKey(key);
      setSortDir("asc");
    }
    setPageIndex(0);
  };

  const handleNext = (): void => {
    if (atBufferEnd) {
      if (hasNext && !searching) {
        loadMore()
          .then(() => setPageIndex((p) => p + 1))
          .catch(() => {
            /* error surfaced in loadMore */
          });
      }
    } else {
      setPageIndex(safePageIndex + 1);
    }
  };

  const handleExport = async (): Promise<void> => {
    try {
      const all = await exportAllMatching(filters);
      if (all.length === 0) {
        setExportError("No matching rows to export.");
        return;
      }

      const XLSX = await import(/* webpackChunkName: "xlsx" */ "xlsx");

      const header = PRODUCT_COLUMNS.map((c) => c.header.split("<br/>").join(" "));
      const body = all.map((r) => PRODUCT_COLUMNS.map((c) => r[c.key]));

      const worksheet = XLSX.utils.aoa_to_sheet([header, ...body]);
      // Without explicit widths the long Thai product names are clipped in Excel.
      worksheet["!cols"] = PRODUCT_COLUMNS.map((c) => ({ wch: c.wch }));

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Product Master");
      XLSX.writeFile(workbook, `ProductMaster_${Date.now()}.xlsx`);
      setExportError("");
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Export failed.";
      setExportError(`Excel export failed: ${msg}`);
    }
  };

  return (
    <div className="pdf-thai flex flex-col h-full min-h-0 overflow-hidden">
      {/* ================= HEADER ================= */}
      <div className="sticky top-0 z-20 bg-white shrink-0 border-b border-gray-200">
        <ItemCategoryHeader onExport={handleExport} disabled={rows.length === 0} />
      </div>

      {/* ================= BODY (ONLY SCROLL AREA) ================= */}
      <div className={`flex-1 overflow-y-auto overflow-x-hidden ${pageBg} py-4 px-6 sm:px-6`}>
        <div className="space-y-3">
          {/* Filters */}
          <ProductFilterCard
            filters={filters}
            onChange={handleFilterChange}
            onClearAll={handleClearAll}
            statusText={statusText}
          />

          {exportError && (
            <div className="text-[12px] text-[#E60023]" role="alert">
              {exportError}
            </div>
          )}

          {loading && (
            <Loading
              title="Loading products"
              description="Fetching live product data."
            />
          )}

          {!loading && loadError && (
            <div className={`${card} p-6`}>
              <div
                className="text-[13px] font-semibold text-[#E60023]"
                role="alert"
              >
                Couldn&apos;t load products
              </div>
              <div className="mt-1 text-[12px] text-gray-600 break-words">
                {loadError}
              </div>
            </div>
          )}

          {!loading && !loadError && (
          <div className="space-y-2">
            {/* Title row - the design puts "Show N entries" on the right of this
                line, not inside the card. */}
            <div className="flex items-center justify-between gap-3 flex-wrap">
              <div className="flex items-center gap-2">
                <div className="text-[13px] font-semibold text-gray-900">Product list</div>
                {searching && (
                  <span className="text-[11px] text-gray-400">Searching…</span>
                )}
              </div>

              <div className="flex items-center gap-2 text-[12px] text-gray-700">
                <label htmlFor="pm-page-size">Show</label>

                <div className="relative">
                  <select
                    id="pm-page-size"
                    value={String(pageSize)}
                    onChange={(e) => {
                      setPageSize(Number(e.target.value));
                      setPageIndex(0);
                    }}
                    className="h-7 rounded-md bg-white border border-[#e7e7e7] pl-2 pr-7 text-[12px] text-gray-700 appearance-none focus:outline-none focus:border-[#2172DA] focus:shadow-[0_0_0_3px_rgba(33,114,218,0.15)]"
                  >
                    {PAGE_SIZE_OPTIONS.map((n) => (
                      <option key={n} value={String(n)}>
                        {n}
                      </option>
                    ))}
                  </select>

                  <span
                    className="pointer-events-none absolute right-1.5 top-1/2 -translate-y-1/2 text-gray-500"
                    aria-hidden="true"
                  >
                    <IconChevronDown />
                  </span>
                </div>

                <span>entries</span>
              </div>
            </div>

            {/* Card - opens straight onto the table, as in the design */}
            <div className={`${card} overflow-hidden`}>
              <ProductMasterTable
                rows={pageRows}
                sortKey={sortKey}
                sortDir={sortDir}
                onSort={handleSort}
              />

              {/* Pagination - summary on the left, controls on the right */}
              <div className="px-4 py-3 border-t border-[#eeeeee] bg-[#fafafa]">
                <div className="flex items-center justify-between gap-3 flex-wrap">
                  <span className="text-[12px] text-gray-500">{summaryText}</span>

                  <div className="flex items-center gap-2 flex-wrap">
                    <button
                      type="button"
                      onClick={() => setPageIndex(0)}
                      disabled={safePageIndex === 0}
                      className={pagerButton}
                    >
                      First
                    </button>

                    <button
                      type="button"
                      onClick={() => setPageIndex(Math.max(0, safePageIndex - 1))}
                      disabled={safePageIndex === 0}
                      className={pagerButton}
                    >
                      Previous
                    </button>

                    <span className="px-3 text-[12px] text-gray-700">
                      Page {safePageIndex + 1} of {totalPages}
                    </span>

                    <button
                      type="button"
                      onClick={handleNext}
                      disabled={!canGoNext || searching}
                      className={pagerButton}
                    >
                      Next
                    </button>

                    <button
                      type="button"
                      onClick={() => setPageIndex(totalPages - 1)}
                      disabled={atBufferEnd}
                      className={pagerButton}
                    >
                      Last
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
          )}
        </div>
      </div>
    </div>
  );
}
