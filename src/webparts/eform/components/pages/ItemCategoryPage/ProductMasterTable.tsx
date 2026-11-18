import * as React from "react";
import { useIsDesktop } from "../../hooks/useIsDesktop";
import { ProductMasterRow, ProductSortKey } from "./itemCategoryTypes";

type SortDir = "asc" | "desc";

type Column = {
  key: ProductSortKey;
  header: string;      // "<br/>" marks the wrap point, same as the request tables
  minW: string;        // Tailwind min-width class — the table scrolls horizontally
  numeric?: boolean;
  wrap?: boolean;      // only the long free-text column wraps; the rest stay on one line
  wch: number;         // Excel column width, same convention as ExportExcel.ts
};

/* The table is wider than the viewport by design: every value is shown in full and
   the card scrolls sideways, instead of truncating cells to fit (the previous
   table-fixed + percentage layout hid Thai product names and 13-digit SKUs). */
export const PRODUCT_COLUMNS: Column[] = [
  { key: "sku", header: "SKU<br/>Code", minW: "min-w-[120px]", wch: 18 },
  { key: "productCode", header: "Product<br/>Code", minW: "min-w-[105px]", wch: 14 },
  { key: "colour", header: "Colour<br/>Desc", minW: "min-w-[90px]", wch: 12 },
  { key: "size", header: "Size<br/>Desc", minW: "min-w-[105px]", wch: 16 },
  { key: "subCollection", header: "SubCollection<br/>Desc", minW: "min-w-[115px]", wch: 16 },
  { key: "brandFamily", header: "Brand Family<br/>Desc", minW: "min-w-[115px]", wch: 16 },
  { key: "cluster", header: "Product/Brand<br/>Cluster Desc", minW: "min-w-[115px]", wch: 18 },
  { key: "category", header: "Category<br/>Desc", minW: "min-w-[90px]", wch: 12 },
  { key: "series", header: "Series<br/>Desc", minW: "min-w-[100px]", wch: 14 },
  { key: "subSeries", header: "SubSeries<br/>Desc", minW: "min-w-[105px]", wch: 14 },
  { key: "productName", header: "Product<br/>Name", minW: "min-w-[240px]", wrap: true, wch: 42 },
  { key: "noffQty", header: "NOFF<br/>Qty", minW: "min-w-[75px]", numeric: true, wch: 10 },
];

const cellValue = (row: ProductMasterRow, key: ProductSortKey): string => {
  const raw = row[key];
  if (key === "noffQty") return Number(raw).toLocaleString();
  return String(raw ?? "").trim();
};

function SortCaret({ active, dir }: { active: boolean; dir: SortDir }): JSX.Element {
  return (
    <svg
      width="7"
      height="11"
      viewBox="0 0 10 14"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      className="shrink-0"
    >
      {/* 0.6 keeps the inactive affordance at >=3:1 against the #f8f8f8 header (WCAG 1.4.11) */}
      <path d="M2 5l3-3 3 3" opacity={active ? (dir === "asc" ? 1 : 0.3) : 0.6} />
      <path d="M2 9l3 3 3-3" opacity={active ? (dir === "desc" ? 1 : 0.3) : 0.6} />
    </svg>
  );
}

function NoResults(): JSX.Element {
  return (
    <div className="px-4 py-10 text-center">
      <div className="flex justify-center mb-3">
        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <circle cx="11" cy="11" r="7" stroke="#e53e3e" strokeWidth="2" />
          <line
            x1="16.5"
            y1="16.5"
            x2="21"
            y2="21"
            stroke="#e53e3e"
            strokeWidth="2"
            strokeLinecap="round"
          />
        </svg>
      </div>
      <div className="text-[13px] font-bold text-gray-800 mb-1">
        No search results found
      </div>
      <div className="text-[12px] font-semibold text-[#9F9F9F]">
        Try changing the filters or search terms.
      </div>
    </div>
  );
}

type Props = {
  rows: ProductMasterRow[];
  sortKey: ProductSortKey;
  sortDir: SortDir;
  onSort: (key: ProductSortKey) => void;
};

export default function ProductMasterTable({
  rows,
  sortKey,
  sortDir,
  onSort,
}: Props): JSX.Element {
  const isTable = useIsDesktop(1024);

  if (rows.length === 0) return <NoResults />;

  /* ===================== MOBILE / CARD VIEW ===================== */
  if (!isTable) {
    return (
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 p-3">
        {rows.map((r) => (
          <div
            key={r.id}
            className="bg-white border border-[#e5e5e5] rounded-xl shadow-sm overflow-hidden"
          >
            <div className="px-4 py-3 border-b border-[#efefef] flex items-start justify-between gap-4">
              <div className="min-w-0">
                <div className="text-[10px] text-gray-500">SKU Code</div>
                <div className="text-[12px] font-semibold text-gray-900 truncate">
                  {r.sku}
                </div>
              </div>
              <div className="text-right shrink-0">
                <div className="text-[10px] text-gray-500">NOFF Qty</div>
                <div className="text-[12px] font-semibold text-gray-900">
                  {r.noffQty.toLocaleString()}
                </div>
              </div>
            </div>

            <div className="px-4 py-3 space-y-2">
              {PRODUCT_COLUMNS.filter(
                (c) => c.key !== "sku" && c.key !== "noffQty"
              ).map((c) => (
                <div key={c.key} className="grid grid-cols-[120px_1fr] gap-3">
                  <div className="text-[11px] text-gray-500">
                    {c.header.split("<br/>").join(" ")}
                  </div>
                  <div className="text-[12px] text-gray-900 text-right break-words">
                    {cellValue(r, c.key) || "—"}
                  </div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    );
  }

  /* ===================== DESKTOP / TABLE VIEW ===================== */
  return (
    <div className="pdf-thai overflow-x-auto" style={{ overflowX: "auto" }}>
      <table
        aria-label="Product master"
        className="w-full min-w-[1380px] text-[11px]"
      >
        <thead className="bg-[#f8f8f8] border-b border-[#eeeeee]">
          <tr>
            {PRODUCT_COLUMNS.map((c) => {
              const active = sortKey === c.key;
              // Every column is left-aligned, including the numeric one —
              // AllRequestPage pins even Total Quantity / Total Tag Price to text-left.
              const headerClass = [
                c.minW,
                "px-3 py-2 font-semibold text-[10px] text-gray-700",
                "whitespace-normal break-normal leading-tight text-left",
              ].join(" ");
              const buttonClass = [
                "w-full flex items-start justify-start gap-1 text-left",
                active ? "text-gray-900" : "text-gray-700",
                "hover:text-gray-900",
              ].join(" ");

              return (
                <th
                  key={c.key}
                  scope="col"
                  aria-sort={
                    active ? (sortDir === "asc" ? "ascending" : "descending") : "none"
                  }
                  className={headerClass}
                >
                  <button type="button" onClick={() => onSort(c.key)} className={buttonClass}>
                    <span>
                      {c.header
                        .split("<br/>")
                        .filter(Boolean)
                        .map((word, i, arr) => (
                          <React.Fragment key={`${c.key}-${word}-${i}`}>
                            {word}
                            {i < arr.length - 1 ? <br /> : null}
                          </React.Fragment>
                        ))}
                    </span>
                    <SortCaret active={active} dir={sortDir} />
                  </button>
                </th>
              );
            })}
          </tr>
        </thead>

        <tbody className="divide-y divide-[#eeeeee] bg-white">
          {rows.map((r) => (
            <tr key={r.id} className="hover:bg-[#fafafa]">
              {PRODUCT_COLUMNS.map((c) => {
                const tdClass = [
                  c.minW,
                  "px-3 py-2 text-[11px] text-gray-700 align-top text-left",
                  // Nothing is truncated any more: short columns stay on one line and
                  // the card scrolls sideways; only Product Name wraps.
                  c.wrap ? "whitespace-normal break-words" : "whitespace-nowrap",
                ].join(" ");

                return (
                  <td key={c.key} className={tdClass}>
                    {cellValue(r, c.key) || "—"}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
