import * as React from "react";
import { useIsDesktop } from "../hooks/useIsDesktop";

type RowId = string | number;

type BaseRow = {
  id: RowId;
  status: React.ReactNode;

};

type ColumnKey = "id" | "status" | "export" | "approval" | (string & {});

type Column<T extends BaseRow> = {
  key: ColumnKey;
  header: string;
  render: (row: T) => React.ReactNode;
  thClassName?: string;
  tdClassName?: string;
  checked?: boolean;
};

type ResponsiveTableProps<T extends BaseRow> = {
  columns: Array<Column<T>>;
  rows: T[];
  minWidthClass?: string;
  tableBreakpoint?: number;
  wrapHeaderWords?: boolean;
  disableHorizontalScroll?: boolean;

  showCheckbox?: boolean;
  selectedIds?: RowId[];
  onToggleRow?: (id: RowId) => void;
  onToggleAll?: () => void;
};

export function ResponsiveTable<T extends BaseRow>({
  columns,
  rows,
  minWidthClass = "min-w-[1200px]",
  tableBreakpoint = 1024,
  wrapHeaderWords = false,
  disableHorizontalScroll = false,
  showCheckbox,
  selectedIds = [],
  onToggleRow,
  onToggleAll,
}: ResponsiveTableProps<T>): JSX.Element {
  const isTable = useIsDesktop(tableBreakpoint);
  const noHorizontalScroll = disableHorizontalScroll || minWidthClass === "min-w-0";
  const forceWrappedHeader = wrapHeaderWords || noHorizontalScroll;

  /* ===================== MOBILE / CARD VIEW ===================== */
  if (!isTable) {
    return (
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        {rows.map((r) => (
          <div
            key={String(r.id)}
            className="bg-white border border-[#e5e5e5] rounded-xl shadow-sm overflow-hidden"
          >
            {/* Header */}
            <div className="px-4 py-3 border-b border-[#efefef] flex items-start justify-between gap-4">
              <div className="min-w-0">
                <div className="text-[10px] text-gray-500">Transaction ID</div>
                <div className="text-[12px] font-semibold text-gray-900 truncate">
                  {String(r.id)}
                </div>
              </div>

              <div className="text-right">
                <div className="text-[10px] text-gray-500">Status</div>
                <div className="text-[12px] font-semibold">
                  {r.status ?? "-"}
                </div>
              </div>

              {showCheckbox && (
                <input
                  type="checkbox"
                  checked={selectedIds.indexOf(String(r.id)) !== -1}
                  onChange={() => onToggleRow?.(r.id)}
                  className="h-3.5 w-3.5 rounded border-gray-300 accent-[#2563eb] cursor-pointer"
                />
              )}
            </div>

            {/* Body */}
            <div className="px-4 py-3 space-y-2">
              {columns.map((c) => {
                if (c.key === "id" || c.key === "status") return null;

                if (c.key === "export" || c.key === "approval") {
                  return (
                    <div
                      key={c.key}
                      className="flex items-start justify-between gap-3 pt-2"
                    >
                      <div className="text-[11px] text-gray-500">
                        {c.header}
                      </div>
                      <div className="flex flex-wrap justify-end gap-2">
                        {c.render(r)}
                      </div>
                    </div>
                  );
                }

                return (
                  <div
                    key={c.key}
                    className="grid grid-cols-[120px_1fr] gap-3"
                  >
                    <div className="text-[11px] text-gray-500">
                      {c.header}
                    </div>
                    <div className="text-[12px] text-gray-900 text-right break-words">
                      {c.render(r)}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>
    );
  }
  const allSelected =
    rows.length > 0 &&
    rows.every(r => selectedIds.indexOf(String(r.id)) !== -1);

  const someSelected = selectedIds.length > 0 && selectedIds.length < rows.length;

  /* ===================== DESKTOP / TABLE VIEW ===================== */
  // Hide table + header entirely when no rows
  if (rows.length === 0) {
    return (
      <div className="px-4 py-10 text-center">
        {/* Red search icon */}
        <div className="flex justify-center mb-3">
          <svg
            width="40"
            height="40"
            viewBox="0 0 24 24"
            fill="none"
            xmlns="http://www.w3.org/2000/svg"
          >
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

        {/* Title */}
        <div className="text-[13px] font-bold text-gray-800 mb-1">
          No search results found
        </div>

        {/* Subtitle */}
        <div className="text-[12px] font-semibold text-[#9F9F9F]">
          Try changing the filters or search terms.
        </div>
      </div>
    );
  }

  return (
    <div
      className={`pdf-thai ${noHorizontalScroll ? "overflow-x-hidden" : "overflow-x-auto"}`}
      style={noHorizontalScroll ? { overflowX: "hidden" } : undefined}
    >
      <table
        className={`w-full ${noHorizontalScroll ? "min-w-0 table-fixed" : `${minWidthClass} table-fixed`} text-[10px]`}

        style={{ tableLayout: "fixed" }}

      >
        <thead className="bg-[#f8f8f8] border-b border-[#eeeeee]">

          <tr>
            {showCheckbox && (
              <th className="w-6 px-2 text-center align-middle">
                {/* NO indeterminate to avoid "-" */}
                <input
                  type="checkbox"
                  checked={allSelected}
                  onChange={onToggleAll}
                  className={[
                    "h-3.5 w-3.5 rounded border-gray-300 accent-[#2563eb] cursor-pointer",
                    // optional: if someSelected, keep it visually normal (no dash)
                    someSelected ? "opacity-100" : "",
                  ].join(" ")}
                />
              </th>
            )}

            {columns.map((c) => (
              <th
                key={c.key}
                className={[
                  `px-1.5 py-1.5 font-semibold text-[9px] text-gray-700 ${forceWrappedHeader ? "whitespace-normal break-normal leading-tight" : "whitespace-nowrap"}`,
                  (!noHorizontalScroll && c.key === "export" && "min-w-[160px]"),
                  (!noHorizontalScroll && c.key === "approval" && "min-w-[180px]"),
                  c.thClassName && /text-(left|center|right)/.test(c.thClassName)
                    ? c.thClassName
                    : `${c.thClassName || ""} text-left`,
                ].join(" ")}
              >

                {forceWrappedHeader
                  ? c.header.split('<br/>').filter(Boolean).map((word, index, arr) => (
                    <React.Fragment key={`${c.key}-${word}-${index}`}>
                      {word}
                      {index < arr.length - 1 ? <br /> : null}
                    </React.Fragment>
                  ))
                  : c.header}
              </th>
            ))}
          </tr>
        </thead>

        <tbody className="divide-y divide-[#eeeeee] bg-white">
          {rows.map((r) => (
            <tr
              key={String(r.id)}
              className="hover:bg-[#fafafa]"
            >
              {showCheckbox && (
                <td className="px-2 text-center align-middle">
                  <input
                    type="checkbox"
                    checked={selectedIds.indexOf(String(r.id)) !== -1}
                    onChange={() => onToggleRow?.(r.id)}
                    className="h-3.5 w-3.5 rounded border-gray-300 accent-[#2563eb] cursor-pointer"
                  />
                </td>
              )}

              {columns.map((c) => (
                <td
                  key={c.key}
                  className={[
                    "px-1.5 py-1.5 text-[10px] text-gray-700",
                    (c.key === "id")
                      ? "whitespace-nowrap break-normal"
                      : (c.key === "export" || c.key === "approval")
                        ? "whitespace-normal"
                        : "break-words",
                    c.tdClassName ?? "align-top",
                  ].join(" ")}
                >
                  <div
                    className={
                      c.key === "export"
                        ? "flex flex-col items-start gap-2"
                        : c.key === "approval"
                          ? "flex flex-col items-start gap-2"
                          : ""
                    }
                  >
                    {c.render(r)}
                  </div>
                </td>

              ))}
            </tr>
          ))}
          {rows.length === 0 && (
            <tr>
              <td
                colSpan={columns.length + (showCheckbox ? 1 : 0)}
                className="px-4 py-10 text-center"
              >
                {/* Red search icon */}
                <div className="flex justify-center mb-3">
                  <svg
                    width="40"
                    height="40"
                    viewBox="0 0 24 24"
                    fill="none"
                    xmlns="http://www.w3.org/2000/svg"
                  >
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

                {/* Title */}
                <div className="text-[13px] font-semibold text-gray-800 mb-1">
                  No search results found
                </div>

                {/* Subtitle */}
                <div className="text-[12px] text-gray-500">
                  Try changing the filters or search terms.
                </div>

              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}
