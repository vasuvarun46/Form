import * as React from "react";
import { ResponsiveTable } from "../ResponsiveTable";

export type RequestTableColumn<T> = {
  key: string;
  header: string;
  render: (row: T) => React.ReactNode;
};

type Props<T> = {
  columns: RequestTableColumn<T>[];
  rows: T[];
  showCheckbox?: boolean;

  selectedIds?: RowId[];              // ✅ ARRAY
  onToggleRow?: (id: RowId) => void;  // ✅
  onToggleAll?: () => void; 

  // pagination
  pageIndex: number;
  totalPages: number;
  pageSize: number;
  onFirst: () => void;
  onPrev: () => void;
  onNext: () => void;
  onLast: () => void;
};
type RowId = string | number;

type BaseRow = {
  id: RowId;
  status: React.ReactNode;
};
export default function RequestTable<T extends BaseRow>({
  columns,
  rows,
  pageIndex,
  totalPages,
  pageSize,
  onFirst,
  onPrev,
  onNext,
  onLast,
}: Props<T>): JSX.Element {
  const [selectedIds, setSelectedIds] = React.useState<RowId[]>([]);

  return (
    <div className="bg-white rounded-xl border border-[#e5e5e5] shadow-sm overflow-hidden">

      {/* Table */}
      <div className="overflow-x-auto">
        <ResponsiveTable<T>
          columns={columns}
          rows={rows}
          minWidthClass="min-w-[1200px]"
          tableBreakpoint={1024}
           selectedIds={selectedIds}
  onToggleRow={(id: RowId) =>
    setSelectedIds(prev =>
      prev.indexOf(id) ? prev.filter(x => x !== id) : [...prev, id]
    )
  }
  onToggleAll={() =>
    setSelectedIds(
      selectedIds.length === rows.length ? [] : rows.map(r => r.id)
    )
  }
        />
      </div>

      {/* Pagination */}
      <div className="px-4 py-3 flex items-center gap-2 text-[12px] text-gray-600 border-t border-[#eeeeee]">
        <button
          onClick={onFirst}
          disabled={pageIndex === 0}
          className="h-8 px-3 rounded-md bg-[#e6e6e6] text-gray-700 disabled:opacity-50"
        >
          First
        </button>

        <button
          onClick={onPrev}
          disabled={pageIndex === 0}
          className="h-8 px-3 rounded-md bg-[#e6e6e6] text-gray-700 disabled:opacity-50"
        >
          Previous
        </button>

        <span className="px-2">
          Page {pageIndex + 1} of {totalPages}
        </span>

        <button
          onClick={onNext}
          disabled={pageIndex >= totalPages - 1}
          className="h-8 px-3 rounded-md bg-[#e6e6e6] text-gray-700 disabled:opacity-50"
        >
          Next
        </button>

        <button
          onClick={onLast}
          disabled={pageIndex >= totalPages - 1}
          className="h-8 px-3 rounded-md bg-[#e6e6e6] text-gray-700 disabled:opacity-50"
        >
          Last
        </button>

        <span className="ml-auto text-gray-500">
          Showing {rows.length} items (pageSize {pageSize})
        </span>
      </div>
    </div>
  );
}
