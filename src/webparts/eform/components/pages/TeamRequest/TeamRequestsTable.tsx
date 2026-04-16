/* =========================================================
   TeamRequestsTable.tsx
   âœ… No changes removed
========================================================= */

import * as React from "react";
import { ResponsiveTable } from "../../Common/ResponsiveTable";
import { Row } from "./TeamService";
import Loading from "../../ui/Loading";

export default function TeamRequestsTable(props: {
  columns: Array<{
    key: string;
    header: string;
    render: (r: Row) => JSX.Element | string | number;
    thClassName?: string;
    tdClassName?: string;
  }>;
  pageRows: Row[];
  filteredRows: Row[];
  loading: boolean;
  error: string;
  pageSize: number;
  safePageIndex: number;
  totalPages: number;
  goFirst: () => void;
  goPrev: () => void;
  goNext: () => void;
  goLast: () => void;
  selectedIds: string[];
  onToggleRow: (id: string | number) => void;
  onToggleAll: () => void;
  canExport: boolean;
}): JSX.Element {
  const {
    columns,
    pageRows,
    //filteredRows,
    loading,
    error,
    //pageSize,
    safePageIndex,
    totalPages,
    goFirst,
    goPrev,
    goNext,
    goLast,
  } = props;

  return (
    <>
      {error ? <div className="px-4 sm:px-5 py-3 text-[12px] text-red-600">{error}</div> : null}
      {loading ? (
        <div className="px-4 sm:px-5 py-6">
          <Loading />
        </div>
      ) : null}

      {!loading && <div className="px-0 pr-0">
        <ResponsiveTable<Row>
          columns={columns}
          rows={pageRows}
          minWidthClass="min-w-0"
          tableBreakpoint={1024}
          wrapHeaderWords={true}
          disableHorizontalScroll={true}
          // âœ… ENABLE CHECKBOX
          showCheckbox={props.canExport}
          selectedIds={props.selectedIds}
          onToggleRow={props.onToggleRow}
          onToggleAll={props.onToggleAll}
        />
      </div>}

      {!loading && <div className="px-4 py-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between text-[12px] text-gray-600 border-t border-[#eeeeee] bg-[#fafafa]">
        <div className="flex flex-wrap items-center gap-2">

        <button

          onClick={goFirst}

          disabled={safePageIndex === 0}

          className="h-7 px-2.5 rounded-md border text-[12px] font-medium bg-white border-gray-300 text-[#0a66d8] hover:bg-[#e6f0fb] disabled:bg-gray-100 disabled:text-gray-500 disabled:border-gray-200 disabled:cursor-not-allowed" >

          First
        </button>

        <button

          onClick={goPrev}

          disabled={safePageIndex === 0}

          className="h-7 px-2.5 rounded-md border text-[12px] font-medium bg-white border-gray-300 text-[#0a66d8]  hover:bg-[#e6f0fb]  disabled:bg-gray-100 disabled:text-gray-500 disabled:border-gray-200 disabled:cursor-not-allowed">

          Previous
        </button>

        <span className="px-3 text-[12px] text-gray-700">

          Page {safePageIndex + 1} of {totalPages}
        </span>

        <button

          onClick={goNext}

          disabled={safePageIndex >= totalPages - 1}

          className="h-7 px-2.5 rounded-md border text-[12px] font-medium bg-white border-gray-300 text-[#0a66d8] hover:bg-[#e6f0fb] disabled:bg-gray-100 disabled:text-gray-500 disabled:border-gray-200 disabled:cursor-not-allowed">

          Next
        </button>

        <button

          onClick={goLast}

          disabled={safePageIndex >= totalPages - 1}

          className="h-7 px-2.5 rounded-md border text-[12px] font-medium bg-white border-gray-300 text-[#0a66d8] hover:bg-[#e6f0fb] disabled:bg-gray-100 disabled:text-gray-500 disabled:border-gray-200 disabled:cursor-not-allowed">

          Last
        </button>

        </div>
        {/* <span className="text-[11px] sm:text-[12px] text-gray-500">

          Showing {pageRows.length} of {filteredRows.length} items (pageSize {pageSize})
        </span> */}

      </div>}


    </>
  );
}

