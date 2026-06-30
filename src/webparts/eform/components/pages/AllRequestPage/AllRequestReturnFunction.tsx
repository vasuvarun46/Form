import * as React from "react";
import RequestEmptyState from "../RequestEmptyState";
import Loading from "../../ui/Loading";
import AllRequestHeader from "../../Common/PageHeader/AllRequestHeader";
import RequestFilterCard from "../../Common/Filters/RequestFilters";
import SelectBox from "../../Common/Filters/SelectBox";
import RequestSearchBar from "../../Common/Search/RequestSearchBar";
import { ResponsiveTable } from "../../Common/ResponsiveTable";
import { Row } from "./AllRequestTypes";
import * as strings from "EformWebPartStrings";
import { SearchFieldKey } from "../../Common/Search/RequestSearchBar";

type TableColumn<T> = {
  key: string;
  header: string;
  render: (row: T) => React.ReactNode;
  thClassName?: string;
  tdClassName?: string;
  checked?: boolean;
};

type Props = {
  pageBg: string;
  container: string;
  card: string;

  loadingPage: boolean;
  tableLoading: boolean;
  isError: boolean;
  isBaseEmpty: boolean;
  hasAnyData: boolean;
  error: string;

  statusOptions: string[];
  purposeOptions: string[];
  costCenterOptions: string[];

  statusFilter: string[];
  purposeFilter: string[];
  costCenterFilter: string[];

  expectedFrom: string;
  expectedTo: string;

  setStatusFilter: (v: string[]) => void;
  setPurposeFilter: (v: string[]) => void;
  setCostCenterFilter: (v: string[]) => void;
  setExpectedFrom: (v: string) => void;
  setExpectedTo: (v: string) => void;

  searchDraft: string;
  setSearchDraft: (v: string) => void;
  onSearch: () => void;
  onResetPage: () => void;
  onSortToggle: () => void;

  pageRows: Row[];
  filteredRows: Row[];
  columns: TableColumn<Row>[];

  pageSize: number;
  safePageIndex: number;
  totalPages: number;

  goFirst: () => void;
  goPrev: () => void;
  goNext: () => void;
  goLast: () => void;

  selectedIds: string[];
  handleToggleRow: (id: string | number) => void;
  handleToggleAll: () => void;
  onSelectAllRequests: () => void;
  onRestoreManualSelection: () => void;
  onExportExcel: (mode: "all" | "selected") => Promise<void>;
  headerChecked: boolean;

  onSearchWithField?: (args: { field: SearchFieldKey; text: string }) => void;
  onClear?: () => void;

  canExport: boolean;
  showDisabledExport: boolean;
  manualSelectedCount: number;
  countIncreased: boolean;
};

export default function AllRequestRender({
  pageBg,
  container,
  card,
  loadingPage,
  tableLoading,
  isError,
  isBaseEmpty,
  hasAnyData,
  error,
  statusOptions,
  purposeOptions,
  costCenterOptions,
  statusFilter,
  purposeFilter,
  costCenterFilter,
  expectedFrom,
  expectedTo,
  setStatusFilter,
  setPurposeFilter,
  setCostCenterFilter,
  setExpectedFrom,
  setExpectedTo,
  searchDraft,
  setSearchDraft,
  onSearch,
  pageRows,
  filteredRows,
  columns,
  safePageIndex,
  totalPages,
  goFirst,
  goPrev,
  goNext,
  goLast,
  onResetPage,
  selectedIds,
  handleToggleRow,
  handleToggleAll,
  onSelectAllRequests,
  onRestoreManualSelection,
  onExportExcel,
  headerChecked,
  onSearchWithField,
  onClear,
  canExport,
  showDisabledExport,
  manualSelectedCount,
  countIncreased
}: Props): JSX.Element {

  const hasActiveFilters =
    statusFilter.length > 0 ||
    purposeFilter.length > 0 ||
    costCenterFilter.length > 0 ||
    !!expectedFrom ||
    !!expectedTo;

  return (
    <div className="pdf-thai flex flex-col h-full min-h-0 overflow-hidden">
      {/* LOADING */}
      {loadingPage && (
        <div className={`${pageBg} flex-1 min-h-0 px-3 sm:px-6`}>
          <div className={`${container}`}>
            <div className={`${card} p-6`}>
              <Loading
                title={strings.LoadingTitle}
                description={strings.LoadingDescription}
              />
            </div>
          </div>
        </div>
      )}

      {/* ERROR */}
      {isError && (
        <div className={`${pageBg} flex-1 min-h-0 px-3 sm:px-6 py-4`}>
          <div className={`${container}`}>
            <RequestEmptyState
              title={strings.EmptyTitle}
              description={strings.EmptyTeamRequests}
              showAction={false}
            />
          </div>
        </div>
      )}

      {/* EMPTY */}
      {isBaseEmpty && (
        <>
          <div className="sticky top-0 z-20 bg-white shrink-0 border-b border-gray-200">
            <AllRequestHeader
              loading={true}
              onSelectAllRequests={onSelectAllRequests}
              onRestoreManualSelection={onRestoreManualSelection}
              onExportExcel={onExportExcel}
              canExport={canExport}
              showDisabledExport={showDisabledExport}
              manualSelectedCount={manualSelectedCount}
            />
          </div>
          <div className={`${pageBg} flex-1 min-h-0 px-3 sm:px-6 py-4`}>
            <div className={`${container}`}>
              <RequestEmptyState
                title={strings.EmptyTitle}
                description={strings.EmptyTeamRequests}
                showAction={false}
              />
            </div>
          </div>
        </>
      )}

      {/* DATA */}
      {hasAnyData && (
        <div className="pdf-thai flex flex-col h-full min-h-0 overflow-hidden">

          {/* ================= HEADER ================= */}
          <div className="sticky top-0 z-20 bg-white shrink-0 border-b border-gray-200">
            <AllRequestHeader
              loading={tableLoading || filteredRows.length === 0}
              onSelectAllRequests={onSelectAllRequests}
              onRestoreManualSelection={onRestoreManualSelection}
              onExportExcel={onExportExcel}
              canExport={canExport}
              manualSelectedCount={manualSelectedCount}
            />
          </div>

          {/* ================= BODY (ONLY SCROLL AREA) ================= */}
          <div className={`flex-1 overflow-y-auto overflow-x-hidden bg-[#f3f3f3] py-4 px-6 sm:px-6`}>

            <div className={` space-y-6`}>

              {/* Filters */}
              <RequestFilterCard
                statusOptions={statusOptions}
                purposeOptions={purposeOptions}
                costCenterOptions={costCenterOptions}
                statusSelected={statusFilter}
                purposeSelected={purposeFilter}
                costCenterSelected={costCenterFilter}
                expectedFrom={expectedFrom}
                expectedTo={expectedTo}
                setStatusSelected={(v) => { setStatusFilter(v); onResetPage(); }}
                setPurposeSelected={(v) => { setPurposeFilter(v); onResetPage(); }}
                setCostCenterSelected={(v) => { setCostCenterFilter(v); onResetPage(); }}
                setExpectedFrom={(v) => { setExpectedFrom(v); onResetPage(); }}
                setExpectedTo={(v) => { setExpectedTo(v); onResetPage(); }}
                onClearAll={() => {
                  setStatusFilter([]);
                  setPurposeFilter([]);
                  setCostCenterFilter([]);
                  setExpectedFrom("");
                  setExpectedTo("");
                  onResetPage();
                }}
                showClearAll={hasActiveFilters}
                SelectBox={SelectBox}
              />

              {/* Title */}
              <div className="text-[13px] font-semibold text-gray-900">
                {strings.RequestListTitle}
              </div>

              {/* Card */}

              <div key={String(countIncreased)} className={`${card} overflow-hidden`}>

                {/* Search */}
                <div className="px-4 py-3 flex items-center justify-between border-b border-[#eeeeee]">
                  <RequestSearchBar
                    searchDraft={searchDraft}
                    setSearchDraft={setSearchDraft}
                    onSearch={onSearch}
                    onSearchWithField={onSearchWithField}
                    onClear={onClear}
                  />
                </div>

                {/* Error */}
                {error && (
                  <div className="px-4 py-3 text-[12px] text-red-600">
                    {error}
                  </div>
                )}

                {/* Table */}
                {tableLoading ? (
                  <div className="px-4 py-6">
                    <Loading />
                  </div>
                ) : (
                  <ResponsiveTable<Row>
                    rows={pageRows}
                    columns={columns}
                    minWidthClass="min-w-0"
                    tableBreakpoint={1024}
                    wrapHeaderWords={true}
                    disableHorizontalScroll={true}
                    showCheckbox={true}
                    selectedIds={selectedIds}
                    onToggleRow={handleToggleRow}
                    onToggleAll={handleToggleAll}
                    headerChecked={headerChecked}
                  />
                )}

                {/* Pagination */}
                <div className="px-4 py-3 border-t border-[#eeeeee] bg-[#fafafa]">
                  <div className="flex items-center gap-2 flex-wrap">

                    <button
                      onClick={goFirst}
                      disabled={safePageIndex === 0}
                      className="h-7 px-2.5 rounded-md border text-[12px] font-medium bg-white border-gray-300 text-[#2172DA] hover:bg-[#2172DA] hover:text-white disabled:cursor-not-allowed disabled:border-[#D9D9D9] disabled:bg-[#D9D9D9] disabled:text-gray-500"

                    >
                      First
                    </button>

                    <button
                      onClick={goPrev}
                      disabled={safePageIndex === 0}
                      className="h-7 px-2.5 rounded-md border text-[12px] font-medium bg-white border-gray-300 text-[#2172DA] hover:bg-[#2172DA] hover:text-white disabled:cursor-not-allowed disabled:border-[#D9D9D9] disabled:bg-[#D9D9D9] disabled:text-gray-500"

                    >
                      Previous
                    </button>

                    <span className="px-3 text-[12px] text-gray-700">
                      Page {safePageIndex + 1} of {totalPages}
                    </span>

                    <button
                      onClick={goNext}
                      disabled={safePageIndex >= totalPages - 1}
                      className="h-7 px-2.5 rounded-md border text-[12px] font-medium bg-white border-gray-300 text-[#2172DA] hover:bg-[#2172DA] hover:text-white disabled:cursor-not-allowed disabled:border-[#D9D9D9] disabled:bg-[#D9D9D9] disabled:text-gray-500"
                    >
                      Next
                    </button>

                    <button
                      onClick={goLast}
                      disabled={safePageIndex >= totalPages - 1}
                      className="h-7 px-2.5 rounded-md border text-[12px] font-medium bg-white border-gray-300 text-[#2172DA] hover:bg-[#2172DA] hover:text-white disabled:cursor-not-allowed disabled:border-[#D9D9D9] disabled:bg-[#D9D9D9] disabled:text-gray-500"
                    >
                      Last
                    </button>

                    {/* <span className="ml-auto text-[12px] text-gray-500">
                      Showing {pageRows.length} of {filteredRows.length} items (pageSize {pageSize})
                    </span> */}

                  </div>
                </div>

              </div>

            </div>

          </div>

        </div>
      )}

    </div>
  );
}
