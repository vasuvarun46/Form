

import * as React from "react";

type SelectBoxProps = {
  placeholder: string;
  kind?: "multi" | "dateRange";
  options?: string[];
  selected?: string[];
  from?: string;
  to?: string;
  onSelectedChange?: (v: string[]) => void;
  onDateRangeChange?: (from: string, to: string) => void;
  searchable?: boolean;
};

type RequestFilterCardProps = {
  statusOptions: string[];
  purposeOptions: string[];
  costCenterOptions: string[];

  statusSelected: string[];
  purposeSelected: string[];
  costCenterSelected: string[];

  expectedFrom: string;
  expectedTo: string;

  setStatusSelected: (v: string[]) => void;
  setPurposeSelected: (v: string[]) => void;
  setCostCenterSelected: (v: string[]) => void;
  setExpectedFrom: (v: string) => void;
  setExpectedTo: (v: string) => void;
  onClearAll?: () => void;
  showClearAll?: boolean;


  SelectBox: React.FC<SelectBoxProps>; // ✅ injected properly
};

export default function RequestFilterCard(
  props: RequestFilterCardProps
): JSX.Element {
  const {
    statusOptions,
    purposeOptions,
    costCenterOptions,

    statusSelected,
    purposeSelected,
    costCenterSelected,

    expectedFrom,
    expectedTo,

    setStatusSelected,
    setPurposeSelected,
    setCostCenterSelected,
    setExpectedFrom,
    setExpectedTo,
    onClearAll,
    showClearAll = false,


    SelectBox, // ✅ FIX: destructured correctly
  } = props;


  
  return (
    <div className="bg-white rounded-xl border border-[#e5e5e5] shadow-sm p-5 sm:p-6 pdf-thai">
      {/* Header */}
      <div className="flex items-center gap-2 mb-3">
        <svg
          width="18"
          height="18"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.6"
          strokeLinecap="round"
          strokeLinejoin="round"
          className="text-gray-600"
        >
          <line x1="4" y1="6" x2="20" y2="6" />
          <circle cx="9" cy="6" r="2" />
          <line x1="4" y1="12" x2="20" y2="12" />
          <circle cx="15" cy="12" r="2" />
        </svg>

        <div className="text-[12px] font-semibold text-gray-900">Filter</div>
        {onClearAll && showClearAll && (
          <button
            type="button"
            onClick={onClearAll}
            className="ml-auto h-7 px-3 rounded-md bg-[#f5f5f5] border border-[#e5e5e5] text-[11px] text-gray-600 hover:bg-[#ececec]"
          >
            Clear all
          </button>
        )}
      </div>

      {/* Filters */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-3">
        <SelectBox
          placeholder="Status"
          kind="multi"
          options={statusOptions}
          selected={statusSelected}
          onSelectedChange={(v) => {
            setStatusSelected(v);
            
          }}
        />

        <SelectBox
          placeholder="Purpose"
          kind="multi"
          options={purposeOptions}
          selected={purposeSelected}
          onSelectedChange={(v) => {
            setPurposeSelected(v);
           
          }}
        />

        <SelectBox
          placeholder="Cost center"
          kind="multi"
          options={costCenterOptions}
          selected={costCenterSelected}
          searchable={true}
          onSelectedChange={(v) => {
            setCostCenterSelected(v);

          }}
        />

        <SelectBox
          placeholder="Expected delivery date"
          kind="dateRange"
          from={expectedFrom}
          to={expectedTo}
          onDateRangeChange={(from, to) => {
            setExpectedFrom(from);
            setExpectedTo(to);
           
          }}
        />
      </div>
    </div>
  );
}
