import * as React from "react";
import { IconSearch } from "../../ui/icons";
import {
  ProductFilterKey,
  ProductFilters,
  PRODUCT_FILTER_FIELDS,
} from "./itemCategoryTypes";

type Props = {
  filters: ProductFilters;
  onChange: (key: ProductFilterKey, value: string) => void;
  onClearAll: () => void;
  statusText: string;
};

export default function ProductFilterCard(props: Props): JSX.Element {
  const { filters, onChange, onClearAll, statusText } = props;

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
          aria-hidden="true"
        >
          <line x1="4" y1="6" x2="20" y2="6" />
          <circle cx="9" cy="6" r="2" />
          <line x1="4" y1="12" x2="20" y2="12" />
          <circle cx="15" cy="12" r="2" />
        </svg>

        <div className="text-[12px] font-semibold text-gray-900">Filter</div>
      </div>

      {/* Fields — label above a single-line input, as in the approved design */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5 gap-x-4 gap-y-3">
        {PRODUCT_FILTER_FIELDS.map((f) => {
          const value = filters[f.key] || "";
          const hasValue = value.trim().length > 0;

          const inputClass = [
            "h-8 w-full rounded-md border bg-[#F8F8F8] pl-8 pr-8",
            "text-[11px] text-gray-700 placeholder-gray-500",
            "hover:border-[#A3A3A3]",
            // The browser's default focus ring is a thick black outline. Replace it
            // with the app's blue: a 1px border plus a soft halo, so focus stays
            // visible (WCAG 2.4.7) without reading as a heavy border.
            "focus:outline-none focus:border-[#2172DA] focus:bg-white",
            "focus:shadow-[0_0_0_3px_rgba(33,114,218,0.15)]",
            hasValue ? "border-[#2172DA] bg-white" : "border-[#e5e7eb]",
          ].join(" ");

          return (
            <div key={f.key} className="min-w-0">
              <label
                htmlFor={`pm-filter-${f.key}`}
                className={`block mb-1 text-[11px] ${
                  hasValue ? "font-semibold text-gray-900" : "font-medium text-gray-700"
                }`}
              >
                {f.label}
              </label>

              <div className="relative">
                <span
                  className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-500"
                  aria-hidden="true"
                >
                  <IconSearch className="h-3.5 w-3.5" />
                </span>

                <input
                  id={`pm-filter-${f.key}`}
                  type="text"
                  value={value}
                  onChange={(e) => onChange(f.key, e.target.value)}
                  placeholder={`Search ${f.label}`}
                  autoComplete="off"
                  className={inputClass}
                />

                {hasValue && (
                  <button
                    type="button"
                    onClick={() => onChange(f.key, "")}
                    aria-label={`Clear ${f.label}`}
                    className="absolute right-1.5 top-1/2 -translate-y-1/2 h-5 w-5 rounded-full bg-[#FBE9EA] hover:bg-[#E8E8E8] flex items-center justify-center text-[#E60023]"
                  >
                    <svg
                      width="10"
                      height="10"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="3"
                      strokeLinecap="round"
                      aria-hidden="true"
                    >
                      <path d="M6 6l12 12M18 6L6 18" />
                    </svg>
                  </button>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {/* Status + Clear all */}
      <div className="mt-4 flex items-center justify-between gap-3 flex-wrap">
        <div className="text-[11px] text-gray-500" aria-live="polite">
          {statusText}
        </div>

        <button
          type="button"
          onClick={onClearAll}
          className="flex items-center gap-1.5 h-7 px-3 rounded-md bg-[#f5f5f5] border border-[#e5e5e5] text-[11px] text-gray-600 hover:bg-[#ececec]"
        >
          <svg
            width="12"
            height="12"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2.2"
            strokeLinecap="round"
            strokeLinejoin="round"
            aria-hidden="true"
          >
            <path d="M3 12a9 9 0 0 1 15.5-6.2L21 8" />
            <path d="M21 3v5h-5" />
            <path d="M21 12a9 9 0 0 1-15.5 6.2L3 16" />
            <path d="M3 21v-5h5" />
          </svg>
          Clear all
        </button>
      </div>
    </div>
  );
}
