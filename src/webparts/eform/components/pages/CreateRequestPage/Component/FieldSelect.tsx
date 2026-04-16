import * as React from "react";
import { IconChevronDown } from "../../../ui/icons";

export default function FieldSelect({
  label,
  value,
  onChange,
  options,
  error,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  options: string[];
  error?: string;
}): JSX.Element {
  return (
    <div className="pdf-thai">
      <div className="relative">
        {/* SELECT */}
        <select
          value={value || ""}
          onChange={(e) => onChange(e.target.value)}
          className={[
            "w-full h-12 rounded-md bg-white px-3 pt-5 text-[12px]",
            "appearance-none",
            error ? "border border-red-300" : "border border-[#e7e7e7]",
          ].join(" ")}
        >
          {/* placeholder */}
          <option value="" disabled>
            Select {label}
          </option>

         {options.length === 0 ? (
  <option value="" disabled>
    No options available
  </option>
) : (
  options.map((op) => (
    <option key={op} value={op}>
      {op}
    </option>
  ))
)}
        </select>
         <span className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-gray-500">
            <IconChevronDown />
              </span>
        {/* LABEL INSIDE BOX */}
        <div
          className="  absolute left-3 top-1.5 text-[10px] text-gray-500 bg-white px-1   pointer-events-none" >
          <span className="text-[#E60023]">*</span> {label}
        </div>
      </div>

      {error && <div className="mt-1 text-[10px] text-red-700">{error}</div>}
    </div>
  );
}