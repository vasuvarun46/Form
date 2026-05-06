import * as React from "react";

import { IconChevronDown } from "../../../ui/icons";
import FieldTextarea from "../Component/FieldTextarea";
import { CostCenterOption, IUsageState, PurposeOption, UsageFieldErrors } from "../createRequestTypes";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { useSharePointTimezoneOffset } from "./SpService"

type InputWithPicker = HTMLInputElement & {
  showPicker?: () => void;
};

type Props = {
  usage: IUsageState;
  setField: <K extends keyof IUsageState>(k: K, v: IUsageState[K]) => void;
  purposes: PurposeOption[];
  costCenters: CostCenterOption[];
  fieldErrors: UsageFieldErrors;
  disabled?: boolean;
  context: WebPartContext;
  textFieldsDisabled?: boolean; // ✅ NEW: force text fields disabled in manual draft mode

};

function FieldSelect({
  label,
  value,
  onChange,
  options,
  error,
  disabled,
  grayBackground = false,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  options: string[];
  error?: string;
  disabled?: boolean;
  grayBackground?: boolean;
}): JSX.Element {
  const bgColor = disabled && grayBackground ? "#f3f4f6" : "#ffffff";
  const textColor = disabled && grayBackground ? "#6b7280" : "#374151";

  return (
    <div>
      <div className="relative">
        <select
          value={value || ""}
          onChange={(e) => onChange(e.target.value)}
          disabled={disabled}
          className={[
            "w-full h-[52px] rounded-md pl-3 pr-8 pt-[18px] pb-1 text-[12px]",
            "appearance-none",
            "disabled:opacity-100",
            "disabled:cursor-not-allowed",
            "border-2 border-[#e7e7e7] focus:outline-none focus:border-blue-500 focus:ring-0 hover:!border-[#76A8E8]",
          ].join(" ")}
          style={{ backgroundColor: bgColor, color: textColor }}
        >
          <option value="" disabled>
            Select {label}
          </option>
          {options.map((op) => (
            <option key={op} value={op}>
              {op}
            </option>
          ))}
        </select>

        <span className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-gray-500">
          <IconChevronDown />
        </span>

        <div
          className="absolute left-3 top-1.5 text-[10px] text-gray-500 pointer-events-none" style={{ backgroundColor: bgColor }}
        >
          <span className="text-[#E56D71]">*</span>{label}        </div>
      </div>

      {error && <div className="mt-1 text-[10px] text-[#E56D71]">{error}</div>}    </div>
  );
}

function FieldDate({
  label,
  value,
  onChange,
  error,
  disabled,
  grayBackground = false,
  timezoneOffsetMinutes = 0,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  error?: string;
  disabled?: boolean;
  grayBackground?: boolean;
  timezoneOffsetMinutes?: number;
}): JSX.Element {
  const today = new Date().toISOString().split("T")[0];
  const inputRef = React.useRef<HTMLInputElement>(null);

  const bgColor = disabled && grayBackground ? "#f3f4f6" : "#ffffff";
  const textColor = disabled && grayBackground ? "#6b7280" : "#1f2937";

  const toLocalDateString = (val: string): string => {
    if (!val) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(val)) return val;
    const utcMs = new Date(val).getTime();
    if (isNaN(utcMs)) return "";
    const localMs = utcMs + timezoneOffsetMinutes * 60 * 1000;
    const localDate = new Date(localMs);
    const yyyy = localDate.getUTCFullYear();
    const mm = String(localDate.getUTCMonth() + 1).padStart(2, "0");
    const dd = String(localDate.getUTCDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  };

  const formatDisplay = (val: string): string => {
    const normalized = toLocalDateString(val);
    if (!normalized) return "";
    const [yyyy, mm, dd] = normalized.split("-");
    if (!yyyy || !mm || !dd) return "";
    return `${dd}-${mm}-${yyyy}`;
  };

  const normalizedValue = toLocalDateString(value);
  const displayValue = formatDisplay(value);

  return (
    <div>
      <div
        className="group relative cursor-pointer"
        onClick={() => {
          if (disabled) return;
          const el = inputRef.current as InputWithPicker | null;
          if (el && typeof el.showPicker === "function") {
            el.showPicker();
          }
        }}
      >

        <input
          ref={inputRef}
          type="date"
          value={normalizedValue}
          min={today}
          onChange={(e) => onChange(e.target.value)}
          className="absolute inset-0 opacity-0 w-full h-full cursor-pointer"
          tabIndex={-1}
        />

        <div
          className={[
            "w-full h-[52px] rounded-md pl-3 pr-8 pt-[18px] pb-1 text-[12px]",
            "appearance-none select-none",
            disabled ? "cursor-not-allowed" : "cursor-pointer",
            "border-2 border-[#e7e7e7] focus-within:border-blue-500 focus-within:ring-0 group-hover:!border-[#76A8E8]",
          ].join(" ")}
          style={{ backgroundColor: bgColor }}
        >
          {displayValue ? (
            <span style={{ color: textColor }}>{displayValue}</span>
          ) : (
            <span className="text-gray-500">Select {label}</span>
          )}
        </div>


        <span className="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-gray-500">
          <IconChevronDown />
        </span>
        <div
          className="absolute left-3 top-1.5 text-[10px] text-gray-500 pointer-events-none"
          style={{ backgroundColor: bgColor }}
        >
          <span className="text-[#E56D71]">*</span>{label}
        </div>
      </div>

      {error && <div className="mt-1 text-[10px] text-[#E56D71]">{error}</div>}

    </div>
  );
}


export default function UsageInfoSection({
  usage,
  setField,
  purposes,
  costCenters,
  fieldErrors,
  disabled = false,
  context,
  textFieldsDisabled = false,

}: Props): JSX.Element {
  // const dropdownsSelected =
  //   !!usage.purpose &&
  //   !!usage.costCenter &&
  //   !!usage.expectedDelivery;

  const textFieldDisabled = disabled;
  const timezoneOffsetMinutes = useSharePointTimezoneOffset(context);

  return (
    <section className="rounded-xl border border-[#e5e5e5] shadow-sm p-4 sm:p-5 pdf-thai bg-white transition-colors duration-200">
      <h2 className="text-[13px] font-bold text-gray-900 mb-4">
        Usage Information
      </h2>

      <div className="grid grid-cols-1 xl:grid-cols-3 gap-x-10 gap-y-4 -ml-3">        {/* ROW 1 — grayBackground only when isExcelMode (disabled=true) */}
        <FieldSelect
          label="Purpose"
          value={usage.purpose}
          onChange={(v: string) => setField("purpose", v)}
          options={purposes.map((p) => p.Title)}
          error={fieldErrors.purpose}
          disabled={disabled}
          grayBackground={disabled}
        />

        <FieldSelect
          label="Cost Center"
          value={usage.costCenter}
          onChange={(v: string) => setField("costCenter", v)}
          options={costCenters.map((c) => c.Title)}
          error={fieldErrors.costCenter}
          disabled={disabled}
          grayBackground={disabled}
        />

        <FieldDate
          label="Expected Delivery Date"
          value={usage.expectedDelivery}
          onChange={(v: string) => setField("expectedDelivery", v)}
          error={fieldErrors.expectedDelivery}
          disabled={disabled}
          grayBackground={disabled}
          timezoneOffsetMinutes={timezoneOffsetMinutes}
        />

        {/* ROW 2 — grayBackground only when isExcelMode (disabled=true), white when just !dropdownsSelected */}
        <div className="col-span-full space-y-4">
          <FieldTextarea
            label="Place of Delivery"
            maxLength={150}
            value={usage.placeOfDelivery}
            onChange={(v: string) => setField("placeOfDelivery", v)}
            error={fieldErrors.placeOfDelivery}
            placeholder="Enter Place of Delivery"
            disabled={textFieldDisabled}
            grayBackground={disabled}
          />

          <FieldTextarea
            label="Purpose Detail"
            maxLength={250}
            value={usage.purposeDetail}
            onChange={(v: string) => setField("purposeDetail", v)}
            error={fieldErrors.purposeDetail}
            placeholder="Enter Purpose Details"
            disabled={textFieldDisabled}
            grayBackground={disabled}
          />
        </div>
      </div>
    </section>
  );
}