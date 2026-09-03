import * as React from "react";

import { IconChevronDown } from "../../../ui/icons";
import FieldTextarea from "../Component/FieldTextarea";
import { CostCenterOption, IUsageState, PurposeOption, UsageFieldErrors } from "../createRequestTypes";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { useSharePointTimezoneOffset } from "./SpService"

const _MONTH_NAMES = [
  "January","February","March","April","May","June",
  "July","August","September","October","November","December",
];
const _DAY_HDRS = ["Su","Mo","Tu","We","Th","Fr","Sa"];

type Props = {
  usage: IUsageState;
  setField: <K extends keyof IUsageState>(k: K, v: IUsageState[K]) => void;
  purposes: PurposeOption[];
  costCenters: CostCenterOption[];
  fieldErrors: UsageFieldErrors;
  disabled?: boolean;
  context: WebPartContext;
  textFieldsDisabled?: boolean;
};

// ── Shared hook: close on outside mousedown or Escape ──────────────────────
function useOutsideClose(ref: React.RefObject<HTMLDivElement>, open: boolean, close: () => void): void {
  React.useEffect(() => {
    if (!open) return;
    const onMouse = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) close();
    };
    const onKey = (e: KeyboardEvent): void => { if (e.key === "Escape") close(); };
    document.addEventListener("mousedown", onMouse);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onMouse);
      document.removeEventListener("keydown", onKey);
    };
  }, [open, ref, close]);
}

// ── Shared trigger style ───────────────────────────────────────────────────
function FieldTrigger({
  label,
  bgColor,
  open,
  disabled,
  children,
  onClick,
  onKeyDown,
}: {
  label: string;
  bgColor: string;
  open: boolean;
  disabled?: boolean;
  children: React.ReactNode;
  onClick: () => void;
  onKeyDown: React.KeyboardEventHandler<HTMLDivElement>;
}): JSX.Element {
  return (
    <div
      className="group relative"
      onClick={onClick}
      role="button"
      tabIndex={disabled ? -1 : 0}
      onKeyDown={onKeyDown}
    >
      <div
        className={[
          "w-full h-[52px] rounded-md pl-3 pr-8 pt-[18px] pb-1 text-[12px] select-none border-2 transition-colors",
          disabled ? "cursor-not-allowed" : "cursor-pointer",
          open ? "border-blue-500" : "border-[#e7e7e7] group-hover:border-[#76A8E8]",
        ].join(" ")}
        style={{ backgroundColor: bgColor }}
      >
        {children}
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
  );
}

// ── FieldSelect — custom popup dropdown ───────────────────────────────────
function FieldSelect({
  label,
  value,
  onChange,
  options,
  error,
  disabled,
  grayBackground = false,
  searchable = false,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  options: string[];
  error?: string;
  disabled?: boolean;
  grayBackground?: boolean;
  searchable?: boolean;
}): JSX.Element {
  const bgColor = disabled && grayBackground ? "#f3f4f6" : "#ffffff";
  const textColor = disabled && grayBackground ? "#6b7280" : "#374151";

  const [open, setOpen] = React.useState(false);
  const [hoveredOp, setHoveredOp] = React.useState<string | null>(null);
  const [search, setSearch] = React.useState("");
  const innerRef = React.useRef<HTMLDivElement>(null);
  const searchRef = React.useRef<HTMLInputElement>(null);
  const close = React.useCallback(() => setOpen(false), []);
  useOutsideClose(innerRef, open, close);

  React.useEffect(() => {
    if (!open) { setSearch(""); return; }
    if (searchable && searchRef.current) {
      const t = setTimeout(() => searchRef.current?.focus(), 30);
      return () => clearTimeout(t);
    }
  }, [open, searchable]);

  const toggle = (): void => { if (!disabled) setOpen((o) => !o); };

  const visibleOptions = searchable && search.trim()
    ? options.filter((op) => op.toLowerCase().includes(search.toLowerCase()))
    : options;

  return (
    <div>
      {/* inner ref wraps ONLY the 52px trigger so top-full isn't offset by error rows */}
      <div ref={innerRef} className="relative">
        <FieldTrigger
          label={label}
          bgColor={bgColor}
          open={open}
          disabled={disabled}
          onClick={toggle}
          onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggle(); } }}
        >
          {value
            ? <span style={{ color: open ? "#9ca3af" : textColor }}>{value}</span>
            : <span className="text-gray-400">Select {label}</span>}
        </FieldTrigger>

        {open && (
          <div
            className="absolute left-0 top-full bg-white border border-[#e0e0e0] rounded-lg shadow-xl overflow-hidden"
            style={{ zIndex: 9999, width: "100%" }}
          >
            {searchable && (
              <div className="p-2 border-b border-gray-100">
                <input
                  ref={searchRef}
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Search..."
                  className="w-full px-2 py-1 text-[12px] border border-gray-200 rounded outline-none focus:border-blue-400"
                  onClick={(e) => e.stopPropagation()}
                />
              </div>
            )}
            <div className="overflow-y-auto eform-thin-scroll" style={{ maxHeight: 240 }}>
              {/* Placeholder row: grey when nothing hovered AND no value; white otherwise */}
              <div
                className="px-4 py-2 text-[12px] cursor-default select-none"
                style={{
                  backgroundColor: (!value && !hoveredOp) ? "#C8C8C8" : "#ffffff",
                  color: (!value && !hoveredOp) ? "#ffffff" : "#9ca3af",
                }}
              >
                Select {label}
              </div>

              {visibleOptions.map((op) => (
                <button
                  key={op}
                  type="button"
                  onClick={() => { onChange(op); setOpen(false); }}
                  onMouseEnter={() => setHoveredOp(op)}
                  onMouseLeave={() => setHoveredOp(null)}
                  style={{
                    backgroundColor: hoveredOp === op ? "#1967D2" : "#ffffff",
                    color: hoveredOp === op ? "white" : "#374151",
                  }}
                  className="w-full text-left px-4 py-2 text-[12px]"
                >
                  {op}
                </button>
              ))}
              {visibleOptions.length === 0 && (
                <div className="px-4 py-2 text-[12px] text-gray-400">No results</div>
              )}
            </div>
          </div>
        )}
      </div>
      {error && <div className="mt-1 text-[10px] text-[#E56D71]">{error}</div>}
    </div>
  );
}

// ── FieldDate — custom calendar popup ────────────────────────────────────
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
  const bgColor = disabled && grayBackground ? "#f3f4f6" : "#ffffff";
  const textColor = disabled && grayBackground ? "#6b7280" : "#1f2937";

  const todayStr = React.useMemo((): string => new Date().toISOString().split("T")[0], []);

  const toLocalDateString = (val: string): string => {
    if (!val) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(val)) return val;
    const utcMs = new Date(val).getTime();
    if (isNaN(utcMs)) return "";
    const localMs = utcMs + timezoneOffsetMinutes * 60 * 1000;
    const d = new Date(localMs);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
  };

  const formatDisplay = (val: string): string => {
    const n = toLocalDateString(val);
    if (!n) return "";
    const [yyyy, mm, dd] = n.split("-");
    return `${dd}-${mm}-${yyyy}`;
  };

  const normalizedValue = toLocalDateString(value);
  const displayValue = formatDisplay(value);

  const [open, setOpen] = React.useState(false);
  const [calView, setCalView] = React.useState<"days" | "years">("days");
  const [viewYear, setViewYear] = React.useState(0);
  const [viewMonth, setViewMonth] = React.useState(0);
  const [yearRangeStart, setYearRangeStart] = React.useState(0);
  const innerRef = React.useRef<HTMLDivElement>(null);
  const close = React.useCallback(() => setOpen(false), []);
  useOutsideClose(innerRef, open, close);

  const toggle = (): void => {
    if (disabled) return;
    if (open) { setOpen(false); return; }
    const base = normalizedValue || todayStr;
    const [y, m] = base.split("-").map(Number);
    setViewYear(y);
    setViewMonth(m - 1);
    setCalView("days");
    setOpen(true);
  };

  const prevMonth = (): void => {
    if (viewMonth === 0) { setViewMonth(11); setViewYear((y) => y - 1); }
    else setViewMonth((m) => m - 1);
  };
  const nextMonth = (): void => {
    if (viewMonth === 11) { setViewMonth(0); setViewYear((y) => y + 1); }
    else setViewMonth((m) => m + 1);
  };

  const openYears = (): void => {
    setYearRangeStart(viewYear - (viewYear % 12));
    setCalView("years");
  };

  // Calendar grid cells
  const firstWeekday = new Date(viewYear, viewMonth, 1).getDay();
  const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
  const cells: (number | null)[] = [
    ...Array(firstWeekday).fill(null),
    ...Array.from({ length: daysInMonth }, (_, i) => i + 1),
  ];
  while (cells.length % 7 !== 0) cells.push(null);

  const toDayStr = (day: number): string =>
    `${viewYear}-${String(viewMonth + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;

  const yearRange = Array.from({ length: 12 }, (_, i) => yearRangeStart + i);

  return (
    <div>
      {/* inner ref wraps ONLY the 52px trigger so top-full isn't offset by sibling error rows */}
      <div ref={innerRef} className="relative">
        <FieldTrigger
          label={label}
          bgColor={bgColor}
          open={open}
          disabled={disabled}
          onClick={toggle}
          onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggle(); } }}
        >
          {displayValue
            ? <span style={{ color: textColor }}>{displayValue}</span>
            : <span className="text-gray-400">Select {label}</span>}
        </FieldTrigger>

        {open && (
          <div
            className="absolute left-0 top-full bg-white border border-[#e0e0e0] rounded-lg shadow-xl"
            style={{ width: 280, zIndex: 9999 }}
          >
            {calView === "days" ? (
              <>
                {/* Month navigation header */}
                <div className="flex items-center justify-between px-3 pt-3 pb-1">
                  <button
                    type="button" onClick={prevMonth}
                    className="w-7 h-7 flex items-center justify-center rounded hover:bg-gray-100 text-gray-600 text-lg leading-none select-none"
                  >‹</button>
                  <button
                    type="button" onClick={openYears}
                    className="text-[12px] font-semibold text-gray-800 hover:bg-gray-100 rounded px-2 py-0.5 select-none"
                  >
                    {_MONTH_NAMES[viewMonth]} {viewYear}
                  </button>
                  <button
                    type="button" onClick={nextMonth}
                    className="w-7 h-7 flex items-center justify-center rounded hover:bg-gray-100 text-gray-600 text-lg leading-none select-none"
                  >›</button>
                </div>

                {/* Day-of-week headers */}
                <div className="grid grid-cols-7 px-2 pb-0.5">
                  {_DAY_HDRS.map((d) => (
                    <div key={d} className="text-center text-[10px] text-gray-400 py-0.5 select-none">{d}</div>
                  ))}
                </div>

                {/* Day cells */}
                <div className="grid grid-cols-7 px-2 pb-2">
                  {cells.map((day, i) => {
                    if (day === null) return <div key={`e-${i}`} />;
                    const ds = toDayStr(day);
                    const isPast = ds < todayStr;
                    const isSelected = ds === normalizedValue;
                    const isToday = ds === todayStr;
                    return (
                      <button
                        key={`d-${i}`}
                        type="button"
                        disabled={isPast}
                        onClick={() => { onChange(ds); setOpen(false); }}
                        className={[
                          "h-7 w-full flex items-center justify-center rounded text-[11px]",
                          isPast ? "text-gray-300 cursor-not-allowed"
                            : isSelected ? "text-white font-semibold"
                            : isToday ? "bg-blue-50 text-blue-600 font-semibold hover:bg-blue-100"
                            : "text-gray-700 hover:bg-blue-100",
                        ].join(" ")}
                        style={isSelected ? { backgroundColor: "#1967D2" } : undefined}
                      >{day}</button>
                    );
                  })}
                </div>

                {/* Footer */}
                <div className="flex items-center justify-between px-3 pb-3 pt-1 border-t border-gray-100">
                  <button
                    type="button"
                    onClick={() => { onChange(""); setOpen(false); }}
                    className="text-[11px] text-gray-500 hover:text-gray-700"
                  >Clear</button>
                  <button
                    type="button"
                    onClick={() => { onChange(todayStr); setOpen(false); }}
                    className="text-[11px] text-blue-500 hover:text-blue-700 font-medium"
                  >Today</button>
                </div>
              </>
            ) : (
              /* Year picker view */
              <>
                <div className="flex items-center justify-between px-3 pt-3 pb-2">
                  <button
                    type="button" onClick={() => setYearRangeStart((s) => s - 12)}
                    className="w-7 h-7 flex items-center justify-center rounded hover:bg-gray-100 text-gray-600 text-lg leading-none select-none"
                  >‹</button>
                  <span className="text-[12px] font-semibold text-gray-800 select-none">
                    {yearRangeStart} – {yearRangeStart + 11}
                  </span>
                  <button
                    type="button" onClick={() => setYearRangeStart((s) => s + 12)}
                    className="w-7 h-7 flex items-center justify-center rounded hover:bg-gray-100 text-gray-600 text-lg leading-none select-none"
                  >›</button>
                </div>
                <div className="grid grid-cols-4 gap-1 px-3 pb-3">
                  {yearRange.map((y) => (
                    <button
                      key={y}
                      type="button"
                      onClick={() => { setViewYear(y); setCalView("days"); }}
                      className={[
                        "py-1.5 rounded text-[11px] font-medium",
                        y === viewYear ? "bg-blue-500 text-white" : "text-gray-700 hover:bg-gray-100",
                      ].join(" ")}
                    >{y}</button>
                  ))}
                </div>
              </>
            )}
          </div>
        )}
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
  const textFieldDisabled = textFieldsDisabled || disabled;
  const timezoneOffsetMinutes = useSharePointTimezoneOffset(context);

  return (
    <section className="rounded-xl border border-[#e5e5e5] shadow-sm p-4 sm:p-5 pdf-thai bg-white transition-colors duration-200">
      <h2 className="text-[13px] font-bold text-gray-900 mb-4">
        Usage Information
      </h2>

      <div className="grid grid-cols-1 xl:grid-cols-3 gap-x-10 gap-y-4 -ml-3">
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
          searchable
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
