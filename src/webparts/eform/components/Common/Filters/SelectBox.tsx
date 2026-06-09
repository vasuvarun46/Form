import * as React from "react";

type Props = {
  placeholder: string;
  kind?: "multi" | "dateRange";
  options?: string[];
  selected?: string[];
  from?: string;
  to?: string;
  onSelectedChange?: (v: string[]) => void;
  onDateRangeChange?: (from: string, to: string) => void;
};

const WEEKDAY_LABELS: string[] = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];

function pad2(v: number): string {
  return String(v).padStart(2, "0");
}

function toIsoLocalDate(date: Date): string {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function parseIsoLocalDate(value: string): Date | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ""));
  if (!m) return null;

  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const parsed = new Date(y, mo - 1, d);

  if (
    parsed.getFullYear() !== y ||
    parsed.getMonth() !== mo - 1 ||
    parsed.getDate() !== d
  ) {
    return null;
  }
  return parsed;
}

function formatDisplayDate(value: string): string {
  const d = parseIsoLocalDate(value);
  if (!d) return "--";
  return `${pad2(d.getDate())}-${pad2(d.getMonth() + 1)}-${d.getFullYear()}`;
}

function addDays(base: Date, days: number): Date {
  const next = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  next.setDate(next.getDate() + days);
  return next;
}

function getGridStart(viewMonth: Date): Date {
  const monthStart = new Date(viewMonth.getFullYear(), viewMonth.getMonth(), 1);
  return addDays(monthStart, -monthStart.getDay());
}

function getMonthLabel(viewMonth: Date): string {
  return viewMonth.toLocaleDateString("en-US", { month: "short", year: "numeric" });
}

export default function SelectBox({
  placeholder,
  kind = "multi",
  options = [],
  selected = [],
  from = "",
  to = "",
  onSelectedChange,
  onDateRangeChange,
}: Props): JSX.Element {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);

  const [viewMonth, setViewMonth] = React.useState<Date>(() => {
    const anchor = parseIsoLocalDate(from) || parseIsoLocalDate(to) || new Date();
    return new Date(anchor.getFullYear(), anchor.getMonth(), 1);
  });

  // 🔹 UPDATED LOGIC HERE (only change)
  React.useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) {
        setOpen(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  React.useEffect(() => {
    if (!open || kind !== "dateRange") return;
    const anchor = parseIsoLocalDate(from) || parseIsoLocalDate(to) || new Date();
    setViewMonth(new Date(anchor.getFullYear(), anchor.getMonth(), 1));
  }, [open, kind, from, to]);

  const display =
    kind === "dateRange"
      ? from || to
        ? `${formatDisplayDate(from)} to ${formatDisplayDate(to)}`
        : ""
      : selected.length === 0
        ? ""
        : selected.length === 1
          ? selected[0]
          : `${selected.length} selected`;

  const toggle = (v: string): void => {
    if (!onSelectedChange) return;

    const idx = selected.indexOf(v);
    const next = idx >= 0
      ? selected.filter((x) => x !== v)
      : [...selected, v];

    onSelectedChange(next);

    // ✅ Auto close after select
    setOpen(false);
  };

  const fromSafe = from || "";
  const toSafe = to || "";

  const rangeStart =
    fromSafe && toSafe
      ? (fromSafe <= toSafe ? fromSafe : toSafe)
      : fromSafe || toSafe;

  const rangeEnd =
    fromSafe && toSafe
      ? (fromSafe <= toSafe ? toSafe : fromSafe)
      : fromSafe || toSafe;

  const dayCells = React.useMemo(() => {
    if (kind !== "dateRange") return [];
    const start = getGridStart(viewMonth);
    return Array.from({ length: 42 }, (_, idx) => addDays(start, idx));
  }, [kind, viewMonth]);

  const pickDate = (day: Date): void => {
    if (!onDateRangeChange) return;
    const iso = toIsoLocalDate(day);

    if (!from && !to) {
      onDateRangeChange(iso, iso);
      return;
    }

    if (from && to && from !== to) {
      onDateRangeChange(iso, iso);
      return;
    }

    if (from && to && from === to) {
      if (iso < from) {
        onDateRangeChange(iso, from);
      } else {
        onDateRangeChange(from, iso);
      }
      return;
    }

    if (from && !to) {
      if (iso < from) {
        onDateRangeChange(iso, from);
      } else {
        onDateRangeChange(from, iso);
      }
      return;
    }

    if (!from && to) {
      if (iso <= to) {
        onDateRangeChange(iso, to);
      } else {
        onDateRangeChange(to, iso);
      }
      return;
    }
  };

  return (
    <div ref={ref} className="relative pdf-thai">
      {/* Trigger */}
      <button
        type="button"
        onClick={() => setOpen((p) => !p)}
        className={`
w-full border-2 border-[#e5e7eb]  bg-[#F8F8F8]
px-3 py-2 text-left rounded-lg
hover:border-[#A3A3A3]
focus:outline-none
${open ? "rounded-b-none border-b-0" : ""}
`}
      >
        {kind === "dateRange" ? (
          <>
            <div className="flex items-start justify-between">
              <div className={`text-[10px] ${open ? "font-bold text-gray-900" : "font-medium text-gray-600"}`}>
                {placeholder}
              </div>

              <span
                className={`shrink-0 h-5 w-5 rounded-full bg-[#FBE9EA] hover:bg-[#E8E8E8] flex items-center justify-center transition-transform duration-200 ${open ? "rotate-180" : ""}`}
              >
                <svg width="12" height="12" viewBox="0 0 24 24" fill="#E60023">
                  <path d="M7.41 8.59 12 13.17l4.59-4.58L18 10l-6 6-6-6z" />
                </svg>
              </span>
            </div>

            {!open && display && (
              <div className="truncate text-[11px] text-gray-700">
                {display}
              </div>
            )}
          </>
        ) : (
          <div className="flex items-center justify-between">
            <div className={`text-[10px] ${open ? "font-bold text-gray-900" : "font-medium text-gray-600"}`}>
              {placeholder}
            </div>

            <span
              className={`shrink-0 h-5 w-5 rounded-full bg-[#FBE9EA] hover:bg-[#E8E8E8] flex items-center justify-center transition-transform duration-200 ${open ? "rotate-180" : ""}`}
            >
              <svg width="12" height="12" viewBox="0 0 24 24" fill="#E60023">
                <path d="M7.41 8.59 12 13.17l4.59-4.58L18 10l-6 6-6-6z" />
              </svg>
            </span>
          </div>
        )}

        {kind !== "dateRange" && display && (
          <div className="truncate text-[11px] text-gray-700">
            {display}
          </div>
        )}
      </button>

      {/* Dropdown */}
      {open && (
        <div
          className={[
            "absolute z-50 w-full border border-[#e5e7eb]  bg-[#F8F8F8] shadow-md",

            kind === "dateRange"
              ? "mt-0 rounded-b-lg rounded-t-none border-t-0"
              : "mt-0 rounded-b-lg rounded-t-none border-t-0"
          ].join(" ")}
        >
          {kind === "dateRange" ? (
            <div className="p-2">
              <div className="flex h-8 items-center justify-between rounded-md border border-[#2172DA] bg-white px-2">
                <span
                  className={[
                    "truncate text-[11px]",
                    from || to ? "text-[#6b7280]" : "text-[#9ca3af]",
                  ].join(" ")}
                >
                  {display}
                </span>
                <svg
                  width="14"
                  height="14"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="#2172DA"
                  strokeWidth="1.8"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  aria-hidden="true"
                >
                  <rect x="3" y="4" width="18" height="17" rx="2" />
                  <line x1="3" y1="9" x2="21" y2="9" />
                  <line x1="8" y1="2" x2="8" y2="6" />
                  <line x1="16" y1="2" x2="16" y2="6" />
                  <circle cx="8" cy="13" r="0.8" fill="#2172DA" />
                  <circle cx="12" cy="13" r="0.8" fill="#2172DA" />
                  <circle cx="16" cy="13" r="0.8" fill="#2172DA" />
                  <circle cx="8" cy="17" r="0.8" fill="#2172DA" />
                  <circle cx="12" cy="17" r="0.8" fill="#2172DA" />
                  <circle cx="16" cy="17" r="0.8" fill="#2172DA" />
                </svg>
              </div>

              <div className="mt-1 rounded-md border border-[#ececec] bg-[#FFFFFF] p-2">
                <div className="flex items-center justify-between">
                  <button
                    type="button"
                    onClick={() =>
                      setViewMonth((prev) => new Date(prev.getFullYear(), prev.getMonth() - 1, 1))
                    }
                    className="h-6 w-6 flex items-center justify-center text-[18px] text-[#6b7280] hover:text-[#2172DA] leading-none select-none"
                    aria-label="Previous month"
                  >
                    ←
                  </button>

                  <div className="text-[11px] font-semibold text-[#4b5563]">
                    {getMonthLabel(viewMonth)}
                  </div>

                  <button
                    type="button"
                    onClick={() =>
                      setViewMonth((prev) => new Date(prev.getFullYear(), prev.getMonth() + 1, 1))
                    }
                    className="h-6 w-6 flex items-center justify-center text-[18px] text-[#6b7280] hover:text-[#2172DA] leading-none select-none"
                    aria-label="Next month"
                  >
                    →
                  </button>
                </div>

                <div className="mt-2 grid grid-cols-7 text-center text-[10px] text-[#9ca3af]">
                  {WEEKDAY_LABELS.map((day) => (
                    <div key={day}>{day}</div>
                  ))}
                </div>

                <div className="mt-1 grid grid-cols-7 gap-y-1">
                  {dayCells.map((day) => {
                    const iso = toIsoLocalDate(day);
                    const inCurrentMonth = day.getMonth() === viewMonth.getMonth();
                    const isStart = iso === fromSafe;
                    const isEnd = iso === toSafe;
                    const isSelected = isStart || isEnd;
                    const isInRange =
                      !!rangeStart && !!rangeEnd && iso > rangeStart && iso < rangeEnd;
                    const hasBothEnds = !!fromSafe && !!toSafe && fromSafe !== toSafe;

                    const wrapperStyle: React.CSSProperties =
                      isInRange
                        ? { backgroundColor: "#E3F0FF" }
                        : isStart && hasBothEnds
                          ? { background: "linear-gradient(to right, transparent 50%, #E3F0FF 50%)" }
                          : isEnd && hasBothEnds
                            ? { background: "linear-gradient(to left, transparent 50%, #E3F0FF 50%)" }
                            : {};

                    return (
                      <div
                        key={iso}
                        className="flex justify-center"
                        style={wrapperStyle}
                      >
                        <button
                          type="button"
                          onClick={() => pickDate(day)}
                          className={[
                            "h-6 w-6 rounded-full text-[10px] transition-colors relative z-10",
                            isSelected
                              ? "bg-[#2172DA] text-white font-semibold"
                              : isInRange
                                ? "text-[#2172DA]"
                                : inCurrentMonth
                                  ? "text-[#4b5563] hover:bg-[#F8F8F8]"
                                  : "text-[#d1d5db] cursor-default",
                          ].join(" ")}
                          aria-label={`Pick ${iso}`}
                        >
                          {day.getDate()}
                        </button>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          ) : (
            <div className="max-h-[260px] overflow-auto p-3 space-y-2  bg-[#F8F8F8]">

              {options.map((opt) => (
                <label
                  key={opt}
                  className="flex items-center gap-2 text-[11px] text-gray-700 cursor-pointer"
                >
                  <input
                    type="checkbox"
                    checked={selected.indexOf(opt) !== -1}
                    onChange={() => toggle(opt)}
                  />
                  {opt}
                </label>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}