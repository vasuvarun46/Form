import * as React from "react";
import BaseInnerHeader from "./BaseInnerHeader";

type Props = {
  loading?: boolean;
  manualSelectedCount: number;
  onExportExcel: (mode: "all" | "selected") => Promise<void>;
  onSelectAllRequests: () => void;
  onRestoreManualSelection: () => void;
  canExport: boolean;
  showDisabledExport?: boolean;
};

export default function AllRequestHeader(props: Props): JSX.Element {
  const {
    loading,
    manualSelectedCount,
    onExportExcel,
    onSelectAllRequests,
    onRestoreManualSelection,
    canExport,
    showDisabledExport,
  } = props;

  const [open, setOpen] = React.useState(false);
  const [isExporting, setIsExporting] = React.useState(false);
  const [mode, setMode] = React.useState<"all" | "selected">("all");
  const ref = React.useRef<HTMLDivElement | null>(null);
  const canShowExport = canExport || showDisabledExport === true;
  const exportButtonDisabled = !canExport || !!loading || isExporting;

  // ✅ Close when clicking outside
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (ref.current && !ref.current.contains(event.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  // ✅ FIX: Use manualSelectedCount, only run when dropdown opens
  React.useEffect(() => {
    if (!open) return;

    if (manualSelectedCount > 0) {
      setMode("selected");
      onRestoreManualSelection();
    } else {
      setMode("all");
      onSelectAllRequests();
    }
  }, [open]); // ✅ intentionally exclude manualSelectedCount & callbacks

  // ✅ Close dropdown if export permission is revoked
  React.useEffect(() => {
    if (!canExport && open) {
      setOpen(false);
    }
  }, [canExport, open]);

  return (
    <BaseInnerHeader
      title="All requests"
      subtitle="Overview"
      right={
        <div className="flex items-center gap-2">
          {canShowExport ? (
            <div className="relative" ref={ref}>
              <button
                type="button"
                disabled={exportButtonDisabled}
                onClick={() => {
                  if (!exportButtonDisabled) setOpen((prev) => !prev);
                }}
                className={[
                  "h-8 px-3 rounded-md text-[12px] font-medium transition",
                  exportButtonDisabled
                    ? "bg-gray-200 text-gray-500 cursor-not-allowed"
                    : "bg-[#eeeeee] text-gray-700 hover:bg-[#e0e0e0]",
                ].join(" ")}
              >
                Export
              </button>

              {open && (
                <div className="absolute right-0 mt-2 w-64 bg-white border border-gray-200 rounded-lg shadow-lg z-50 p-4 space-y-3">
                  <div className="text-sm font-semibold text-gray-700">
                    Export request for Excel
                  </div>

                  {/* ✅ All requests radio */}
                  <label className="flex items-center gap-2 text-sm cursor-pointer">
                    <input
                      type="radio"
                      checked={mode === "all"}
                      onChange={() => {
                        setMode("all");
                        onSelectAllRequests();
                      }}
                    />
                    All requests
                  </label>

                  {/* ✅ Selected radio — always shows manualSelectedCount */}
                  <label className="flex items-center gap-2 text-sm cursor-pointer">
                    <input
                      type="radio"
                      checked={mode === "selected"}
                      disabled={manualSelectedCount === 0}   // ✅ disable if no manual selection
                      onChange={() => {
                        setMode("selected");
                        onRestoreManualSelection();
                      }}
                    />
                    Selected: {manualSelectedCount} requests  {/* ✅ FIXED */}
                  </label>

                  <button
                    onClick={async () => {
                      if (isExporting) return;
                      setIsExporting(true);
                      try {
                        await onExportExcel(mode);
                        setOpen(false);
                      } finally {
                        setIsExporting(false);
                      }
                    }}
                    disabled={isExporting}
                    className="w-full rounded-md bg-blue-600 py-1.5 text-[12px] text-white hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-60"
                  >
                    {isExporting ? "Exporting..." : "Export for Excel"}
                  </button>
                </div>
              )}
            </div>
          ) : null}
        </div>
      }
    />
  );
}