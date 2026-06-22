import * as React from "react";
import { useIsDesktop } from "../../hooks/useIsDesktop";

type MyRequestHeaderProps = {
  onCreate: () => void;
  selectedCount: number;
  manualSelectedCount: number;
  onExportExcel: (mode: "all" | "selected") => Promise<void>;
  onExportFirst: () => Promise<void>;
  onSelectAllRequests: () => void;
  onRestoreManualSelection: () => void;
  canExport: boolean;
  showDisabledExport?: boolean;
};

export default function MyRequestHeader(
  props: MyRequestHeaderProps
): JSX.Element {
  const {
    onCreate,
    selectedCount,
    manualSelectedCount,
    onExportExcel,
    onExportFirst,
    onSelectAllRequests,
    onRestoreManualSelection,
    canExport,
    showDisabledExport,
  } = props;

  const isDesktop = useIsDesktop(900);

  const [open, setOpen] = React.useState(false);
  const [isExporting, setIsExporting] = React.useState(false);
  const [mode, setMode] = React.useState<"all" | "selected">("all");
  const ref = React.useRef<HTMLDivElement | null>(null);

  const canShowExport = canExport || showDisabledExport === true;
  const exportButtonDisabled = !canExport || isExporting;

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

  // ✅ FIX: When dropdown opens, set correct default mode
  // Use manualSelectedCount (not selectedCount) to decide
  React.useEffect(() => {
    if (!open) return;

    if (manualSelectedCount > 0) {
      // User has manually checked rows → show "selected" mode
      setMode("selected");
      onRestoreManualSelection();
    } else {
      // No manual selection → default to "all"
      setMode("all");
      onSelectAllRequests();
    }
  }, [open]); // ✅ only run when dropdown opens/closes

  // ✅ Close dropdown if export permission is revoked
  React.useEffect(() => {
    if (!canExport && open) {
      setOpen(false);
    }
  }, [canExport, open]);

  return (
    <div className="bg-[#f5f5f5] border-b border-[#e5e5e5] pdf-thai">
      <div className={isDesktop ? "w-full mx-auto px-6 py-4" : "w-full mx-auto px-4 py-3"}>
        <div className={isDesktop ? "flex items-center justify-between" : "flex flex-col items-start gap-3"}>

          {/* LEFT */}
          <div>
            <div className="text-[18px] font-bold text-gray-900">
              Internal Usage Inventory Request Form
            </div>
            <div className="text-[12px] text-gray-500">
              Overview
            </div>
          </div>

          {/* RIGHT */}
          <div className={isDesktop ? "flex items-center gap-3" : "flex w-full items-center gap-2"}>

            {/* EXPORT DROPDOWN */}
            {canShowExport && (
              <div className="relative" ref={ref}>
                <button
                  type="button"
                  disabled={exportButtonDisabled}
                  onClick={() => {
                    if (exportButtonDisabled) return;
                    if (selectedCount === 0) {
                      // Nothing selected → directly download the first visible record
                      setIsExporting(true);
                      onExportFirst().finally(() => setIsExporting(false));
                    } else {
                      setOpen((prev) => !prev);
                    }
                  }}
                  className={[
                    isDesktop
                      ? "h-8 px-3 rounded-md text-[12px] font-medium transition"
                      : "h-9 px-3 rounded-md text-[12px] font-medium transition",
                    exportButtonDisabled
                      ? "bg-gray-200 text-gray-500 cursor-not-allowed"
                      : "bg-[#E8E8E8] text-gray-700 hover:bg-[#E8E8E8]",
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
                          onSelectAllRequests(); // select all rows in table
                        }}
                      />
                      All requests
                    </label>

                    {/* ✅ Selected radio — always shows manualSelectedCount */}
                    <label className="flex items-center gap-2 text-sm cursor-pointer">
                      <input
                        type="radio"
                        checked={mode === "selected"}
                        onChange={() => {
                          setMode("selected");
                          onRestoreManualSelection(); // restore manually checked rows
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
            )}

            {/* CREATE */}
            <button
              onClick={onCreate}
              className={
                isDesktop
                  ? "h-8 px-4 rounded-md bg-[#000000] text-white text-[12px] font-semibold hover:bg-[#3B3B3B]"
                  : "h-9 px-4 rounded-md bg-[#000000] text-white text-[12px] font-semibold hover:bg-[#3B3B3B]"

              }
            >
              Create request
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}