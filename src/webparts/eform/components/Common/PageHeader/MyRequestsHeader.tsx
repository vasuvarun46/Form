import * as React from "react";
import { useIsDesktop } from "../../hooks/useIsDesktop";

type MyRequestHeaderProps = {
  onCreate: () => void;
  manualSelectedCount: number;
  onExportExcel: (mode: "all" | "selected") => Promise<void>;
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
    manualSelectedCount,
    onExportExcel,
    onSelectAllRequests,
    onRestoreManualSelection,
    canExport,
    showDisabledExport,
  } = props;

  const isDesktop = useIsDesktop(900);

  // Set to true to re-enable the dropdown popup
  const SHOW_EXPORT_POPUP = false;

  const [open, setOpen] = React.useState(false);
  const [isExporting, setIsExporting] = React.useState(false);
  const [mode, setMode] = React.useState<"all" | "selected">("all");
  const ref = React.useRef<HTMLDivElement | null>(null);

  const canShowExport = canExport || showDisabledExport === true;
  const exportButtonDisabled = !canExport || isExporting;

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent): void => {
      if (ref.current && !ref.current.contains(event.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  React.useEffect(() => {
    if (!open) return;
    if (manualSelectedCount > 0) {
      setMode("selected");
      onRestoreManualSelection();
    } else {
      setMode("all");
      onSelectAllRequests();
    }
  }, [open]);

  React.useEffect(() => {
    if (!canExport && open) setOpen(false);
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

            {/* EXPORT */}
            {canShowExport && (
              <div className="relative" ref={ref}>
                <button
                  type="button"
                  disabled={exportButtonDisabled}
                  onClick={async () => {
                    if (exportButtonDisabled) return;
                    if (SHOW_EXPORT_POPUP) {
                      setOpen((prev) => !prev);
                    } else {
                      const hasSelection = manualSelectedCount > 0;
                      setIsExporting(true);
                      try {
                        if (hasSelection) {
                          onRestoreManualSelection();
                          await onExportExcel("selected");
                        } else {
                          onSelectAllRequests();
                          await onExportExcel("all");
                        }
                      } finally {
                        setIsExporting(false);
                      }
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
                  {isExporting ? "Exporting..." : "Export"}
                </button>

                {/* Popup — hidden while SHOW_EXPORT_POPUP is false */}
                {SHOW_EXPORT_POPUP && open && (
                  <div className="absolute right-0 mt-2 w-64 bg-white border border-gray-200 rounded-lg shadow-lg z-50 p-4 space-y-3">
                    <div className="text-sm font-semibold text-gray-700">
                      Export request for Excel
                    </div>

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

                    <label className="flex items-center gap-2 text-sm cursor-pointer">
                      <input
                        type="radio"
                        checked={mode === "selected"}
                        onChange={() => {
                          setMode("selected");
                          onRestoreManualSelection();
                        }}
                      />
                      Selected: {manualSelectedCount} requests
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
