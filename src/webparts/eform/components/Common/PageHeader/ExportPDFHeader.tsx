

import * as React from "react";
import { Route } from "../../Eform";
import { SecurityPermissions } from "../../spHelpers";
import { useIsDesktop } from "../../hooks/useIsDesktop";


interface Props {
  transactionId?: string;
  exportSource?: Route;
  canAct?: boolean;
  isApproveMode?: boolean;
  isBypassMode?: boolean;
  showReject?: boolean;
  canExportPdf?: boolean;
  pdfBusy?: boolean;
  actionBusy?: boolean;
  permissions: SecurityPermissions;

  onNavigate: (route: Route) => void;
  onApprove?: () => void;
  onReject?: () => void;
  onBypass?: () => void;
  onExportPdf?: () => void;
}

export default function ExportPDFHeader({
  transactionId,
  exportSource = "my",
  canAct = false,
  isApproveMode = false,
  isBypassMode = false,
  showReject = false,
  canExportPdf = false,
  pdfBusy = false,
  actionBusy = false,
  permissions,
  onNavigate,
  onApprove,
  onReject,
  onBypass,
  onExportPdf,
}: Props): JSX.Element {
  const isDesktop = useIsDesktop(700);
  const actionButtonSharedClass =
    "font-semibold transition-colors duration-150 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 disabled:cursor-not-allowed disabled:opacity-60";

  const backLabel =
    exportSource === "team"
      ? "Team's requests"
      : exportSource === "all"
        ? "All requests"
        : "My requests";
  const canUseActionButtons =
    exportSource === "team" ? permissions?.approveTeam : permissions?.approve;

  return (
    <div className="shrink-0 sticky top-0 z-20 bg-[#f6f6f6] border-b border-gray-200">

      <div className={isDesktop ? "w-full px-8 lg:px-14 py-4" : "w-full px-3 py-3"}>

        <div className={isDesktop ? "flex flex-row justify-between items-center gap-3" : "flex flex-col items-start gap-3"}>

          {/* LEFT SIDE */}
          <div>
            <h1 className={isDesktop ? "text-xl font-semibold text-gray-900 tracking-tight" : "text-lg font-semibold text-gray-900 tracking-tight"}>
              Transaction ID : {transactionId || ""}
            </h1>

            <div className={isDesktop ? "text-sm text-gray-500 mt-1" : "text-xs text-gray-500 mt-1"}>
              <button
                type="button"
                className="text-blue-600 hover:underline"
                onClick={() => onNavigate(exportSource)}
              >
                {backLabel}
              </button>{" "}
              / Transaction details
            </div>
          </div>

          {/* RIGHT SIDE ACTIONS */}
          <div className={isDesktop ? "w-auto flex flex-wrap items-center justify-end gap-2" : "w-full flex flex-wrap items-center justify-start gap-2"}>

            {canExportPdf && (
              <button
                onClick={onExportPdf}
                disabled={pdfBusy}
                className={isDesktop
                  ? "h-11 min-w-[130px] px-8 bg-[#000000] text-white rounded-md text-[14px] font-semibold hover:bg-[#3B3B3B] disabled:opacity-60"
                  : "h-8 min-w-[100px] px-3 bg-[#000000] text-white rounded-md text-[12px] font-semibold hover:bg-[#3B3B3B] disabled:opacity-60"}

              >
                {pdfBusy ? "Exporting..." : "Export PDF"}
              </button>
            )}

            {canAct && canUseActionButtons && (
              <>
                {(isApproveMode || canAct) && (
                  <button
                    onClick={onApprove}
                    disabled={actionBusy}
                    className={isDesktop
                      ? `h-11 min-w-[130px] px-4 rounded-md border border-[#12a020] bg-[#e6f2e8] text-[#12a020] text-[14px] hover:bg-[#12a020] hover:text-white active:bg-[#0f8a1b] active:text-white focus-visible:outline-[#12a020] ${actionButtonSharedClass}`
                      : `h-8 min-w-[92px] px-3 rounded-md border border-[#12a020] bg-[#e6f2e8] text-[#12a020] text-[12px] hover:bg-[#12a020] hover:text-white active:bg-[#0f8a1b] active:text-white focus-visible:outline-[#12a020] ${actionButtonSharedClass}`}
                  >
                    Approve
                  </button>
                )}

                {isBypassMode && (
                  <button
                    onClick={onBypass}
                    disabled={actionBusy}
                    className={isDesktop
                      ? `h-11 min-w-[130px] px-4 rounded-md border border-[#2563eb] bg-[#eff6ff] text-[#2563eb] text-[14px] hover:bg-[#2563eb] hover:text-white active:bg-[#1d4ed8] active:text-white focus-visible:outline-[#2563eb] ${actionButtonSharedClass}`
                      : `h-8 min-w-[92px] px-3 rounded-md border border-[#2563eb] bg-[#eff6ff] text-[#2563eb] text-[12px] hover:bg-[#2563eb] hover:text-white active:bg-[#1d4ed8] active:text-white focus-visible:outline-[#2563eb] ${actionButtonSharedClass}`}
                  >
                    Bypass
                  </button>
                )}

                {showReject && (
                  <button
                    onClick={onReject}
                    disabled={actionBusy}
                    className={isDesktop
                      ? `h-11 min-w-[130px] px-4 rounded-md border border-[#ef4444] bg-[#fff1f2] text-[#dc2626] text-[14px] hover:bg-[#dc2626] hover:text-white active:bg-[#b91c1c] active:text-white focus-visible:outline-[#ef4444] ${actionButtonSharedClass}`
                      : `h-8 min-w-[92px] px-3 rounded-md border border-[#ef4444] bg-[#fff1f2] text-[#dc2626] text-[12px] hover:bg-[#dc2626] hover:text-white active:bg-[#b91c1c] active:text-white focus-visible:outline-[#ef4444] ${actionButtonSharedClass}`}
                  >
                    Reject
                  </button>
                )}
              </>
            )}

          </div>
        </div>
      </div>
    </div>
  );
}
