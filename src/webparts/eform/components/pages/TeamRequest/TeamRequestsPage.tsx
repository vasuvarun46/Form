/* =========================================================
   TeamRequestsPage.tsx
   ✅ ApproverNames column added
   ✅ Multiple approver pending text → "Awaiting from the first approver"
   ✅ Save-and-Approve edit qty popup removed
   ✅ No other styling or logic removed
========================================================= */

import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import TeamRequestHeader from "../../Common/PageHeader/TeamRequestHeader";
import RequestFilterCard from "../../Common/Filters/RequestFilters";
import SelectBox from "../../Common/Filters/SelectBox";
import RequestSearchBar, { SearchFieldKey } from "../../Common/Search/RequestSearchBar";

import TeamRequestsTable from "./TeamRequestsTable";
import RequestEmptyState from "../RequestEmptyState";
import Loading from "../../ui/Loading";
import { exportToExcel, fetchExportItemsByRequestIds } from "../ExportExcel";

import {
  useTeamRequestsLogic,
  TeamRequestsPageProps,
  Row,
} from "./TeamService";
import NotificationBar from "./NotificationBar";

type PendingAction = "Approve" | "Reject";

function StatusText({ status }: { status: string }): JSX.Element {
  const key = String(status || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");

  const isInApproval =
    key === "pending" ||
    key === "in approval" ||
    key === "in-approval";

  const display =
    key === "submitted"
      ? "Submitted"
      : isInApproval
        ? "In-Approval"
        : key === "approved"
          ? "Approved"
          : key === "rejected"
            ? "Rejected"
            : key === "draft"
              ? "Draft"
              : status;

  const cls =
    key === "approved"
      ? "text-green-600"
      : key === "rejected"
        ? "text-red-600"
        : key === "draft"
          ? "text-gray-600"
          : key === "submitted" || isInApproval
            ? "text-blue-600"
            : "text-gray-500";

  return <span className={`font-medium ${cls}`}>{display}</span>;
}

function toUserMessage(value: unknown, fallback: string): string {
  if (typeof value === "string" && value.trim()) return value.trim();
  if (value instanceof Error) {
    const msg = String(value.message || "").trim();
    return msg || fallback;
  }
  if (value && typeof value === "object") {
    const obj = value as Record<string, unknown>;
    const nestedError = obj.error && typeof obj.error === "object"
      ? (obj.error as Record<string, unknown>)
      : undefined;
    const candidates: unknown[] = [
      obj.message, obj.error_description, obj.detail, obj.title,
      nestedError?.message, nestedError?.error_description, nestedError?.detail, nestedError?.title,
    ];
    for (const candidate of candidates) {
      if (typeof candidate === "string" && candidate.trim()) return candidate.trim();
    }
  }
  return fallback;
}

function normalizeToastMessage(message: string): string {
  return String(message || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/^[^\wA-Za-z]+/, "")
    .trim();
}

function CellText(props: { value?: string; maxWidth?: string }): JSX.Element {
  const raw = String(props.value || "").trim();
  const display = raw === "-" || raw === "" ? "" : raw;
  return (
    <div className={`${props.maxWidth || "max-w-[110px]"} truncate`} title={display}>
      {display}
    </div>
  );
}

export default function TeamRequestsPage(props: TeamRequestsPageProps): JSX.Element {
  const ctx: WebPartContext = props.context;
  const canExport = props.permissions.viewExportTeamReport && props.permissions.exportTeamReport;

  const [searchField, setSearchField] = React.useState<SearchFieldKey>("TransactionID");
  const [searchDraft, setSearchDraft] = React.useState<string>("");
  const [sortDir, setSortDir] = React.useState<"asc" | "desc">("desc");

  const [selectedIds, setSelectedIds] = React.useState<string[]>([]);
  const [manualSelectedIds, setManualSelectedIds] = React.useState<string[]>([]);

  const [toast, setToast] = React.useState<null | { type: "success" | "error"; message: string }>(null);

  React.useEffect(() => {
    if (!toast) return;
  }, [toast]);

  React.useEffect(() => {
    const incomingType = props.routeState?.toastType;
    if (!incomingType) return;
    const fallbackMessage =
      incomingType === "success" ? "Action completed successfully." : "Action failed.";
    if (incomingType === "success" || incomingType === "error") {
      setToast({
        type: incomingType,
        message: normalizeToastMessage(props.routeState?.toastMessage || fallbackMessage),
      });
    }
    props.onRouteStateConsumed?.();
  }, [props.routeState, props.onRouteStateConsumed]);

  // ✅ Action modal state (edit qty popup state removed)
  const [actionModalOpen, setActionModalOpen] = React.useState<boolean>(false);
  const [pendingRow, setPendingRow] = React.useState<Row | null>(null);
  const [pendingAction, setPendingAction] = React.useState<PendingAction>("Approve");
  const [actionComments, setActionComments] = React.useState<string>("");
  const [actionSubmitting, setActionSubmitting] = React.useState<boolean>(false);
  const [isAllSelected, setIsAllSelected] = React.useState(false);

  const {
    loading,
    tableLoading,
    error,
    filteredRows,
    pageSize,
    safePageIndex,
    totalPages,
    goFirst,
    goPrev,
    goNext,
    goLast,
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
    setSearchText,
    getActionInfo,
    // validateRequestItemsBeforeApprove,
    // updateRequestItemQuantity,
    submitDecision,
    reload,
    rows,
  } = useTeamRequestsLogic(ctx, props.ApprovalFlowUrl, searchField, props.permissions, props.productFlowUrl);

  React.useEffect(() => {
    if (!toast) return;
    setActionModalOpen(false);
    setPendingRow(null);
    setActionComments("");
    if (toast.type === "success") {
      reload().catch(() => undefined);
    }
  }, [toast, reload]);

  const [appliedSearchText, setAppliedSearchText] = React.useState<string>("");

  /* =========================================================
     SORT
  ========================================================== */
  const sortedRows = React.useMemo((): Row[] => {
    const actionPriority = (row: Row): number => (getActionInfo(row).canAct ? 0 : 1);
    const statusPriority = (status: string): number => {
      const key = String(status || "").trim().toLowerCase();
      return key === "submitted" || key === "pending" || key === "in approval" || key === "in-approval" ? 0 : 1;
    };
    const out = [...filteredRows];
    out.sort((a, b) => {
      const aAction = actionPriority(a);
      const bAction = actionPriority(b);
      if (aAction !== bAction) return aAction - bAction;
      const p1 = statusPriority(a.status);
      const p2 = statusPriority(b.status);
      if (p1 !== p2) return p1 - p2;
      const cmp = a.id.localeCompare(b.id, undefined, { numeric: true });
      return sortDir === "asc" ? cmp : -cmp;
    });
    return out;
  }, [filteredRows, sortDir, getActionInfo]);

  const hasActiveFilters =
    statusSelected.length > 0 ||
    purposeSelected.length > 0 ||
    costCenterSelected.length > 0 ||
    !!expectedFrom ||
    !!expectedTo;
  const hasAppliedFilters = hasActiveFilters || appliedSearchText.trim().length > 0;

  /* =========================================================
     SELECTION
  ========================================================== */
  const handleToggleRow = (id: string | number): void => {
    const idStr = String(id);
    setIsAllSelected(false);
    setSelectedIds((prev) => {
      const updated = prev.includes(idStr) ? prev.filter((x) => x !== idStr) : [...prev, idStr];
      setManualSelectedIds(updated);
      return updated;
    });
  };

  const handleToggleAll = (): void => {
    if (isAllSelected) {
      setIsAllSelected(false);
      setSelectedIds([]);
      setManualSelectedIds([]);
      return;
    }
    const pageIds = sortedRows
      .slice(safePageIndex * pageSize, safePageIndex * pageSize + pageSize)
      .map((r) => String(r.id));
    const allSelectedOnPage = pageIds.every((id) => selectedIds.includes(id));
    if (allSelectedOnPage) {
      const updated = selectedIds.filter((id) => !pageIds.includes(id));
      setSelectedIds(updated);
      setManualSelectedIds(updated);
    } else {
      const updated = [...selectedIds, ...pageIds.filter((id) => !selectedIds.includes(id))];
      setSelectedIds(updated);
      setManualSelectedIds(updated);
    }
  };

  const selectAllFilteredRows = (): void => {
    const ids = sortedRows.map((r) => String(r.id));
    setSelectedIds(ids);
    setManualSelectedIds(ids);
    setIsAllSelected(true);
  };

  const restoreManualSelection = (): void => {
    setIsAllSelected(false);
    setSelectedIds(manualSelectedIds);
  };

  /* =========================================================
     APPROVE/REJECT HANDLERS
     ✅ Edit qty popup removed — submitDecision now always
        proceeds directly, carrying quantityShort + skuMismatches
        to Power Automate via the service layer.
  ========================================================== */
  const openActionModal = (row: Row, action: PendingAction): void => {
    setPendingRow(row);
    setPendingAction(action);
    setActionComments("");
    setActionModalOpen(true);
  };

  const closeActionModal = (force?: boolean): void => {
    if (actionSubmitting && !force) return;
    setActionModalOpen(false);
    setPendingRow(null);
    setActionComments("");
  };

  const onApproveReject = (row: Row, action: PendingAction): void => {
    openActionModal(row, action);
  };

  const confirmApproveReject = async (): Promise<void> => {
    if (!pendingRow) return;

    if (pendingAction === "Reject" && !actionComments.trim()) {
      setToast({ type: "error", message: "Rejection reason is required." });
      return;
    }

    try {
      setActionSubmitting(true);

      // ✅ submitDecision now handles stock validation internally and
      //    sends quantityShort + skuMismatches to Power Automate.
      //    No edit popup is opened here anymore.
      const res = await submitDecision(pendingRow, pendingAction, actionComments.trim());

      if (res.success) {
        setToast({ type: "success", message: res.message || `${pendingAction} success` });
        closeActionModal(true);
      } else {
        setToast({ type: "error", message: res.message || `${pendingAction} failed` });
      }
    } catch (e: unknown) {
      setToast({ type: "error", message: toUserMessage(e, "Action failed") });
    } finally {
      setActionSubmitting(false);
    }
  };

  /* =========================================================
     COLUMNS
  ========================================================== */
  const columns = [
    {
      key: "id",
      header: "Transaction<br/>ID",
      thClassName: "w-[9%]",
      tdClassName: "w-[9%] truncate",
      render: (r: Row) =>
        r.id === "-" ? (
          <span className="text-gray-400"></span>
        ) : (
          <button
            className="text-[#0a66d8] underline truncate max-w-full block"
            onClick={() => props.onOpenExportPdf(r.spId, getActionInfo(r))}
          >
            {r.id}
          </button>
        ),
    },
    {
      key: "status",
      header: "Status",
      thClassName: "w-[6%]",
      tdClassName: "w-[6%]",
      render: (r: Row) => <StatusText status={r.status} />,
    },
    {
      key: "requestor",
      header: "Requestor<br/>Name",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%]",
      render: (r: Row) => (
        <CellText value={r.requestor === "-" ? "" : r.requestor} maxWidth="max-w-full" />
      ),
    },
    
    {
      key: "approverNames",
      header: "Approver<br/>Names",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%]",
      render: (r: Row) => (
        <CellText value={r.approverNames || ""} maxWidth="max-w-full" />
      ),
    },
    {
      key: "purpose",
      header: "Purpose",
      thClassName: "w-[6%]",
      tdClassName: "w-[6%]",
      render: (r: Row) => (
        <CellText value={r.purpose === "-" ? "" : r.purpose} maxWidth="max-w-full" />
      ),
    },
    {
      key: "purposedetails",
      header: "Purpose<br/>Detail",
      thClassName: "w-[6%]",
      tdClassName: "w-[6%]",
      render: (r: Row) => (
        <CellText value={r.purposeDetail === "-" ? "" : (r.purposeDetail || "")} maxWidth="max-w-full" />
      ),
    },
    {
      key: "costCenter",
      header: "Cost<br/>Center",
      thClassName: "w-[6%]",
      tdClassName: "w-[6%]",
      render: (r: Row) => (
        <CellText value={r.costCenter === "-" ? "" : r.costCenter} maxWidth="max-w-full" />
      ),
    },
    {
      key: "submitted",
      header: "Submitted<br/>Date",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%] whitespace-nowrap",
      render: (r: Row) => (r.submitted === "-" || !r.submitted ? "" : r.submitted),
    },
    {
      key: "approved",
      header: "Approved<br/>Date",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%] whitespace-nowrap",
      render: (r: Row) => (r.approved === "-" || !r.approved ? "" : r.approved),
    },
    {
      key: "expected",
      header: "Expected<br/>Delivery Date",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%] whitespace-nowrap",
      render: (r: Row) => (r.expected === "-" || !r.expected ? "" : r.expected),
    },
    {
      key: "qty",
      header: "Total<br/>Quantity",
      thClassName: "w-[4%] text-left",
      tdClassName: "w-[4%] whitespace-nowrap text-left",
      render: (r: Row) => r.qty
    },
    {
      key: "price",
      header: "Total Tag<br/>Price",
      thClassName: "w-[7%] text-left",
      tdClassName: "w-[7%] whitespace-nowrap text-left",
      render: (r: Row) =>
        r.price.toLocaleString(undefined, {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        }),
    },
    {
      key: "export",
      header: "Export<br/>Request",
      thClassName: "w-[7%] text-left",
      tdClassName: "w-[7%] text-left",
      render: (r: Row) => {
        if (!props.permissions.exportTeamReport || r.status === "Draft" || r.status === "Failed") {
          return <div className="text-center text-gray-500"></div>;
        }
        return (
          <div className="w-full flex justify-start">
            <button
              className="inline-flex w-full max-w-[58px] items-center justify-center px-1.5 py-1.5 rounded-md text-[#81AFEB] border border-[#EBEBEB] bg-[#F6F6F6] text-[#2563eb] text-[8px] font-semibold hover:bg-[#F0F7FF] hover:border hover:text-blue-400 hover:border-[#2172DA]"
              onClick={() => props.onDownloadPdf(r.spId)}
            >
              Export.PDF
            </button>
          </div>
        );
      },
    },
    {
      key: "approval",
      header: "Approval",
      thClassName: "w-[10%]",
      tdClassName: "w-[10%]",
      render: (r: Row) => {

        // ── 0. Permission guard ────────────────────────────────────────────────
        if (!props.permissions.approveTeam) {
          return <div className="text-center text-gray-500"></div>;
        }
        const normalizedStatus = String(r.status || "").trim().toLowerCase();

        // ── 1. Terminal states → always dash ──────────────────────────────────
        if (normalizedStatus === "approved" || normalizedStatus === "rejected") {
          return <div className="text-center text-gray-400">–</div>;
        }

        // ── 2. Only act on submitted / in-approval ────────────────────────────
        const isActionableStatus =
          normalizedStatus === "submitted" ||
          normalizedStatus === "in approval" ||
          normalizedStatus === "in-approval";

        if (!isActionableStatus) {
          return <div className="text-center text-gray-400">–</div>;
        }

        const loginEmail = String(
          props.context?.pageContext?.user?.email || ""
        ).trim().toLowerCase();

        const isMultipleApprover =
          String(r.approvalType || "").trim().toLowerCase() === "multipleapprover";

        // ── 3. Non-MultipleApprover: original behaviour ───────────────────────
        if (!isMultipleApprover) {
          const info = getActionInfo(r);
          if (info.canAct) {
            return (
              <div className="w-full flex items-center justify-center gap-0.5 flex-wrap">
                <button
                  className="inline-flex flex-1 min-w-[40px] items-center justify-center px-1.5 py-1.5 rounded-md border border-[#12a020] bg-[#e6f2e8] text-[#12a020] text-[8px] font-semibold hover:bg-[#12a020] hover:text-white"
                  onClick={() => onApproveReject(r, "Approve")}
                >
                  Approve
                </button>
                {info.showReject && (
                  <button
                    className="inline-flex flex-1 min-w-[40px] items-center justify-center px-1.5 py-1.5 rounded-md border border-[#ef4444] bg-[#fff1f2] text-[#dc2626] text-[8px] font-semibold hover:bg-[#dc2626] hover:text-white"
                    onClick={() => onApproveReject(r, "Reject")}
                  >
                    Reject
                  </button>
                )}
              </div>
            );
          }
          return <div className="text-center text-gray-400">–</div>;
        }

        // ── 4. MultipleApprover: parse stages from ApprovalChainJSON ──────────
        //
        // Your JSON shape (from r.allApprovalLevels which is already parsed in Row):
        //   stages: [{ stage: 1, approverEmail: "...", ... }, { stage: 2, ... }]
        //
        // We use r.allApprovalLevels (already mapped by TeamService) as the
        // source of truth for canApproveRequest, plus r.currentApproverEmail
        // from the SP column as the "whose turn is it now" pointer.

        const currentApproverEmail = String(r.currentApproverEmail || "")
          .trim()
          .toLowerCase();

        // ── Parse raw stages from ApprovalChainJSON stored on the Row ─────────
        // TeamService maps MultipleApproverData → r.allApprovalLevels already,
        // but we need the ordered stage list specifically. We read it from the
        // raw JSON that is stored in the SP column (accessed via r directly).
        // If not available, fall back to allApprovalLevels sorted by order.

        type StageEntry = { stage: number; approverEmail: string };
        let parsedStages: StageEntry[] = [];

        try {
          // ApprovalChainJSON is mapped into r as approvalStages in TeamService
          // Use approvalStages if populated, otherwise derive from allApprovalLevels
          if (r.approvalStages && r.approvalStages.length > 0) {
            // approvalStages in Row only has approverId — we need email
            // So derive from allApprovalLevels which has email + canApproveRequest
            parsedStages = r.allApprovalLevels
              .filter((l) => l.canApproveRequest === true)
              .sort((a, b) => a.order - b.order)
              .map((l, idx) => ({
                stage: idx + 1,
                approverEmail: String(l.approverEmail || "").trim().toLowerCase(),
              }));
          }

          // If approvalStages is empty, still derive from allApprovalLevels
          if (parsedStages.length === 0) {
            parsedStages = r.allApprovalLevels
              .filter((l) => l.canApproveRequest === true)
              .sort((a, b) => a.order - b.order)
              .map((l, idx) => ({
                stage: idx + 1,
                approverEmail: String(l.approverEmail || "").trim().toLowerCase(),
              }));
          }
        } catch {
          parsedStages = [];
        }

        // ── Stage 1 and Stage 2 emails ────────────────────────────────────────
        const stage1Email = parsedStages.find((s) => s.stage === 1)?.approverEmail ?? "";
        const stage2Email = parsedStages.find((s) => s.stage === 2)?.approverEmail ?? "";

        const loginIsStage1 = !!stage1Email && stage1Email === loginEmail;
        const loginIsStage2 = !!stage2Email && stage2Email === loginEmail;
        const loginIsCurrentApprover = !!currentApproverEmail && currentApproverEmail === loginEmail;

        const info = getActionInfo(r);

        // ── Case A: Login user is Stage 1 ─────────────────────────────────────
        if (loginIsStage1) {

          // A1: Stage 1 is current approver → their turn → Approve + Reject
          if (loginIsCurrentApprover) {
            return (
              <div className="w-full flex items-center justify-center gap-0.5 flex-wrap">
                <button
                  className="inline-flex flex-1 min-w-[40px] items-center justify-center px-1.5 py-1.5 rounded-md border border-[#12a020] bg-[#e6f2e8] text-[#12a020] text-[8px] font-semibold hover:bg-[#12a020] hover:text-white"
                  onClick={() => onApproveReject(r, "Approve")}
                >
                  Approve
                </button>
                {info.showReject && (
                  <button
                    className="inline-flex flex-1 min-w-[40px] items-center justify-center px-1.5 py-1.5 rounded-md border border-[#ef4444] bg-[#fff1f2] text-[#dc2626] text-[8px] font-semibold hover:bg-[#dc2626] hover:text-white"
                    onClick={() => onApproveReject(r, "Reject")}
                  >
                    Reject
                  </button>
                )}
              </div>
            );
          }

          // A2: Stage 1 already approved (current approver pointer moved away)
          //     → waiting for stage 2 (final approver) to act
          return (
            <div className="text-center text-[11px] text-gray-500 leading-tight px-1">
              Pending approval from the final approver.
            </div>
          );
        }

        // ── Case B: Login user is Stage 2 ─────────────────────────────────────
        if (loginIsStage2) {

          // B1: Stage 2 is current approver → stage 1 already done → their turn
          if (loginIsCurrentApprover) {
            return (
              <div className="w-full flex items-center justify-center gap-0.5 flex-wrap">
                <button
                  className="inline-flex flex-1 min-w-[40px] items-center justify-center px-1.5 py-1.5 rounded-md border border-[#12a020] bg-[#e6f2e8] text-[#12a020] text-[8px] font-semibold hover:bg-[#12a020] hover:text-white"
                  onClick={() => onApproveReject(r, "Approve")}
                >
                  Approve
                </button>
                {info.showReject && (
                  <button
                    className="inline-flex flex-1 min-w-[40px] items-center justify-center px-1.5 py-1.5 rounded-md border border-[#ef4444] bg-[#fff1f2] text-[#dc2626] text-[8px] font-semibold hover:bg-[#dc2626] hover:text-white"
                    onClick={() => onApproveReject(r, "Reject")}
                  >
                    Reject
                  </button>
                )}
              </div>
            );
          }

          // B2: Stage 2 is not yet the current approver
          //     → stage 1 hasn't acted yet → waiting
          return (
            <div className="text-center text-[11px] text-gray-500 leading-tight px-1">
              Awaiting from the first approver.
            </div>
          );
        }

        // ── Case C: Login user is not in stage 1 or stage 2 ──────────────────
        return <div className="text-center text-gray-400">–</div>;
      },
    },
  ];

  const card = "bg-white rounded-xl border border-[#e5e5e5] shadow-sm";
  const rejectReasonEmpty = pendingAction === "Reject" && !actionComments.trim();
  const confirmDisabled = actionSubmitting || rejectReasonEmpty;
  const isRejectModal = pendingAction === "Reject";

  /* =========================================================
     UI
  ========================================================== */

  return (
    <>
      {loading && <Loading />}

      {!loading && (
        <div className="pdf-thai flex flex-col h-full min-h-0 overflow-hidden">
          {/* HEADER */}
          <div className="relative sticky top-0 z-20 bg-white shrink-0">
            {toast && (
              <div className="absolute top-3 left-0 right-0 z-[40] px-7 sm:px-6">
                <NotificationBar
                  type={toast.type}
                  message={toast.message}
                  onClose={() => setToast(null)}
                  inline={true}
                />
              </div>
            )}
            <TeamRequestHeader
              loading={loading || (rows.length === 0 && !hasAppliedFilters)}
              selectedCount={selectedIds.length}
              onSelectAllRequests={selectAllFilteredRows}
              manualSelectedCount={manualSelectedIds.length}
              onRestoreManualSelection={restoreManualSelection}
              onExportExcel={async (mode) => {
                if (!canExport) return;
                let data: Row[] = [];
                if (mode === "all") {
                  data = sortedRows;
                } else {
                  data = sortedRows.filter((r) => selectedIds.indexOf(String(r.id)) !== -1);
                }
                const itemsByRequestId = await fetchExportItemsByRequestIds(ctx, data.map((r) => r.spId));
                const exportRows = data.map((r) => ({ ...r, items: itemsByRequestId[r.spId] || [] }));
                await exportToExcel(exportRows);
              }}
              canExport={canExport}
              showDisabledExport={props.permissions.viewExportTeamReport && rows.length === 0 && !hasAppliedFilters}
            />
          </div>

          {/* BODY */}
          <div className="flex-1 overflow-y-auto overflow-x-hidden bg-[#f3f3f3] py-4 px-5 sm:px-5">
            <div className="space-y-6">
              {rows.length === 0 && !hasAppliedFilters ? (
                <RequestEmptyState
                  title="Empty team's requests"
                  description="There is no any submitted requests from the team yet"
                  showAction={false}
                />
              ) : (
                <>
                  <RequestFilterCard
                    statusOptions={statusOptions}
                    purposeOptions={purposeOptions}
                    costCenterOptions={costCenterOptions}
                    statusSelected={statusSelected}
                    purposeSelected={purposeSelected}
                    costCenterSelected={costCenterSelected}
                    expectedFrom={expectedFrom}
                    expectedTo={expectedTo}
                    setStatusSelected={setStatusSelected}
                    setPurposeSelected={setPurposeSelected}
                    setCostCenterSelected={setCostCenterSelected}
                    setExpectedFrom={setExpectedFrom}
                    setExpectedTo={setExpectedTo}
                    onClearAll={() => {
                      setStatusSelected([]);
                      setPurposeSelected([]);
                      setCostCenterSelected([]);
                      setExpectedFrom("");
                      setExpectedTo("");
                    }}
                    showClearAll={hasActiveFilters}
                    SelectBox={SelectBox}
                  />
                  <div className="text-[13px] font-semibold text-gray-900">Request lists</div>

                  <div className={`${card} overflow-hidden`}>
                    <div className="px-4 py-3 flex items-center justify-between border-b border-[#eeeeee]">
                      <RequestSearchBar
                        searchDraft={searchDraft}
                        setSearchDraft={setSearchDraft}
                        onSearch={() => {
                          setSearchText(searchDraft);
                          setAppliedSearchText(searchDraft);
                        }}
                        onSearchWithField={({ field, text }) => {
                          setSearchField(field);
                          setSearchText(text);
                          setAppliedSearchText(text);
                        }}
                      />
                      <button
                        onClick={() => setSortDir((d) => (d === "asc" ? "desc" : "asc"))}
                        className="inline-flex h-7 items-center gap-1.5 rounded-md border border-gray-200 bg-gray-100 px-2.5 text-[12px] font-medium text-gray-800 transition-colors duration-200 hover:bg-gray-200"
                      >
                        <svg
                          xmlns="http://www.w3.org/2000/svg"
                          className="h-3.5 w-3.5 text-gray-500"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                        >
                          <path strokeLinecap="round" strokeLinejoin="round" d="M8 7l4-4 4 4M16 17l-4 4-4-4" />
                        </svg>
                        Sort
                      </button>
                    </div>

                    <TeamRequestsTable
                      columns={columns}
                      pageRows={sortedRows.slice(safePageIndex * pageSize, safePageIndex * pageSize + pageSize)}
                      filteredRows={sortedRows}
                      loading={tableLoading}
                      error={error}
                      pageSize={pageSize}
                      safePageIndex={safePageIndex}
                      totalPages={totalPages}
                      goFirst={goFirst}
                      goPrev={goPrev}
                      goNext={goNext}
                      goLast={goLast}
                      selectedIds={selectedIds}
                      onToggleRow={handleToggleRow}
                      onToggleAll={handleToggleAll}
                      canExport={canExport}
                    />
                  </div>
                </>
              )}
            </div>
          </div>

          {/* ✅ Edit qty popup REMOVED entirely */}

          {/* APPROVE / REJECT MODAL — unchanged */}
          {actionModalOpen && (
            <div className="fixed inset-0 z-[999] flex items-center justify-center">
              <div className="absolute inset-0 bg-black/30" onClick={() => closeActionModal()} />
              <div
                className={
                  isRejectModal
                    ? "relative w-[92%] max-w-[460px] rounded-[16px] border border-[#d9d9d9] bg-[#f7f7f7] px-6 pb-6 pt-5 shadow-[0_16px_36px_rgba(0,0,0,0.22)]"
                    : "relative w-[92%] max-w-[430px] rounded-[24px] border border-[#ebebeb] bg-white px-9 pb-8 pt-6 shadow-[0_20px_45px_rgba(0,0,0,0.18)]"
                }
              >
                <button
                  className={
                    isRejectModal
                      ? "absolute right-4 top-4 text-[#b9b9b9] hover:text-[#9a9a9a] disabled:opacity-50"
                      : "absolute right-5 top-5 text-[#bcbcbc] hover:text-[#9a9a9a] disabled:opacity-50"
                  }
                  onClick={() => closeActionModal()}
                  disabled={actionSubmitting}
                  aria-label="Close"
                >
                  <svg className={isRejectModal ? "h-6 w-6" : "h-7 w-7"} viewBox="0 0 24 24" fill="none" aria-hidden="true">
                    <path d="M6 6L18 18M18 6L6 18" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
                  </svg>
                </button>

                <h2 className={isRejectModal ? "mt-2 text-center text-[20px] leading-tight font-semibold text-black" : "mt-8 text-center text-[17px] leading-tight font-semibold text-black"}>
                  {pendingAction === "Approve"
                    ? "Do you want to approve the request?"
                    : "Do you want to reject the request?"}
                </h2>

                <p className={isRejectModal ? "mt-2 text-center text-[13px] font-semibold leading-tight text-black" : "mt-3 text-center text-[18px] font-semibold leading-tight text-black"}>
                  ID &ldquo;{pendingRow?.id ?? "-"}&rdquo;
                </p>

                {pendingAction === "Approve" ? (
                  <p className="mx-auto mt-4 max-w-[560px] text-center text-[14px] leading-[1.45] text-[#5f5f5f]">
                    Once the request is approved, it cannot be edited or rolled back. The system will verify required quantity against
                    available stock. Only fully available items will be processed. Approver(s), requestor, and admin will be notified
                    for unavailable items by email which must be resubmitted as a new request. Please confirm to proceed.
                  </p>
                ) : (
                  <div className="mt-4">
                    <p className="mx-auto mb-3 max-w-[430px] text-center text-[11px] leading-[1.45] text-[#6a6a6a]">
                      The request cannot be edited or rolled-back once it is rejected. Please provide reason for rejection and confirm to proceed.
                    </p>
                    <div className="relative mt-2">
                      <label className="absolute left-3 top-2 text-[11px] font-semibold text-black pointer-events-none z-10">
                        *Require reject reason ({actionComments.length}/100)
                      </label>
                      <textarea
                        className="min-h-[130px] w-full rounded-[3px] border border-[#3b82f6] bg-white px-3 pb-2 pt-7 text-[12px] text-black outline-none focus:outline-none focus-visible:outline-none resize-none"
                        placeholder=""
                        value={actionComments}
                        onChange={(e) => setActionComments(e.target.value)}
                        disabled={actionSubmitting}
                        maxLength={100}
                      />
                    </div>
                  </div>
                )}

                <div className={isRejectModal ? "mt-6 flex items-center justify-center gap-4" : "mt-7 grid grid-cols-2 gap-4"}>
                  <button
                    className={
                      isRejectModal
                        ? "h-[38px] w-[96px] rounded-[6px] border border-[#d7d7d7] bg-[#e4e4e4] text-[16px] leading-none font-medium text-[#474747] hover:bg-[#d8d8d8] disabled:opacity-50"
                        : "h-[46px] rounded-[10px] border border-[#ececec] bg-[#ececec] text-[16px] leading-none font-semibold text-[#494949] hover:bg-[#dfdfdf] disabled:opacity-50"
                    }
                    onClick={() => closeActionModal()}
                    disabled={actionSubmitting}
                  >
                    Cancel
                  </button>

                  <button
                    className={
                      pendingAction !== "Reject"
                        ? "h-[46px] rounded-[10px] bg-[#12a020] text-[16px] leading-none font-semibold text-white hover:bg-[#0f8a1b] disabled:opacity-60"
                        : confirmDisabled
                          ? isRejectModal
                            ? "h-[38px] w-[96px] rounded-[6px] bg-[#d9d9d9] text-[16px] leading-none font-semibold text-[#8f8f8f] cursor-not-allowed"
                            : "h-[46px] rounded-[10px] bg-[#d9d9d9] text-[16px] leading-none font-semibold text-[#8f8f8f] cursor-not-allowed"
                          : isRejectModal
                            ? "h-[38px] w-[96px] rounded-[6px] bg-[#df1f26] text-[16px] leading-none font-semibold text-white hover:bg-[#c71920]"
                            : "h-[46px] rounded-[10px] bg-[#dc2626] text-[16px] leading-none font-semibold text-white hover:bg-[#b91c1c]"
                    }
                    onClick={confirmApproveReject}
                    disabled={confirmDisabled}
                  >
                    {actionSubmitting
                      ? pendingAction === "Approve"
                        ? "Approving..."
                        : "Rejecting..."
                      : pendingAction}
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
      )}
    </>
  );
}