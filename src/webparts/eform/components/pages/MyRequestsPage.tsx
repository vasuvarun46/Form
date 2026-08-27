import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import MyRequestHeader from "../Common/PageHeader/MyRequestsHeader";
import RequestFilterCard from "../Common/Filters/RequestFilters";
import SelectBox from "../Common/Filters/SelectBox";
import RequestSearchBar from "../Common/Search/RequestSearchBar";
import { ResponsiveTable } from "../Common/ResponsiveTable";
import Loading from "../ui/Loading";
import RequestEmptyState from "./RequestEmptyState";
import { exportToExcel, fetchExportItemsByRequestIds } from "./ExportExcel";
import {
  canonicalizeStatusFilterKey,
  dedupeLabels,
} from "./AllRequestPage/allRequestFilterUtils";
import { Icon } from "@fluentui/react";
import { RouteState, SecurityPermissions, ToastType } from "../spHelpers";


/* ================= TYPES ================= */

type Status = "Submitted" | "In Approval" | "Approved" | "Rejected" | "Draft" | "Pending";

interface MyRequestsPageProps {
  routeState?: RouteState;
  onRouteStateConsumed?: () => void;
  context: WebPartContext;
  onOpenExportPdf: (spId: number) => void;
  onDownloadPdf: (spId: number) => void;
  onOpenDraft: (draftId: number) => void;
  onCreate: () => void;
  permissions: SecurityPermissions;
}

interface Row {
  id: string;
  status: Status;
  requestor: string;
  approver: string;
  department: string;
  purpose: string;
  purposeDetail: string;
  costCenter: string;
  submitted: string;
  approved: string;
  approvedDates?: string;
  expected: string;
  qty: number;
  price: number;
  spId: number;
  expectedIso?: string;
  submittedIso?: string;
  requestCreatedIso?: string;
  modifiedIso?: string;
  lastActivityIso?: string;
  ApproverNames?: string;
  multipleApproverData?: string;
}

interface SPListResponse<T> {
  value: T[];
}

interface SPPerson {
  Id: number;
  Title: string;
  EMail: string;
}

interface InventoryRequestSPItem {
  Id: number;
  Title: string;
  RequestNumber?: number;
  Requester?: SPPerson;
  CurrentApprover?: { Title: string };
  Purpose?: { Title: string };
  CostCenter?: { Title: string };
  Departments?: string;
  PurposeDetails?: string;
  RequestSubmittedOn?: string;
  FinalDecisionDate?: string;
  ApprovedDates?: string;
  ExpectedDeliveryDate?: string;
  TotalQuantity?: number;
  TotalAmount?: number;
  WorkflowStatus?: Status;
  RequestCreatedOn?: string;
  Modified?: string;
  ApproverNames?: string;
  ApprovalChainJSON?: string;
  MultipleApproverData?: string;
}

/* ================= HELPERS ================= */

function StatusPill({ status }: { status: Status }): JSX.Element {
  const key = String(status || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");

  const isInApproval =
    key === "in approval" ||
    key === "in-approval";

  const isPending = key === "pending";

  const display =
    key === "submitted"
      ? "Submitted"
      : isPending
        ? "Pending"
        : isInApproval
          ? "Submitted"
          : key === "approved"
            ? "Approved"
            : key === "rejected"
              ? "Rejected"
              : key === "draft"
                ? "Draft"
                : status;

  const colorMap: Record<string, string> = {
    approved: "text-green-600",
    rejected: "text-red-600",
    pending: "text-[#C8C8C8]",
    submitted: "text-blue-600",
    draft: "text-gray-600",
    "in approval": "text-blue-600",
    "in-approval": "text-blue-600"
  };

  const cls = colorMap[key] || "text-gray-600";

  return <span className={`font-medium ${cls}`}>{display}</span>;
}

function normalizeToastMessage(message: string): string {
  return String(message || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/^[^\wA-Za-z]+/, "")
    .trim();
}

function CellText(props: { value: string; maxWidth?: string }): JSX.Element {
  const text = String(props.value || "").trim();
  const display = text === "-" || text === "" ? "" : text;
  return (
    <div className={`${props.maxWidth || "max-w-[110px]"} truncate`} title={display}>
      {display}
    </div>
  );
}

/* ================= MAIN ================= */

export default function MyRequestsPage(props: MyRequestsPageProps): JSX.Element {
  const { context } = props;
  const canExport = props.permissions.viewExportOwnReport && props.permissions.exportOwnReport;

  const card = "bg-white rounded-xl border border-[#e5e5e5] shadow-sm";

  const [rows, setRows] = React.useState<Row[]>([]);
  const [allRows, setAllRows] = React.useState<Row[]>([]);

  const [loading, setLoading] = React.useState(false);
  const [tableLoading, setTableLoading] = React.useState(false);
  const [hasLoadedOnce, setHasLoadedOnce] = React.useState(false);
  const [, setError] = React.useState("");

  // filters
  const [statusSelected, setStatusSelected] = React.useState<string[]>([]);
  const [purposeSelected, setPurposeSelected] = React.useState<string[]>([]);
  const [costCenterSelected, setCostCenterSelected] = React.useState<string[]>([]);
  const [expectedFrom, setExpectedFrom] = React.useState("");
  const [expectedTo, setExpectedTo] = React.useState("");
  const [selectedIds, setSelectedIds] = React.useState<string[]>([]);
  const [manualSelectedIds, setManualSelectedIds] = React.useState<string[]>([]);

  const [searchDraft, setSearchDraft] = React.useState("");
  const [searchText, setSearchText] = React.useState("");


  const [toast, setToast] = React.useState<null | { type: ToastType; message: string }>(null);
  const [pageIndex, setPageIndex] = React.useState(0);
  const [searchField, setSearchField] = React.useState<"TransactionID" | "RequesterName">("TransactionID");
  // Storage keys — shared constants
  const STORAGE_KEY_APPROVED = "myreq_approved_count";
  const STORAGE_KEY_REJECTED = "myreq_rejected_count";
  const STORAGE_KEY_DRAFT = "myreq_draft_count";
  const STORAGE_KEY_FAILED = "myreq_failed_count";



  // Tracks if approved/rejected count increased → triggers re-render
  const [countIncreased, setCountIncreased] = React.useState<boolean>(false);
  const pageSize = 15;

  const handleToggleRow = (id: string | number): void => {
    const idStr = String(id);
    setSelectedIds((prev) => {
      const updated = prev.includes(idStr)
        ? prev.filter((x) => x !== idStr)
        : [...prev, idStr];
      setManualSelectedIds(updated);
      return updated;
    });
  };

  /**
 * Format an ISO UTC date string to dd/MM/yyyy in SharePoint's regional timezone.
 * Returns "-" for empty / invalid input.
 */
  /**
 * Mirrors CreateRequest's isoToYyyyMmDdSp, but returns dd/MM/yyyy for display.
 * Same math — guarantees MyRequest displays the same date that CreateRequest saved.
 */
  function formatDateDdMmYyyySp(iso: string | undefined): string {
    if (!iso) return "-";
    const d = new Date(iso);
    if (isNaN(d.getTime())) return "-";
    return new Intl.DateTimeFormat("en-GB", {
      timeZone: "Asia/Bangkok",
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    }).format(d).replace(/\//g, "-"); // Returns "DD-MM-YYYY" in Bangkok time
  }

  /* ================= FILTERED (API-SIDE) ================= */

  const filteredRows = React.useMemo(() => {
    let data = [...allRows];

    // Status filter
    if (statusSelected.length > 0) {
      const normalized = statusSelected.map(s => s.toLowerCase());
      data = data.filter(r =>
        normalized.includes(String(r.status).toLowerCase())
      );
    }

    // Purpose filter
    if (purposeSelected.length > 0) {
      data = data.filter(r =>
        purposeSelected.includes(r.purpose)
      );
    }

    // Cost Center filter
    if (costCenterSelected.length > 0) {
      data = data.filter(r =>
        costCenterSelected.includes(r.costCenter)
      );
    }

    // Date filter
    if (expectedFrom) {
      data = data.filter(r =>
        r.expectedIso && new Date(r.expectedIso) >= new Date(expectedFrom)
      );
    }

    if (expectedTo) {
      data = data.filter(r =>
        r.expectedIso && new Date(r.expectedIso) <= new Date(expectedTo)
      );
    }

    // Search
    if (searchText.trim()) {
      const q = searchText.toLowerCase();
      if (searchField === "TransactionID") {
        data = data.filter(r =>
          String(r.id).toLowerCase().includes(q)
        );
      } else {
        data = data.filter(r =>
          r.requestor.toLowerCase().includes(q)
        );
      }
    }

    return data;

  }, [
    allRows,
    statusSelected,
    purposeSelected,
    costCenterSelected,
    expectedFrom,
    expectedTo,
    searchText,
    searchField,
  ]);

  const sortedRows: Row[] = React.useMemo(() => {
    const isPending = (r: Row): boolean => r.status.trim().toLowerCase() === "pending";
    const isDraft = (r: Row): boolean => r.status.trim().toLowerCase() === "draft";
    const pending = filteredRows.filter(isPending);
    const drafts = filteredRows.filter(isDraft);
    const rest = filteredRows.filter(r => !isPending(r) && !isDraft(r));
    const sortedRest = [...rest].sort((a, b) => String(b.id).localeCompare(String(a.id)));
    return [...pending, ...drafts, ...sortedRest];
  }, [filteredRows]);

  React.useEffect(() => {
    setPageIndex(0);
  }, [
    statusSelected,
    purposeSelected,
    costCenterSelected,
    expectedFrom,
    expectedTo,
    searchText,
    searchField,
  ]);


  const totalPages = Math.max(1, Math.ceil(sortedRows.length / pageSize));
  const safePageIndex = Math.min(Math.max(pageIndex, 0), totalPages - 1);
  const pageRows = sortedRows.slice(
    safePageIndex * pageSize,
    safePageIndex * pageSize + pageSize
  );

  const selectAllFilteredRows = (): void => {
    const ids = sortedRows.map((r) => String(r.id));
    setSelectedIds(ids);
    setManualSelectedIds(ids);
  };

  const restoreManualSelection = (): void => {
    setSelectedIds(manualSelectedIds);
  };

  const handleToggleAll = (): void => {
    const allIds = sortedRows.map((r) => String(r.id));
    const everythingSelected =
      allIds.length > 0 && allIds.every((id) => selectedIds.includes(id));
    if (everythingSelected) {
      setSelectedIds([]);
      setManualSelectedIds([]);
    } else {
      setSelectedIds(allIds);
      setManualSelectedIds(allIds);
    }
  };

  React.useEffect(() => {
    if (!toast) return;
  }, [toast]);

  React.useEffect(() => {
    const incomingType = props.routeState?.toastType;
    if (!incomingType) return;

    const fallbackMessage =
      incomingType === "success"
        ? "Your request is submitted to the approver(s)."
        : incomingType === "pending"
          ? "Your request status will be shown as Pending while it is being processed."
          : incomingType === "draft"
            ? "Your request is saved as draft."
            : "Your request is discarded.";

    setToast({
      type: incomingType,
      message: normalizeToastMessage(props.routeState?.toastMessage || fallbackMessage),
    });

    props.onRouteStateConsumed?.();
  }, [props.routeState, props.onRouteStateConsumed]);

  const spGet = async <T,>(url: string): Promise<T> => {
    const res: SPHttpClientResponse = await context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: "application/json;odata=nometadata", "odata-version": "" } }
    );
    if (!res.ok) throw new Error(await res.text());
    return (await res.json()) as T;
  };

  /* ================= LOAD ================= */
  const loadAll = React.useCallback(async (silent: boolean = false): Promise<void> => {
    if (!silent) {
      if (!hasLoadedOnce) {
        setLoading(true);
      } else {
        setTableLoading(true);
      }
    }

    setError("");
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const myEmail = String(context.pageContext.user.email || "").trim();
      const esc = (v: string): string => String(v || "").replace(/'/g, "''");

      // Always show only the logged-in user's own requests
      const clauses: string[] = [
        `Requester/EMail eq '${esc(myEmail)}'`,
        `WorkflowStatus ne 'Failed'`,
      ];

      if (statusSelected.length > 0) {
        const statusTerms: string[] = [];
        statusSelected.forEach((status) => {
          const key = canonicalizeStatusFilterKey(status);
          if (key === canonicalizeStatusFilterKey("Pending")) {
            statusTerms.push(`WorkflowStatus eq 'Pending'`);
          } else if (key === canonicalizeStatusFilterKey("In Approval")) {
            statusTerms.push(`WorkflowStatus eq 'In Approval'`);
          } else if (key === canonicalizeStatusFilterKey("Submitted")) {
            statusTerms.push(`WorkflowStatus eq 'Submitted'`);
          } else if (key === canonicalizeStatusFilterKey("Approved")) {
            statusTerms.push(`WorkflowStatus eq 'Approved'`);
          } else if (key === canonicalizeStatusFilterKey("Rejected")) {
            statusTerms.push(`WorkflowStatus eq 'Rejected'`);
          } else if (key === canonicalizeStatusFilterKey("Draft")) {
            statusTerms.push(`WorkflowStatus eq 'Draft'`);
          }
        });
        const uniqueTerms = Array.from(new Set(statusTerms));
        if (uniqueTerms.length > 0) {
          clauses.push(`(${uniqueTerms.join(" or ")})`);
        }
      }

      if (purposeSelected.length > 0) {
        const purposeTerms = purposeSelected.map(
          (p) => `Purpose/Title eq '${esc(p)}'`
        );
        clauses.push(`(${purposeTerms.join(" or ")})`);
      }

      if (costCenterSelected.length > 0) {
        const ccTerms = costCenterSelected.map(
          (c) => `CostCenter/Title eq '${esc(c)}'`
        );
        clauses.push(`(${ccTerms.join(" or ")})`);
      }

      if (expectedFrom) {
        clauses.push(`ExpectedDeliveryDate ge datetime'${expectedFrom}T00:00:00Z'`);
      }
      if (expectedTo) {
        clauses.push(`ExpectedDeliveryDate le datetime'${expectedTo}T23:59:59Z'`);
      }

      if (searchText.trim() && searchField === "RequesterName") {
        const q = esc(searchText.trim());
        clauses.push(`substringof('${q}',Requester/Title)`);
      }

      const url =
        `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items` +
        `?$top=5000&$select=Id,Title,RequestNumber,WorkflowStatus,TotalQuantity,TotalAmount,Modified,` +
        `PurposeDetails,RequestCreatedOn,RequestSubmittedOn,FinalDecisionDate,ApprovedDates,ExpectedDeliveryDate,ApproverNames,MultipleApproverData,` +
        `Requester/Title,Requester/EMail,CurrentApprover/Title,Purpose/Title,CostCenter/Title,Departments` +
        `&$expand=Requester,CurrentApprover,Purpose,CostCenter` +
        `&$filter=${clauses.join(" and ")}` +
        `&$orderby=Id desc`;

      const allItems: InventoryRequestSPItem[] = [];
      let nextUrl: string | undefined = url;
      while (nextUrl) {
        const json: SPListResponse<InventoryRequestSPItem> & { "@odata.nextLink"?: string } =
          await spGet(nextUrl);
        allItems.push(...(json.value || []));
        nextUrl = json["@odata.nextLink"];
      }

      const visibleItems = allItems;

      const mappedRows: Row[] = (visibleItems || []).map((it) => {
        const workflowStatus = String(it.WorkflowStatus || "Submitted");
        const isPending = workflowStatus.trim().toLowerCase() === "pending";
        const isDraft = workflowStatus.trim().toLowerCase() === "draft";
        const titleTxId = String(it.Title || "").trim();
        const hasFinalTitle =
          !!titleTxId &&
          !/^draft$/i.test(titleTxId) &&
          !/^temp$/i.test(titleTxId);
        const draftId =
          it.RequestNumber
            ? String(it.RequestNumber).padStart(4, "0").slice(-4)
            : String(it.Id).padStart(4, "0").slice(-4);

        return {
          id: isPending
            ? "-"
            : isDraft
              ? draftId
              : hasFinalTitle
                ? titleTxId
                : it.RequestNumber
                  ? String(it.RequestNumber)
                  : String(it.Id),
          status: workflowStatus as Status,
          requestor: it.Requester?.Title || "",
          approver: it.ApproverNames || "",
          multipleApproverData: it.MultipleApproverData,
          department: it.Departments || "",
          purpose: it.Purpose?.Title || "",
          purposeDetail: it.PurposeDetails || "",
          costCenter: it.CostCenter?.Title || "",
          submitted: formatDateDdMmYyyySp(it.RequestSubmittedOn),
          approved: formatDateDdMmYyyySp(it.FinalDecisionDate),
          approvedDates: it.ApprovedDates ? it.ApprovedDates.split(",").map((d: string) => { const s = d.trim(); if (!s) return ""; const p = s.split("-"); if (p.length === 3 && p[0].length === 4) return `${p[2]}-${p[1]}-${p[0]}`; return s.replace(/\//g, "-"); }).filter(Boolean).join(", ") : "",
          expected: formatDateDdMmYyyySp(it.ExpectedDeliveryDate),
          qty: Number(it.TotalQuantity || 0),
          price: Number(it.TotalAmount || 0),
          spId: it.Id,
          requestCreatedIso: it.RequestCreatedOn || "",
          modifiedIso: it.Modified || "",
          lastActivityIso: it.Modified || it.RequestSubmittedOn || it.RequestCreatedOn || "",
          expectedIso: it.ExpectedDeliveryDate,
          submittedIso: it.RequestSubmittedOn || it.ExpectedDeliveryDate || "",
        };
      });

      // ✅ CHANGED: Deduplicate — keep only latest record per spId, then per transaction ID
      const seenSpIds = new Set<number>();
      const uniqueBySpId = mappedRows.filter((r) => {
        if (seenSpIds.has(r.spId)) return false;
        seenSpIds.add(r.spId);
        return true;
      });
      // For non-draft rows with same transaction ID, keep only highest spId (latest)
      const txMap = new Map<string, Row>();
      const drafts: Row[] = [];
      for (const r of uniqueBySpId) {
        if (r.status.trim().toLowerCase() === "draft" || r.id === "-") {
          drafts.push(r);
        } else {
          const existing = txMap.get(r.id);
          if (!existing || r.spId > existing.spId) {
            txMap.set(r.id, r);
          }
        }
      }
      const deduped = drafts.concat(Array.from(txMap.values()));

      setRows(deduped);
      setAllRows(deduped);


      setPageIndex(0);

    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Failed to load data");
    } finally {
      if (!silent) {

        setLoading(false);
        setTableLoading(false);
      }
      setHasLoadedOnce(true);
    }

  }, [
    context,
    hasLoadedOnce,
  ]);

  React.useEffect(() => {
    loadAll(false).catch(console.error);
  }, [loadAll]);


  const getApprovedRejectedCounts = React.useCallback(async (): Promise<{
    approved: number;
    rejected: number;
    draft: number;
    failed: number;
  }> => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const myEmail = String(context.pageContext.user.email || "").replace(/'/g, "''");

    const buildUrl = (status: string): string =>
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items` +
      `?$select=Id,Requester/EMail` +
      `&$filter=Requester/EMail eq '${myEmail}' and WorkflowStatus eq '${status}'` +
      `&$expand=Requester` +
      `&$top=5000`;

    const fetchCount = async (status: string): Promise<number> => {
      const res: SPHttpClientResponse = await context.spHttpClient.get(
        buildUrl(status),
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "odata-version": "",
          },
        }
      );
      if (!res.ok) throw new Error(`Failed to fetch ${status} count`);
      const data = await res.json();
      return (data?.value?.length ?? 0) as number;
    };

    const [approved, rejected, draft, failed] = await Promise.all([
      fetchCount("Approved"),
      fetchCount("Rejected"),
      fetchCount("Draft"),
      fetchCount("Failed"),
    ]);

    return { approved, rejected, draft, failed };
  }, [context]);

  /* ── 3. Compare counts, re-render + reload if increased ── */
  const checkAndUpdateCounts = React.useCallback(
    async (isInitialLoad: boolean): Promise<void> => {
      try {
        const { approved: currentApproved, rejected: currentRejected, draft: currentDraft, failed: currentFailed } =
          await getApprovedRejectedCounts();

        const prevApproved = Number(localStorage.getItem(STORAGE_KEY_APPROVED) ?? 0);
        const prevRejected = Number(localStorage.getItem(STORAGE_KEY_REJECTED) ?? 0);
        const prevDraft = Number(localStorage.getItem(STORAGE_KEY_DRAFT) ?? 0);
        const prevFailed = Number(localStorage.getItem(STORAGE_KEY_FAILED) ?? 0);

        const hasIncreased =
          !isInitialLoad &&
          (
            currentApproved > prevApproved ||
            currentRejected > prevRejected ||
            currentDraft > prevDraft ||
            currentFailed > prevFailed
          );

        if (hasIncreased) {
          setCountIncreased((prev) => !prev);
          await loadAll(true);
        }

        localStorage.setItem(STORAGE_KEY_APPROVED, String(currentApproved));
        localStorage.setItem(STORAGE_KEY_REJECTED, String(currentRejected));
        localStorage.setItem(STORAGE_KEY_DRAFT, String(currentDraft));
        localStorage.setItem(STORAGE_KEY_FAILED, String(currentFailed));

      } catch (err) {
        console.error("[MyRequestsPage] checkAndUpdateCounts error:", err);
      }
    },
    [getApprovedRejectedCounts, loadAll]
  );

  React.useEffect(() => {
    checkAndUpdateCounts(true).catch(console.error);  // initial — no re-render

    const intervalId = setInterval(() => {
      checkAndUpdateCounts(false).catch(console.error); // poll — re-renders if increased
    }, 2 * 60 * 1000); // 3 minutes

    return () => clearInterval(intervalId);
  }, [checkAndUpdateCounts]);


  /* ================= COLUMNS ================= */

  const columns = [
    {
      key: "id",
      header: "Transaction<br/>ID",
      thClassName: "w-[12%]",
      tdClassName: "w-[12%] truncate",
      render: (r: Row) =>
        r.id === "-" || r.id === " " || r.id === "" ? (
          <span></span>
        ) : r.status === "Draft" ? (
          <button
            className="text-[#0a66d8] underline truncate max-w-full block"
            onClick={() => props.onOpenDraft(r.spId)}
          >
            {r.id}
          </button>
        ) : (
          <button
            className="text-[#0a66d8] underline truncate max-w-full block"
            onClick={() => props.onOpenExportPdf(r.spId)}
          >
            {r.id}
          </button>
        ),
    },
    {
      key: "status",
      header: "Status",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%]",
      render: (r: Row) => <StatusPill status={r.status} />,
    },
    {
      key: "requestor",
      header: "Requestor<br/>Name",
      thClassName: "w-[8%]",
      tdClassName: "w-[8%]",
      render: (r: Row) => (
        <CellText value={r.requestor === "-" ? "" : r.requestor} maxWidth="max-w-full" />
      ),
    },
    {
      key: "approver",
      header: "Approver<br/>Name",
      thClassName: "w-[8%]",
      tdClassName: "w-[8%]",
      render: (r: Row) => (
        <CellText value={r.approver === "-" ? "" : (r.approver || "")} maxWidth="max-w-full" />
      ),
    },
    {
      key: "purpose",
      header: "Purpose",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%]",
      render: (r: Row) => (
        <CellText value={r.purpose === "-" ? "" : r.purpose} maxWidth="max-w-full" />
      ),
    },
    {
      key: "purposedetails",
      header: "Purpose<br/>Detail",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%]",
      render: (r: Row) => (
        <CellText value={r.purposeDetail === "-" ? "" : r.purposeDetail} maxWidth="max-w-full" />
      ),
    },
    {
      key: "costCenter",
      header: "Cost<br/>Center",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%]",
      render: (r: Row) => (
        <CellText value={r.costCenter === "-" ? "" : r.costCenter} maxWidth="max-w-full" />
      ),
    },
    {
      key: "submitted",
      header: "Submitted<br/>Date",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%] whitespace-nowrap",
      render: (r: Row) => (r.submitted === "-" ? "" : r.submitted),
    },
    {
      key: "approved",
      header: "Approved<br/>Date",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%]",
      render: (r: Row) => r.approvedDates || (r.approved === "-" ? "" : r.approved),
    },
    {
      key: "expected",
      header: "Expected<br/>Delivery Date",
      thClassName: "w-[7%]",
      tdClassName: "w-[7%] whitespace-nowrap",
      render: (r: Row) => (r.expected === "-" ? "" : r.expected),
    },
    {
      key: "qty",
      header: "Total<br/>Quantity",
      thClassName: "w-[5%] text-left",
      tdClassName: "w-[5%] whitespace-nowrap text-left",
      render: (r: Row) => r.status === "Pending" ? "" : r.qty.toLocaleString(),
    },
    {
      key: "price",
      header: "Total Tag<br/>Price",
      thClassName: "w-[7%] text-left",
      tdClassName: "w-[7%] whitespace-nowrap text-left",
      render: (r: Row) =>
        r.status === "Pending"
          ? ""
          : Number(r.price).toLocaleString(undefined, {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            }),
    },
    {
      key: "export",
      header: "Export<br/>Request",
      thClassName: "w-[9%] text-left",
      tdClassName: "w-[9%] text-left",
      render: (r: Row) => {
        if (!props.permissions.exportOwnReport) {
          return <div className="text-center text-gray-500"> </div>;
        }
        if (r.status === "Draft" || r.status === "Pending") {
          return <div className="text-center text-gray-500"></div>;
        }
        return (
          <div className="flex justify-start">
            <button
              className="w-full max-w-[80px] rounded-md bg-[#F6F6F6] px-1.5 py-1.5 text-[9px] font-semibold whitespace-nowrap text-[#81AFEB] border border-[#EBEBEB] hover:bg-[#F0F7FF] hover:border hover:border-[#2172DA] hover:text-blue-400"
              onClick={() => props.onDownloadPdf(r.spId)}
            >
              Export .PDF
            </button>
          </div>
        );
      },
    },
  ];

  // Load Purpose & Cost Center from master lists
  const [purposeOptions, setPurposeOptions] = React.useState<string[]>([]);

  const [costCenterOptions, setCostCenterOptions] = React.useState<string[]>([]);
  React.useEffect(() => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('CostCenterMaster')/items?$top=5000&$select=Title&$filter=IsActive eq 1&$orderby=Title asc`,
      SPHttpClient.configurations.v1
    ).then(r => r.json()).then(json => {
      const titles: string[] = (json?.value || []).map((x: { Title?: string }) => String(x?.Title || "")).filter(Boolean);
      setCostCenterOptions(dedupeLabels(titles).sort((a, b) => a.localeCompare(b)));
    }).catch(() => undefined);
  }, []);

  React.useEffect(() => {
    const siteUrl = context.pageContext.web.absoluteUrl;

    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('PurposeMaster')/items?$top=5000&$select=Title&$filter=IsActive eq 1&$orderby=Title asc`,
      SPHttpClient.configurations.v1
    ).then(r => r.json()).then(json => {
      const titles: string[] = (json?.value || []).map((x: { Title?: string }) => String(x?.Title || "")).filter(Boolean);
      setPurposeOptions(dedupeLabels(titles));
    }).catch(() => undefined);
  }, []);

  const statusOptions = React.useMemo(
    () =>
      dedupeLabels(
        ["Approved", "Rejected", "Submitted", "Draft", "Pending"],
        canonicalizeStatusFilterKey
      ),
    []
  );

  const hasActiveFilters =
    statusSelected.length > 0 ||
    purposeSelected.length > 0 ||
    costCenterSelected.length > 0 ||
    !!expectedFrom ||
    !!expectedTo;
  const hasAppliedFilters = hasActiveFilters || searchText.trim().length > 0;
  const displayToastMessage = toast ? normalizeToastMessage(toast.message) : "";
  const isDiscardToast =
    toast?.type === "error" && /\bdiscard(?:ed)?\b/i.test(displayToastMessage);

  return (
    <>
{loading && <Loading />}

      {!loading && (
        <div className="pdf-thai relative flex h-full flex-col">
          {/* HEADER — ALWAYS SHOWN */}
          <div className="relative shrink-0 sticky top-0 z-20 bg-[#f5f5f5]">
            {toast && (
              <div className="absolute top-3 left-0 right-0 z-[40]">
                <div>
                  <div
                    className={[
                      "w-full text-white rounded-md shadow-lg",
                      toast.type === "success"
                        ? "bg-green-600"
                        : toast.type === "error"
                          ? "bg-red-600"
                          : "bg-[#606060]",
                    ].join(" ")}
                  >
                    <div className="px-4 py-2 flex justify-between items-center gap-3">
                      <div className="flex items-center gap-3 min-w-0">
                        <span className="inline-flex h-5 w-5 items-center justify-center shrink-0">
                          {toast.type === "draft" ? (
                            <Icon iconName="Save" className="text-[14px] leading-none" />
                          ) : toast.type === "success" ? (
                            <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" aria-hidden="true">
                              <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2" />
                              <path d="M7 12.5L10.5 16L17 9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          ) : toast.type === "error" ? (
                            isDiscardToast ? (
                              <Icon iconName="Delete" className="text-[14px] leading-none" />
                            ) : (
                              <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" aria-hidden="true">
                                <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2" />
                                <path d="M9 9L15 15M15 9L9 15" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                              </svg>
                            )
                          ) : (
                            <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                              <g transform="scale(-1,1) translate(-24,0)">
                                <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                                <path d="M3 3v5h5" />
                              </g>
                              <path d="M12 7v5l4 2" />
                            </svg>
                          )}
                        </span>
                        <span className="text-sm font-normal truncate">{displayToastMessage}</span>
                      </div>
                      <button
                        onClick={() => setToast(null)}
                        className="text-white opacity-90 hover:opacity-100 shrink-0"
                        aria-label="Close"
                      >
                        <svg className="h-5 w-5" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                          <path d="M6 6L18 18M18 6L6 18" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                        </svg>
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )}

            <MyRequestHeader
              onCreate={props.onCreate}
              onSelectAllRequests={selectAllFilteredRows}
              onRestoreManualSelection={restoreManualSelection}
              manualSelectedCount={manualSelectedIds.length}
              onExportExcel={async (mode) => {
                if (!canExport) return;
                const data =
                  mode === "all"
                    ? sortedRows
                    : sortedRows.filter((r) =>
                      selectedIds.indexOf(String(r.id)) !== -1
                    );
                const itemsByRequestId = await fetchExportItemsByRequestIds(
                  context,
                  data.map((r) => r.spId)
                );
                const exportRows = data.map((r) => ({
                  ...r,
                  items: itemsByRequestId[r.spId] || [],
                }));
                await exportToExcel(exportRows);
              }}
              canExport={canExport}
              showDisabledExport={props.permissions.viewExportOwnReport && rows.length === 0 && !hasAppliedFilters}
            />
          </div>

          {/* BODY — SINGLE SCROLL CONTAINER */}
          <div className="relative flex-1 overflow-auto bg-[#f3f3f3] py-4 px-7 sm:px-6">
            <div className="space-y-6">
              {rows.length === 0 && !hasAppliedFilters ? (
                <RequestEmptyState
                  title="Empty requests"
                  description="You can create a new request here"
                  actionLabel="Let's Create Request"
                  onAction={props.onCreate}
                  showAction={true}
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
                    setStatusSelected={(v) => { setStatusSelected(v); setPageIndex(0); setSelectedIds([]); setManualSelectedIds([]); }}
                    setPurposeSelected={(v) => { setPurposeSelected(v); setPageIndex(0); setSelectedIds([]); setManualSelectedIds([]); }}
                    setCostCenterSelected={(v) => { setCostCenterSelected(v); setPageIndex(0); setSelectedIds([]); setManualSelectedIds([]); }}
                    setExpectedFrom={(v) => { setExpectedFrom(v); setPageIndex(0); setSelectedIds([]); setManualSelectedIds([]); }}
                    setExpectedTo={(v) => { setExpectedTo(v); setPageIndex(0); setSelectedIds([]); setManualSelectedIds([]); }}
                    onClearAll={() => {
                      setStatusSelected([]);
                      setPurposeSelected([]);
                      setCostCenterSelected([]);
                      setExpectedFrom("");
                      setExpectedTo("");
                      setPageIndex(0);
                      setSelectedIds([]);
                      setManualSelectedIds([]);
                    }}
                    showClearAll={hasActiveFilters}
                    SelectBox={SelectBox}
                  />
                  <div>
                    <div className="text-[13px] font-semibold text-gray-900">
                      Request lists
                    </div>
                    <div className="flex items-center gap-2 mt-1">
                      <svg
                        width="14"
                        height="14"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="#0a66d8"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                        <path d="M3 3v5h5" />
                        <path d="M12 7v5l4 2" />
                      </svg>
                      <p className="text-xs text-gray-500">
                        All draft will be deleted on 15 January of every year.
                      </p>
                    </div>
                  </div>

                  <div key={String(countIncreased)} className={`${card} overflow-hidden`}>
                    <div className="px-4 py-3 flex items-center justify-between border-b border-[#eeeeee]">
                      <RequestSearchBar
                        searchDraft={searchDraft}
                        setSearchDraft={setSearchDraft}
                        onSearch={() => {
                          setSearchText(searchDraft);
                          setPageIndex(0);
                        }}
                        onSearchWithField={({ field, text }) => {
                          setSearchField(field);
                          setSearchText(text);
                          setPageIndex(0);
                        }}
                        onClear={() => {
                          setSearchText("");
                          setSearchField("TransactionID");
                          setPageIndex(0);
                        }}
                      />
                    </div>

                    {tableLoading ? (
                      <div className="p-6">
                        <Loading />
                      </div>
                    ) : (
                      <>
                        <div className="p-0.1">
                          <ResponsiveTable<Row>
                            columns={columns}
                            rows={pageRows}
                            minWidthClass="min-w-0"
                            tableBreakpoint={1024}
                            wrapHeaderWords={true}
                            disableHorizontalScroll={true}
                            showCheckbox
                            selectedIds={selectedIds}
                            onToggleRow={handleToggleRow}
                            onToggleAll={handleToggleAll}
                            headerChecked={
                              sortedRows.length > 0 &&
                              sortedRows.every((r) => selectedIds.includes(String(r.id)))
                            }
                          />
                        </div>

                        <div className="px-4 py-3 flex items-center justify-between text-[12px] text-gray-600 border-t border-[#eeeeee] bg-[#fafafa]">

                          {/* LEFT SIDE BUTTONS */}
                          <div className="flex items-center gap-2">
                            <button
                              disabled={safePageIndex === 0}
                              onClick={() => setPageIndex(0)}
                              className="rounded-md border border-gray-300 bg-[#FFFFFF] px-2.5 py-1 text-[12px] font-medium text-[#2172DA] hover:bg-[#2172DA] hover:text-white disabled:cursor-not-allowed disabled:border-[#D9D9D9] disabled:bg-[#D9D9D9] disabled:text-gray-500" >
                              First
                            </button>

                            <button
                              disabled={safePageIndex === 0}
                              onClick={() => setPageIndex((p) => Math.max(0, p - 1))}
                              className="rounded-md border border-gray-300 bg-[#FFFFFF] px-2.5 py-1 text-[12px] font-medium text-[#2172DA] hover:bg-[#2172DA] hover:text-white disabled:cursor-not-allowed disabled:border-[#D9D9D9] disabled:bg-[#D9D9D9] disabled:text-gray-500" >
                              Previous
                            </button>

                            <span className="px-3 text-[12px] text-gray-700">
                              Page {safePageIndex + 1} of {totalPages}
                            </span>

                            <button
                              disabled={safePageIndex >= totalPages - 1}
                              onClick={() =>
                                setPageIndex((p) => Math.min(totalPages - 1, p + 1))
                              }
                              className="rounded-md border border-gray-300 bg-[#FFFFFF] px-2.5 py-1 text-[12px] font-medium text-[#2172DA] hover:bg-[#2172DA] hover:text-white disabled:cursor-not-allowed disabled:border-[#D9D9D9] disabled:bg-[#D9D9D9] disabled:text-gray-500">
                              Next
                            </button>

                            <button
                              disabled={safePageIndex >= totalPages - 1}
                              onClick={() => setPageIndex(totalPages - 1)}
                              className="rounded-md border border-gray-300 bg-[#FFFFFF] px-2.5 py-1 text-[12px] font-medium text-[#2172DA] hover:bg-[#2172DA] hover:text-white disabled:cursor-not-allowed disabled:border-[#D9D9D9] disabled:bg-[#D9D9D9] disabled:text-gray-500" >
                              Last
                            </button>

                          </div>


                          {/* RIGHT SIDE INFO */}
                          {/* <div>
                            Showing {pageRows.length} of {sortedRows.length} items (pageSize {pageSize})
                          </div> */}
                        </div>

                      </>
                    )}
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      )}
    </>
  );
}