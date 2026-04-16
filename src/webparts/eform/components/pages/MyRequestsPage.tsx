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
import { formatDateDdMmYyyy } from "./AllRequestPage/dateUtils";
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
  expected: string;
  qty: number;
  price: number;
  spId: number;
  expectedIso?: string;
  submittedIso?: string;
  requestCreatedIso?: string;
  modifiedIso?: string;
  lastActivityIso?: string;
  ApproverNames?:string
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
  Department?: { Title: string };
  PurposeDetails?: string;
  RequestSubmittedOn?: string;
  FinalDecisionDate?: string;
  ExpectedDeliveryDate?: string;
  TotalQuantity?: number;
  TotalAmount?: number;
  WorkflowStatus?: Status;
  RequestCreatedOn?: string;
  Modified?: string;
  ApproverNames?:string
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

  // Display text
  const display =
    key === "submitted"
      ? "Submitted"
      : isPending
        ? "Pending"
        : isInApproval
          ? "In-Approval"
          : key === "approved"
            ? "Approved"
            : key === "rejected"
              ? "Rejected"
              : key === "draft"
                ? "Draft"
                : status;

  // Color mapping
  const colorMap: Record<string, string> = {
    approved: "text-green-600",
    rejected: "text-red-600",
    pending: "text-yellow-600",
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
  // ✅ Show empty string instead of "-" when no value
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
  const [allRows, setAllRows] = React.useState<Row[]>([]); // ✅ ADD THIS

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
  // âœ… store only manual checkbox selections
  const [manualSelectedIds, setManualSelectedIds] = React.useState<string[]>([]);

  const [searchDraft, setSearchDraft] = React.useState("");
  const [searchText, setSearchText] = React.useState("");
  // SORT
  const [sortDir, setSortDir] = React.useState<"asc" | "desc">("desc");
  const [toast, setToast] = React.useState<null | { type: ToastType; message: string }>(null);
  const [pageIndex, setPageIndex] = React.useState(0);
  const [searchField, setSearchField] = React.useState<"TransactionID" | "RequesterName">("TransactionID");
  const [isAllSelected, setIsAllSelected] = React.useState(false);

  const pageSize = 15;
  // âœ… Toggle single row
  const handleToggleRow = (id: string | number): void => {
  const idStr = String(id);
  setIsAllSelected(false); // exit "all selected" mode when manually toggling
  setSelectedIds((prev) => {
    const updated = prev.includes(idStr)
      ? prev.filter((x) => x !== idStr)
      : [...prev, idStr];
    setManualSelectedIds(updated);
    return updated;
  });
};


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
    searchField
  ]);


  const sortedRows = React.useMemo(() => {
    const out = [...filteredRows];

    // Sort by transaction ID
    out.sort((a, b) => {
      const cmp = a.id.localeCompare(b.id, undefined, { numeric: true });
      return sortDir === "asc" ? cmp : -cmp;
    });

    return out;
  }, [filteredRows, sortDir]);

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
  setIsAllSelected(true); 
};


const restoreManualSelection = (): void => {
  setIsAllSelected(false);
  setSelectedIds(manualSelectedIds);
};



  // âœ… Toggle all rows (current page only)
  const handleToggleAll = (): void => {
  if (isAllSelected) {
    
    setIsAllSelected(false);
    setSelectedIds([]);
    setManualSelectedIds([]);
    return;
  }

  const pageIds = pageRows.map((r) => String(r.id));
  const allSelectedOnPage = pageIds.every((id) => selectedIds.includes(id));

  if (allSelectedOnPage) {
    // deselect current page
    const updated = selectedIds.filter((id) => !pageIds.includes(id));
    setSelectedIds(updated);
    setManualSelectedIds(updated);
  } else {
    // select current page
    const updated = [...selectedIds, ...pageIds.filter((id) => !selectedIds.includes(id))];
    setSelectedIds(updated);
    setManualSelectedIds(updated);
  }
};



  React.useEffect(() => {
    if (!toast) return;
    // const timer = setTimeout(() => setToast(null), 4000);
    // return () => clearTimeout(timer);
  }, [toast]);

  React.useEffect(() => {
    const incomingType = props.routeState?.toastType;
    if (!incomingType) return;

    const fallbackMessage =
      incomingType === "success"
        ? "Your request is submitted to the approver(s)."
        : incomingType === "pending"
          ? "Your request status will be shown as “Pending” while it is being processed."
          : incomingType === "draft"
            ? "Your request is saved as draft."
            : "Your request is discarded.";

    setToast({
      type: incomingType,
      message: normalizeToastMessage(props.routeState?.toastMessage || fallbackMessage),
    });

    props.onRouteStateConsumed?.();
  }, [props.routeState, props.onRouteStateConsumed]);

  // paging



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
  const loadAll = React.useCallback(async (): Promise<void> => {
    if (!hasLoadedOnce) {
      setLoading(true);
    } else {
      setTableLoading(true);
    }
    setError("");
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const myEmail = String(context.pageContext.user.email || "").trim();
      const esc = (v: string): string => String(v || "").replace(/'/g, "''");
      const clauses: string[] = [
        `Requester/EMail eq '${esc(myEmail)}'`,
        `WorkflowStatus ne 'Failed'` // 🚫 Exclude Failed records
      ];

      if (statusSelected.length > 0) {
        const statusTerms: string[] = [];

        statusSelected.forEach((status) => {
          const key = canonicalizeStatusFilterKey(status);

          if (key === canonicalizeStatusFilterKey("Pending")) {
            // ✅ ONLY Pending
            statusTerms.push(`WorkflowStatus eq 'Pending'`);
          }
          else if (key === canonicalizeStatusFilterKey("In Approval")) {
            // ✅ ONLY In Approval
            statusTerms.push(`WorkflowStatus eq 'In Approval'`);
          }
          else if (key === canonicalizeStatusFilterKey("Submitted")) {
            // ✅ ONLY Submitted
            statusTerms.push(`WorkflowStatus eq 'Submitted'`);
          }
          else if (key === canonicalizeStatusFilterKey("Approved")) {
            statusTerms.push(`WorkflowStatus eq 'Approved'`);
          }
          else if (key === canonicalizeStatusFilterKey("Rejected")) {
            statusTerms.push(`WorkflowStatus eq 'Rejected'`);
          }
          else if (key === canonicalizeStatusFilterKey("Draft")) {
            statusTerms.push(`WorkflowStatus eq 'Draft'`);
          }
        });

        // ✅ Remove duplicates safely
        const uniqueTerms = Array.from(new Set(statusTerms));

        // ✅ Apply filter
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
        "?$top=5000&$select=Id,Title,RequestNumber,WorkflowStatus,TotalQuantity,TotalAmount,Modified," +
        "PurposeDetails,RequestCreatedOn,RequestSubmittedOn,FinalDecisionDate,ExpectedDeliveryDate,ApproverNames," +
        "Requester/Title,Requester/EMail,CurrentApprover/Title,Purpose/Title,CostCenter/Title,Department/Title" +
        "&$expand=Requester,CurrentApprover,Purpose,CostCenter,Department" +
        `&$filter=${clauses.join(" and ")}` +
        "&$orderby=Id desc";

      const allItems: InventoryRequestSPItem[] = [];
      let nextUrl: string | undefined = url;
      while (nextUrl) {
        const json: SPListResponse<InventoryRequestSPItem> & { "@odata.nextLink"?: string } =
          await spGet(nextUrl);
        allItems.push(...(json.value || []));
        nextUrl = json["@odata.nextLink"];
      }

      // ✅ Extract mappedRows as a named variable
      const mappedRows: Row[] = (allItems || []).map((it) => {
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
          department: it.Department?.Title || "",
          purpose: it.Purpose?.Title || "",
          purposeDetail: it.PurposeDetails || "",
          costCenter: it.CostCenter?.Title || "",
          submitted: formatDateDdMmYyyy(it.RequestSubmittedOn),
          approved: formatDateDdMmYyyy(it.FinalDecisionDate),
          expected: formatDateDdMmYyyy(it.ExpectedDeliveryDate),
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

      setRows(mappedRows);      // ✅ existing
      setAllRows(mappedRows);   // ✅ ADD — only set on first load (no filters active)
      setPageIndex(0);

    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Failed to load data");
    } finally {
      setLoading(false);
      setTableLoading(false);
      setHasLoadedOnce(true);
    }
  }, [
    context,
    hasLoadedOnce,
  ]);

  React.useEffect(() => {
    loadAll().catch(console.error);
  }, []); // ✅ ONLY ONCE



  /* ================= COLUMNS ================= */

  const columns = [
    {
  key: "id",
  header: "Transaction<br/>ID",
  thClassName: "w-[10%]",
  tdClassName: "w-[10%] truncate",
  render: (r: Row) =>
    r.id === "-" || r.id === " " || r.id === "" ? (
      // ✅ Show empty for Pending rows (id is "-")
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
        <CellText value={r.status === "Approved" || r.status === "Rejected" ? (r.approver === "-" ? "" : r.approver) : ""} maxWidth="max-w-full" />
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
      tdClassName: "w-[7%] whitespace-nowrap",
      render: (r: Row) => (r.approved === "-" ? "" : r.approved),
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
      render: (r: Row) => r.qty,
    },

    {
      key: "price",
      header: "Total Tag<br/>Price",
      thClassName: "w-[7%] text-left",
      tdClassName: "w-[7%] whitespace-nowrap text-left",
      render: (r: Row) =>
        Number(r.price).toLocaleString(undefined, {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2
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
              className=" w-full max-w-[80px] rounded-md  bg-[#F6F6F6] px-1.5 py-1.5 text-[9px] font-semibold whitespace-nowrap text-[#81AFEB] border border-[#EBEBEB] hover:bg-[#F0F7FF] hover:border hover:border-[#2172DA] hover:text-blue-400"
              onClick={() => props.onDownloadPdf(r.spId)}
            >
              Export .PDF
            </button>
          </div>
        );
      },
    },
  ];
  // Load Purpose & Cost Center from master lists (same source as Create Request)
  const [purposeOptions, setPurposeOptions] = React.useState<string[]>([]);
  const [costCenterOptions, setCostCenterOptions] = React.useState<string[]>([]);

  React.useEffect(() => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const email = context.pageContext.user.email.replace(/'/g, "''");

    // Load Purpose from PurposeMaster
    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('PurposeMaster')/items?$top=5000&$select=Title&$filter=IsActive eq 1&$orderby=Title asc`,
      SPHttpClient.configurations.v1
    ).then(r => r.json()).then(json => {
      const titles: string[] = (json?.value || []).map((x: { Title?: string }) => String(x?.Title || "")).filter(Boolean);
      setPurposeOptions(dedupeLabels(titles));
    }).catch(() => undefined);

    // Get department ID from EForm Users (same as Create Request page)
    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('EForm%20Users')/items?$top=1&$select=Department/Id&$expand=Department&$filter=Email eq '${email}'`,
      SPHttpClient.configurations.v1
    ).then(r => r.json()).then(userJson => {
      const deptId = userJson?.value?.[0]?.Department?.Id;
      const deptFilter = deptId ? ` and Department/Id eq ${deptId}` : "";
      return context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('CostCenterMaster')/items?$top=5000&$select=Title,Department/Id&$expand=Department&$filter=IsActive eq 1${deptFilter}&$orderby=Title asc`,
        SPHttpClient.configurations.v1
      );
    }).then(r => r.json()).then(json => {
      const titles: string[] = (json?.value || []).map((x: { Title?: string }) => String(x?.Title || "")).filter(Boolean);
      setCostCenterOptions(dedupeLabels(titles));
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
          {/* âœ… HEADER â€“ ALWAYS SHOWN */}
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
                          : "bg-[#595959]",
                    ].join(" ")}
                  >
                    <div className="px-4 py-2 flex justify-between items-center gap-3">
                      <div className="flex items-center gap-3 min-w-0">
                        <span className="inline-flex h-4 w-4 items-center justify-center shrink-0">
                          {toast.type === "draft" ? (
                            <Icon iconName="Save" className="text-[14px] leading-none" />
                          ) : toast.type === "success" ? (
                            "✓"
                          ) : toast.type === "error" ? (
                            isDiscardToast ? (
                              <Icon iconName="Delete" className="text-[14px] leading-none" />
                            ) : (
                              "✕"
                            )
                          ) : (
                            "↻"
                          )}
                        </span>
                        <span className="text-sm font-medium truncate">{displayToastMessage}</span>
                      </div>
                      <button
                        onClick={() => setToast(null)}
                        className="text-white text-base leading-none opacity-90 hover:opacity-100"
                        aria-label="Close"
                      >
                        ×
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* âœ… HEADER â€“ ALWAYS SHOWN */}
            <MyRequestHeader
              onCreate={props.onCreate}
              selectedCount={selectedIds.length}
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


          {/* âœ… BODY â€“ SINGLE SCROLL CONTAINER */}
          <div
            className={`relative flex-1 overflow-auto bg-[#f3f3f3] py-4 px-7 sm:px-6`}
          >
            <div className={`space-y-6`}>
              {/* âœ… SWITCH ONLY CONTENT */}
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
                    setStatusSelected={(v) => { setStatusSelected(v); setPageIndex(0); }}
                    setPurposeSelected={(v) => { setPurposeSelected(v); setPageIndex(0); }}
                    setCostCenterSelected={(v) => { setCostCenterSelected(v); setPageIndex(0); }}
                    setExpectedFrom={(v) => { setExpectedFrom(v); setPageIndex(0); }}
                    setExpectedTo={(v) => { setExpectedTo(v); setPageIndex(0); }}
                    onClearAll={() => {
                      setStatusSelected([]);
                      setPurposeSelected([]);
                      setCostCenterSelected([]);
                      setExpectedFrom("");
                      setExpectedTo("");
                      setPageIndex(0);
                    }}
                    showClearAll={hasActiveFilters}
                    SelectBox={SelectBox}
                  />
                  <div>
                    {/* Title */}
                    <div className="text-[13px] font-semibold text-gray-900">
                      Request lists
                    </div>

                    {/* Clock + Content */}
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


                  <div className={`${card} overflow-hidden`}>
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
                          />
                        </div>

                        <div className="px-4 py-3 flex items-center justify-between text-[12px] text-gray-600 border-t border-[#eeeeee] bg-[#fafafa]">

                          {/* LEFT SIDE BUTTONS */}
                          <div className="flex items-center gap-2">
                            <button
                              disabled={safePageIndex === 0}
                              onClick={() => setPageIndex(0)}
                              className="rounded-md border border-gray-300 bg-white px-2.5 py-1 text-[12px] font-medium text-blue-600 hover:bg-blue-50 disabled:cursor-not-allowed disabled:border-gray-200 disabled:bg-gray-100 disabled:text-gray-500" >
                              First
                            </button>

                            <button
                              disabled={safePageIndex === 0}
                              onClick={() => setPageIndex((p) => Math.max(0, p - 1))}
                              className="rounded-md border border-gray-300 bg-white px-2.5 py-1 text-[12px] font-medium text-blue-600 hover:bg-blue-50 disabled:cursor-not-allowed disabled:border-gray-200 disabled:bg-gray-100 disabled:text-gray-500" >
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
                              className="rounded-md border border-gray-300 bg-white px-2.5 py-1 text-[12px] font-medium text-blue-600 hover:bg-blue-50 disabled:cursor-not-allowed disabled:border-gray-200 disabled:bg-gray-100 disabled:text-gray-500">
                              Next
                            </button>

                            <button
                              disabled={safePageIndex >= totalPages - 1}
                              onClick={() => setPageIndex(totalPages - 1)}
                              className="rounded-md border border-gray-300 bg-white px-2.5 py-1 text-[12px] font-medium text-blue-600 hover:bg-blue-50 disabled:cursor-not-allowed disabled:border-gray-200 disabled:bg-gray-100 disabled:text-gray-500" >
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
