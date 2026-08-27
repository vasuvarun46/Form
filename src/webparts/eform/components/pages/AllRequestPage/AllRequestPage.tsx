import * as React from "react";
import {
  AllRequestPageProps,
  InventoryRequestSPItem,
  Row,
  Status,
} from "./AllRequestTypes";
import {
  canonicalizeStatusFilterKey,
  dedupeLabels,
} from "./allRequestFilterUtils";
import AllRequestRender from "./AllRequestReturnFunction";
import { exportToExcel, fetchExportItemsByRequestIds } from "../ExportExcel";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

/* =====================================================
   STATUS TEXT COMPONENT
===================================================== */

function StatusText({ status }: { status: string }): JSX.Element {
  const key = String(status || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");

  const isInApproval =
    key === "pending" ||
    key === "in approval" ||
    key === "in-approval";

  // Display text
  const display =
    key === "submitted" || isInApproval
      ? "Submitted"
      : key === "approved"
          ? "Approved"
          : key === "rejected"
            ? "Rejected"
            : key === "draft"
              ? "Draft"
              : status;

  // Color logic (UPDATED)
  const cls =
    key === "approved"
      ? "text-green-600"
      : key === "rejected"
        ? "text-red-600"
        : key === "draft"
          ? "text-gray-600" // 
          : key === "submitted" || isInApproval
            ? "text-blue-600"
            : "text-gray-500";


  return <span className={`font-medium ${cls}`}>{display}</span>;
}

function CellText(props: { value: string; maxWidth?: string }): JSX.Element {
  const raw = String(props.value || "").trim();
  const display = raw === "-" || raw === "" ? "" : raw;
  return (
    <div className={`${props.maxWidth || "max-w-[110px]"} truncate`} title={display}>
      {display}
    </div>
  );
}

/* =====================================================
   MAIN COMPONENT
===================================================== */

export default function AllRequestPage(
  props: AllRequestPageProps
): JSX.Element {
  const { context } = props;

  const pageBg = "bg-[#f3f3f3]";
  const container = "w-full max-w-[1320px] mx-auto";
  const card = "bg-white rounded-xl border border-[#e5e5e5] shadow-sm";

  const [rows, setRows] = React.useState<Row[]>([]);
  const [loadingPage, setLoadingPage] = React.useState<boolean>(false);
  const [tableLoading, setTableLoading] = React.useState<boolean>(false);
  const [hasLoadedOnce, setHasLoadedOnce] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>("");

  const [pageSize] = React.useState<number>(15);
  const [pageIndex, setPageIndex] = React.useState<number>(0);

  const [searchDraft, setSearchDraft] = React.useState<string>("");
  const [searchText, setSearchText] = React.useState<string>("");
  const [searchField, setSearchField] = React.useState<"TransactionID" | "RequesterName">("TransactionID");
  

  const [statusFilter, setStatusFilter] = React.useState<string[]>([]);
  const [purposeFilter, setPurposeFilter] = React.useState<string[]>([]);
  const [costCenterFilter, setCostCenterFilter] = React.useState<string[]>([]);
  const [expectedFrom, setExpectedFrom] = React.useState<string>("");
  const [expectedTo, setExpectedTo] = React.useState<string>("");

  const [selectedIds, setSelectedIds] = React.useState<string[]>([]);
  const [manualSelectedIds, setManualSelectedIds] = React.useState<string[]>([]);

  /* ================= DROPDOWN OPTION STATES ================= */

  const [statusOptions, setStatusOptions] = React.useState<string[]>([]);
  const [purposeOptions, setPurposeOptions] = React.useState<string[]>([]);

  const STORAGE_KEY_APPROVED = "allreq_approved_count";
  const STORAGE_KEY_REJECTED = "allreq_rejected_count";
  const STORAGE_KEY_SUBMITTED = "allreq_submitted_count";

  // ✅ Toggles to force re-render when any count increases
  const [countIncreased, setCountIncreased] = React.useState<boolean>(false);


  /* =====================================================
     LOAD DROPDOWN OPTIONS
  ===================================================== */
  const canExport = props.permissions?.exportAllReport === true;

  const loadStatusOptions = async (): Promise<void> => {
    const siteUrl = context.pageContext.web.absoluteUrl;

    const response = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/fields/getbyinternalnameortitle('WorkflowStatus')`,
      SPHttpClient.configurations.v1
    );

    const json = await response.json();
    const choices: string[] = Array.isArray(json?.Choices)
      ? json.Choices.map((choice: unknown) => String(choice || ""))
      : [];
    const draftKey      = canonicalizeStatusFilterKey("Draft");
    const failedKey     = canonicalizeStatusFilterKey("Failed");
    const inApprovalKey = canonicalizeStatusFilterKey("In Approval");
    const exclude = (status: string): boolean => {
      const key = canonicalizeStatusFilterKey(status);
      return key !== draftKey && key !== failedKey && key !== inApprovalKey;
    };
    const STATUS_ORDER = ["Approved", "Rejected", "Submitted"];
    const cleaned = dedupeLabels(choices, canonicalizeStatusFilterKey)
      .filter(exclude)
      .sort((a, b) => {
        const ai = STATUS_ORDER.findIndex(s => canonicalizeStatusFilterKey(s) === canonicalizeStatusFilterKey(a));
        const bi = STATUS_ORDER.findIndex(s => canonicalizeStatusFilterKey(s) === canonicalizeStatusFilterKey(b));
        return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
      });
    setStatusOptions(cleaned);
    setStatusFilter((prev) => prev.filter(exclude));
  };

  const loadPurposeOptions = async (): Promise<void> => {
    const siteUrl = context.pageContext.web.absoluteUrl;

    const response = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('PurposeMaster')/items?$select=Title&$filter=IsActive eq 1&$orderby=Title asc`,
      SPHttpClient.configurations.v1
    );

    const json = await response.json();
    const titles: string[] = (Array.isArray(json?.value) ? json.value : []).map(
      (item: { Title?: string }) => String(item?.Title || "")
    );

    setPurposeOptions(dedupeLabels(titles));
  };

  const [costCenterOptions, setCostCenterOptions] = React.useState<string[]>([]);
  const loadCostCenterOptions = async (): Promise<void> => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const response = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('CostCenterMaster')/items?$select=Title&$filter=IsActive eq 1&$orderby=Title asc`,
      SPHttpClient.configurations.v1
    );
    const json = await response.json();
    const titles: string[] = (Array.isArray(json?.value) ? json.value : []).map(
      (item: { Title?: string }) => String(item?.Title || "")
    ).filter(Boolean);
    setCostCenterOptions(dedupeLabels(titles).sort((a, b) => a.localeCompare(b)));
  };

  /* =====================================================
     FILTERING LOGIC (UNCHANGED)
  ===================================================== */

  const filteredRows: Row[] = React.useMemo(() => {
    const seen = new Set<string>();
    const deduped: Row[] = [];
    for (const r of rows) {
      const dedupKey = r.id && r.id !== "-" ? r.id : String(r.spId);
      if (!seen.has(dedupKey)) {
        seen.add(dedupKey);
        deduped.push(r);
      }
    }
    return deduped;
  }, [rows, statusFilter, purposeFilter, costCenterFilter, expectedFrom, expectedTo, searchText]);

  const sortedRows: Row[] = filteredRows;
  const totalPages = Math.max(1, Math.ceil(sortedRows.length / pageSize));

  const safePageIndex = Math.min(pageIndex, totalPages - 1);

  const pageRows: Row[] = sortedRows.slice(
    safePageIndex * pageSize,
    safePageIndex * pageSize + pageSize
  );

  /* =====================================================
     SELECTION LOGIC (UNCHANGED)
  ===================================================== */

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

  const onSelectAllRequests = (): void => {
    const ids = sortedRows.map((r) => String(r.id));
    setSelectedIds(ids);
    setManualSelectedIds(ids);
  };

  const onRestoreManualSelection = (): void => {
    setSelectedIds(manualSelectedIds);
  };



  function formatDateDdMmYyyy(iso: string | undefined): string {
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

  const onExportExcel = async (mode: "all" | "selected") => {
    if (!canExport) return;

    const data =
      mode === "all"
        ? sortedRows
        : sortedRows.filter((r) =>
          manualSelectedIds.includes(String(r.id))
        );

    if (data.length === 0) {
      setError("No rows selected to export.");
      return;
    }

    try {
      const sortedByInternalIdDesc = [...data].sort((a, b) => b.spId - a.spId);
      const itemsByRequestId = await fetchExportItemsByRequestIds(
        context,
        sortedByInternalIdDesc.map((r) => r.spId)
      );
      const exportRows = sortedByInternalIdDesc.map((r) => ({
        ...r,
        items: itemsByRequestId[r.spId] || [],
      }));

      await exportToExcel(exportRows);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : "Export failed.";
      setError(`Excel export failed: ${msg}`);
    }
  };

  /* =====================================================
     LOAD DATA
  ===================================================== */
  const loadAll = React.useCallback(async (silent: boolean = false): Promise<void> => {
    if (!silent) {
      if (!hasLoadedOnce) {
        setLoadingPage(true);
      } else {
        setTableLoading(true);
      }
    }

    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const esc = (v: string): string => String(v || "").replace(/'/g, "''");
      const clauses: string[] = [`WorkflowStatus ne 'Draft' and WorkflowStatus ne 'Pending' and WorkflowStatus ne 'Failed'`];

      if (statusFilter.length > 0) {
        const statusTerms: string[] = [];
        const draftKey = canonicalizeStatusFilterKey("Draft");
        statusFilter.forEach((status) => {
          const key = canonicalizeStatusFilterKey(status);
          if (key === draftKey) {
            return;
          }
          if (key === canonicalizeStatusFilterKey("Pending")) {
            statusTerms.push(`WorkflowStatus eq 'Pending'`);
            statusTerms.push(`WorkflowStatus eq 'In Approval'`);
          } else {
            statusTerms.push(`WorkflowStatus eq '${esc(status)}'`);
          }
        });
        if (statusTerms.length > 0) {
          clauses.push(`(${Array.from(new Set(statusTerms)).join(" or ")})`);
        }
      }

      if (purposeFilter.length > 0) {
        const purposeTerms = purposeFilter.map((p) => `Purpose/Title eq '${esc(p)}'`);
        clauses.push(`(${purposeTerms.join(" or ")})`);
      }

      if (costCenterFilter.length > 0) {
        const ccTerms = costCenterFilter.map((c) => `CostCenter/Title eq '${esc(c)}'`);
        clauses.push(`(${ccTerms.join(" or ")})`);
      }

      if (expectedFrom) {
        clauses.push(`ExpectedDeliveryDate ge datetime'${expectedFrom}T00:00:00Z'`);
      }
      if (expectedTo) {
        clauses.push(`ExpectedDeliveryDate le datetime'${expectedTo}T23:59:59Z'`);
      }

      if (searchText.trim()) {
        const q = esc(searchText.trim());
        if (searchField === "RequesterName") {
          clauses.push(`substringof('${q}',Requester/Title)`);
        } else {
          clauses.push(`substringof('${q}',Title)`);
        }
      }

      // ✅ Your same query (Department is Lookup)
      const firstUrl =
        `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items` +
        `?$top=5000` +
        `&$orderby=Id desc` +
        `&$filter=${clauses.join(" and ")}` +
        `&$select=` +
        [
          "Id",
          "RequestNumber",
          "Title",
          "Departments",
          "WorkflowStatus",
          "TotalQuantity",
          "ApproverNames",
          "TotalAmount",
          "PurposeDetails",
          "RequestSubmittedOn",
          "FinalDecisionDate",
          "ApprovedDates",
          "ExpectedDeliveryDate",
          "Modified",
          "Requester/Title",
          "CurrentApprover/Title",
          "Purpose/Title",
          "CostCenter/Title",
          "MultipleApproverData",
        ].join(",") +
        `&$expand=` +
        ["Requester", "CurrentApprover", "Purpose", "CostCenter"].join(",");

      // ✅ Fetch all pages using @odata.nextLink
      const reqItems: InventoryRequestSPItem[] = [];
      let nextUrl: string | undefined = firstUrl;

      while (nextUrl) {
        const res: SPHttpClientResponse = await context.spHttpClient.get(
          nextUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              'odata-version': ''
            },
          }
        );

        if (!res.ok) {
          const text = await res.text();
          throw new Error(`SharePoint request failed (${res.status}): ${text}`);
        }

        const json: any = await res.json();

        // With odata=nometadata, items are usually in "value"
        const batch: any[] = Array.isArray(json?.value) ? json.value : [];
        reqItems.push(...(batch as InventoryRequestSPItem[]));

        // Next link name can vary depending on metadata mode
        nextUrl =
          json?.["@odata.nextLink"] ||
          json?.["odata.nextLink"] ||
          json?.d?.__next ||
          undefined;
      }


      const mapped: Row[] = reqItems.map((it: any) => {
        const workflowStatus = String(it.WorkflowStatus ?? "Submitted");
        const isPending = workflowStatus.trim().toLowerCase() === "pending";
        const titleTxId = String(it?.Title || "").trim();
        const hasFinalTitle =
          !!titleTxId &&
          !/^draft$/i.test(titleTxId) &&
          !/^temp$/i.test(titleTxId);

        return {
          id: isPending
            ? "-"
            : hasFinalTitle
              ? titleTxId
              : it?.RequestNumber
                ? String(it.RequestNumber)
                : String(it?.Id || ""),
          status: workflowStatus as Status,
          requestor: it.Requester?.Title ?? "-",
          approver: it.ApproverNames,
          multipleApproverData: it.MultipleApproverData,
          department: String(it.Departments || "-"),
          purpose: it.Purpose?.Title ?? "-",
          PurposeDetails: it.PurposeDetails ?? "-",
          costCenter: it.CostCenter?.Title ?? "-",
          submitted: formatDateDdMmYyyy(it.RequestSubmittedOn),
          approved: formatDateDdMmYyyy(it.FinalDecisionDate),
          approvedDates: it.ApprovedDates ? it.ApprovedDates.split(",").map((d: string) => { const s = d.trim(); if (!s) return ""; const p = s.split("-"); if (p.length === 3 && p[0].length === 4) return `${p[2]}-${p[1]}-${p[0]}`; return s.replace(/\//g, "-"); }).filter(Boolean).join(", ") : "",
          expected: formatDateDdMmYyyy(it.ExpectedDeliveryDate),
          qty: Number(it.TotalQuantity ?? 0),
          price: Number(it.TotalAmount ?? 0),
          spId: it.Id,
          expectedIso: it.ExpectedDeliveryDate,
          submittedIso: it.RequestSubmittedOn,
          modifiedIso: it.Modified,
          lastActivityIso: it.Modified || it.RequestSubmittedOn || "",
        };
      }).filter((row) =>
        canonicalizeStatusFilterKey(String(row.status || "")) !==
        canonicalizeStatusFilterKey("Draft")
      );

      setRows(mapped);
      setPageIndex(0);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Failed to load data.");
      setRows([]);
      setPageIndex(0);
    } finally {
      if (!silent) {

        setLoadingPage(false);
        setTableLoading(false);
      }
      setHasLoadedOnce(true);
    }
  }, [
    context,
    hasLoadedOnce,
    statusFilter,
    purposeFilter,
    costCenterFilter,
    expectedFrom,
    expectedTo,
    searchText,
    searchField,
  ]);

  React.useEffect(() => {
    loadAll(false).catch(console.error);
  }, [loadAll]);

  React.useEffect(() => {
    loadStatusOptions().catch(() => undefined);
    loadPurposeOptions().catch(() => undefined);
    loadCostCenterOptions().catch(() => undefined);
  }, []);

  // ✅ FIXED — no user filter, counts ALL records across all users
  const getApprovedRejectedCounts = React.useCallback(async (): Promise<{
    approved: number;
    rejected: number;
    submitted: number;
  }> => {
    const siteUrl = context.pageContext.web.absoluteUrl;

    // ✅ No Requester/EMail filter — AllRequestPage sees ALL users' records
    const buildUrl = (status: string): string =>
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items` +
      `?$select=Id` +
      `&$filter=WorkflowStatus eq '${status}'` +
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

    const [approved, rejected, submitted] = await Promise.all([
      fetchCount("Approved"),
      fetchCount("Rejected"),
      fetchCount("Submitted"),  // ✅ tracks new submissions from any user
    ]);

    return { approved, rejected, submitted };
  }, [context]);

  /* ── 3. Compare counts, re-render + reload if increased ── */
  const checkAndUpdateCounts = React.useCallback(
    async (isInitialLoad: boolean): Promise<void> => {
      try {
        // ✅ correctly named variables matching what they actually hold
        const {
          approved: currentApproved,
          rejected: currentRejected,
          submitted: currentSubmitted,   // ✅ was wrongly named currentDraft
        } = await getApprovedRejectedCounts();

        const prevApproved = Number(localStorage.getItem(STORAGE_KEY_APPROVED) ?? 0);
        const prevRejected = Number(localStorage.getItem(STORAGE_KEY_REJECTED) ?? 0);
        const prevSubmitted = Number(localStorage.getItem(STORAGE_KEY_SUBMITTED) ?? 0); // ✅ correct name

        const hasIncreased =
          !isInitialLoad &&
          (
            currentApproved > prevApproved ||
            currentRejected > prevRejected ||
            currentSubmitted > prevSubmitted  // ✅ correct name
          );

        if (hasIncreased) {
          setCountIncreased((prev) => !prev);
          await loadAll(true);
        }

        // ✅ persist all 3 with correct variable names
        localStorage.setItem(STORAGE_KEY_APPROVED, String(currentApproved));
        localStorage.setItem(STORAGE_KEY_REJECTED, String(currentRejected));
        localStorage.setItem(STORAGE_KEY_SUBMITTED, String(currentSubmitted)); // ✅ correct

      } catch (err) {
        console.error("[AllRequestPage] checkAndUpdateCounts error:", err);
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

  const goFirst = () => setPageIndex(0);
  const goPrev = () => setPageIndex((p) => Math.max(0, p - 1));
  const goNext = () =>
    setPageIndex((p) => Math.min(totalPages - 1, p + 1));
  const goLast = () => setPageIndex(totalPages - 1);

  const hasFiltersApplied =
    statusFilter.length > 0 ||
    purposeFilter.length > 0 ||
    costCenterFilter.length > 0 ||
    !!expectedFrom ||
    !!expectedTo ||
    searchText.trim().length > 0;

  const isError = !loadingPage && !!error;
  const isBaseEmpty = !loadingPage && !tableLoading && !error && rows.length === 0 && !hasFiltersApplied;
  const hasAnyData = !loadingPage && !error && (rows.length > 0 || hasFiltersApplied || tableLoading);


  const columns = [
    {
      key: "id",
      header: "Transaction<br/>ID",
      thClassName: "w-[11%]",
      tdClassName: "w-[11%] truncate",
      render: (r: Row) =>
        r.id === "-" ? (
          <span className="text-gray-400"></span>  // ✅ was "-"
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
      render: (r: Row) => <StatusText status={r.status} />,
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
        <CellText value={r.PurposeDetails === "-" ? "" : r.PurposeDetails} maxWidth="max-w-full" />
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
      thClassName: "w-[8%]",
      tdClassName: "w-[8%] whitespace-nowrap",
      render: (r: Row) => (r.submitted === "-" || !r.submitted ? "" : r.submitted),
    },
    {
      key: "approved",
      header: "Approved<br/>Date",
      thClassName: "w-[8%]",
      tdClassName: "w-[8%]",
      render: (r: Row) => r.approvedDates || (r.approved === "-" || !r.approved ? "" : r.approved),
    },
    {
      key: "expected",
      header: "Expected<br/>Delivery Date",
      thClassName: "w-[8%]",
      tdClassName: "w-[8%] whitespace-nowrap",
      render: (r: Row) => (r.expected === "-" || !r.expected ? "" : r.expected),
    },
    {
      key: "qty",
      header: "Total<br/>Quantity",
      thClassName: "w-[5%] text-left",
      tdClassName: "w-[5%] whitespace-nowrap text-left",
      render: (r: Row) => r.qty.toLocaleString(),
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
      thClassName: "w-[9%] text-left",
      tdClassName: "w-[9%] text-left",
      render: (r: Row) => (
        <div className="flex justify-start">
          <button
            type="button"
            className="w-full max-w-[76px] inline-flex items-center justify-center px-1.5 py-1.5 bg-[#F6F6F6] rounded-md border border-[#EBEBEB] text-[#81AFEB]  text-[9px] font-semibold hover:bg-[#F0F7FF] hover:border hover:border-[#2172DA] hover:text-blue-400"
            onClick={() => props.onDownloadPdf(r.spId)}
          >
            Export .PDF
          </button>
        </div>
      ),
    },
  ];


  return (
    <AllRequestRender
      pageBg={pageBg}
      container={container}
      card={card}
      loadingPage={loadingPage}
      tableLoading={tableLoading}
      isError={isError}
      isBaseEmpty={isBaseEmpty}
      hasAnyData={hasAnyData}
      error={error}
      statusOptions={statusOptions}
      purposeOptions={purposeOptions}
      costCenterOptions={costCenterOptions}
      statusFilter={statusFilter}
      purposeFilter={purposeFilter}
      costCenterFilter={costCenterFilter}
      expectedFrom={expectedFrom}
      expectedTo={expectedTo}
      setStatusFilter={setStatusFilter}
      setPurposeFilter={setPurposeFilter}
      setCostCenterFilter={setCostCenterFilter}
      setExpectedFrom={setExpectedFrom}
      setExpectedTo={setExpectedTo}
      searchDraft={searchDraft}
      manualSelectedCount={manualSelectedIds.length}   // ✅ ADD
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
        setSearchDraft("");
        setSearchText("");
        setSearchField("TransactionID");
        setPageIndex(0);
      }}

      onResetPage={() => setPageIndex(0)}
      pageRows={pageRows}
      filteredRows={sortedRows}
      columns={columns}
      pageSize={pageSize}
      safePageIndex={safePageIndex}
      totalPages={totalPages}
      goFirst={goFirst}
      goPrev={goPrev}
      goNext={goNext}
      goLast={goLast}
      selectedIds={selectedIds}
      handleToggleRow={handleToggleRow}
      handleToggleAll={handleToggleAll}
      onSelectAllRequests={onSelectAllRequests}
      onRestoreManualSelection={onRestoreManualSelection}
      onExportExcel={onExportExcel}
      headerChecked={
        sortedRows.length > 0 &&
        sortedRows.every((r) => selectedIds.includes(String(r.id)))
      }
      canExport={canExport}
      showDisabledExport={props.permissions?.viewExportAllReport === true}
      countIncreased={countIncreased}
      onResetSelection={() => { setSelectedIds([]); setManualSelectedIds([]); }}

    />
  );
}
