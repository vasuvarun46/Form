/* =========================================================
   TeamService.ts
   ✅ Login user fetched from EForm Users list (Email column)
   ✅ Records shown if:
       - CurrentApproverEmail eq loginUserEmail  (they are the current approver)
       - OR loginUserEmail exists in ApprovalChainJSON.allApprovalLevels
         with canApproveRequest = true
   ✅ Action buttons: Approve + Reject only (NO bypass)
   ✅ No approval-matrix / role-based scoping
   ✅ No expand on list name
   ✅ quantityShort boolean + skuMismatch array sent to Power Automate
   ✅ Save-and-Approve (edit qty popup) removed
========================================================= */

import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { HttpClient, SPHttpClient } from "@microsoft/sp-http";

import {
  canonicalizeStatusFilterKey,
  dedupeLabels,
} from "../AllRequestPage/allRequestFilterUtils";
import { RouteState, SecurityPermissions } from "../../spHelpers";
import { ExportActionInfo } from "../ExportActionInfo";
import {
  fetchProductFromAws,
  normalizeFlowToRow,
} from "../CreateRequestPage/CreateRequestService";

/* =========================================================
   TYPES
========================================================= */

export interface Row {
  id: string;
  spId: number;

  status: string;
  requestor: string;
  approver: string;
  approverNames: string; // ✅ NEW: ApproverNames column value

  purpose: string;
  purposeDetail?: string;

  costCenter: string;

  submitted: string;
  approved: string;
  approvedDates?: string;
  expected: string;

  qty: number;
  price: number;

  currentApproverId: number;
  currentApproverEmail: string;
  finalDecisionEmail: string;
  workflowStage: number;

  approvalType: string;
  hasMultipleApprovals: boolean;
  approvalStages: Array<{
    stage: number;
    approverId: number;
    approverJobPosition?: string;
  }>;

  allApprovalLevels: AllApprovalLevel[];

  departmentId?: number;
  departmentTitle?: string;
  modifiedIso?: string;
  lastActivityIso?: string;
}

// ✅ Shape from ApprovalChainJSON.allApprovalLevels
export interface AllApprovalLevel {
  order: number;
  approvalRole: string;
  approverEmail: string;
  approverId: number;
  canApproveRequest: boolean;
  canViewRequest: boolean;
  isCurrentApprover: boolean;
}

export interface GetActionInfoResult {
  canAct: boolean;
  showReject: boolean;
  isApproverWaiting?: boolean;  // in chain but not yet their turn
}

// ✅ EForm Users profile
export interface EFormUserProfile {
  email: string;
  employeeId?: string;
  jobPosition?: string;
  departmentId?: number;
  departmentTitle?: string;
  managerEmail?: string;
  executiveEmail?: string;
  c1Email?: string;
  c2Email?: string;
}

export interface TeamRequestsPageProps {
  context: WebPartContext;
  ApprovalFlowUrl: string;
  productFlowUrl?: string;
  onOpenExportPdf: (spId: number, actionInfo?: ExportActionInfo) => void;
  onDownloadPdf: (spId: number) => void;
  permissions: SecurityPermissions;
  routeState?: RouteState;
  onRouteStateConsumed?: () => void;
}

export type FlowAction = "Approve" | "Reject";

export interface ApprovalFlowPayload {
  // ── Core approval ──────────────────────────────────────────
  requestId: number;    // InventoryRequests SP Id
  actorEmail: string;    // login user email
  position: string;    // Jobposition from EForm Users
  action: FlowAction; // "Approve" | "Reject"
  comments: string;    // reject reason or empty string

  // ── Stock shortage (Approve only; empty/false on Reject) ───
  quantityShort: boolean;       // true if skuMismatches.length > 0
  skuMismatches: SkuMismatch[]; // empty array on Reject or when all stock OK

  // ── Request-level fields for shortage email body ───────────
  transactionId: string; // Title / display TX ID
  detailViewUrl: string; // link to RequestDetail.aspx?requestId=N
  totalQuantity: number; // TotalQuantity column
  totalTagPrice: number; // TotalAmount column
  purpose: string; // Purpose lookup title
  status: string; // WorkflowStatus at time of decision
}

export interface RequestItemRow {
  Id: number;
  SKUCode?: string;
  Quantity?: number;
  UnitPrice?: number;
  RowStatus?: string;
  RequestIDId?: number;
}

export interface ApprovalInvalidItem {
  itemId: number;
  sku: string;
  qty: number;
  unitPrice: number;
  message: string;
}

// ✅ NEW: mismatch shape sent to Power Automate
export interface SkuMismatch {
  itemId: number;
  sku: string;
  tagPricePerUnit: number;      // UnitPrice from SP item (was: requestPrice in schema)
  requestedQty: number;         // Quantity from SP item
  totalTagPrice: number;        // tagPricePerUnit × requestedQty (was: totalRequestPrice)
  availableQty: number;         // from WMS  (-1 = unverifiable)
  actualTotalTagPrice: number;  // tagPricePerUnit × availableQty (was: totalActualPrice)
  differenceQty: number;        // availableQty - requestedQty  (-1 = unverifiable)
}


/* =========================================================
   HELPERS
========================================================= */

const escOdata = (v: string): string => String(v || "").replace(/'/g, "''");

const statusFilterKey = (value: string): string =>
  canonicalizeStatusFilterKey(String(value || "").replace(/-/g, " "));

const isInApprovalStatus = (value: string): boolean =>
  statusFilterKey(value) === statusFilterKey("In Approval");

/** Parse ApprovalChainJSON → allApprovalLevels array */
const parseAllApprovalLevels = (rawJson: unknown): AllApprovalLevel[] => {
  const text = String(rawJson || "").trim();
  if (!text) return [];
  try {
    const parsed = JSON.parse(text) as {
      allApprovalLevels?: AllApprovalLevel[];
    };
    if (!Array.isArray(parsed?.allApprovalLevels)) return [];
    return parsed.allApprovalLevels
      .map((l) => ({
        order: Number(l?.order ?? 0),
        approvalRole: String(l?.approvalRole || "").trim(),
        approverEmail: String(l?.approverEmail || "").trim().toLowerCase(),
        approverId: Number(l?.approverId ?? 0),
        canApproveRequest: l?.canApproveRequest === true,
        canViewRequest: l?.canViewRequest === true,
        isCurrentApprover: l?.isCurrentApprover === true,
      }))
      .filter((l) => l.order > 0 && !!l.approverEmail);
  } catch {
    return [];
  }
};

/** Legacy stages parser kept for backward compat */
const parseApprovalStages = (
  rawJson: unknown
): Array<{ stage: number; approverId: number; approverJobPosition?: string }> => {
  const text = String(rawJson || "").trim();
  if (!text) return [];
  try {
    const parsed = JSON.parse(text) as {
      stages?: Array<{
        stage?: number;
        approverId?: number;
        approverJobPosition?: string;
      }>;
    };
    const stages = Array.isArray(parsed?.stages) ? parsed.stages : [];
    return stages
      .map((s) => ({
        stage: Number(s?.stage ?? 0),
        approverId: Number(s?.approverId ?? 0),
        approverJobPosition: String(s?.approverJobPosition || "").trim() || undefined,
      }))
      .filter((s) => s.stage > 0 && s.approverId > 0)
      .sort((a, b) => a.stage - b.stage);
  } catch {
    return [];
  }
};

function toUserMessage(value: unknown, fallback: string): string {
  if (typeof value === "string" && value.trim()) return value.trim();
  if (value instanceof Error) {
    const msg = String(value.message || "").trim();
    return msg || fallback;
  }
  if (value && typeof value === "object") {
    const obj = value as Record<string, unknown>;
    const nestedError =
      obj.error && typeof obj.error === "object"
        ? (obj.error as Record<string, unknown>)
        : undefined;
    const candidates: unknown[] = [
      obj.message,
      obj.error_description,
      obj.detail,
      obj.title,
      nestedError?.message,
      nestedError?.error_description,
      nestedError?.detail,
      nestedError?.title,
    ];
    for (const candidate of candidates) {
      if (typeof candidate === "string" && candidate.trim())
        return candidate.trim();
    }
  }
  return fallback;
}

async function spGetJson(
  context: WebPartContext,
  url: string
): Promise<any> {
  const res = await context.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`SP GET failed (${res.status}): ${text}`);
  }
  return await res.json();
}

async function spPostJson(
  context: WebPartContext,
  url: string,
  body: any,
  extraHeaders?: Record<string, string>
): Promise<any> {


  const res = await context.httpClient.post(
    url,
    HttpClient.configurations.v1,
    {
      headers: {
        "Content-Type": "application/json",
        ...(extraHeaders || {}),
      },
      body: JSON.stringify(body),
    }
  );

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`SP POST failed (${res.status}): ${text}`);
  }
  try {
    return await res.json();
  } catch {
    const text = await res.text();
    return text ? { message: text } : {};
  }
}

function pickApiKey(configItems: any[]): string | undefined {
  const candidates = [
    "APPROVAL_API_KEY",
    "APPROVE_REJECT_API_KEY",
    "SECUID",
    "SECURITY_ID",
    "X_API_KEY",
    "X-API-KEY",
    "API_KEY",
  ];
  const byKey = new Map<string, string>();
  for (const it of configItems || []) {
    const k = String(it?.ConfigKey || it?.Title || "").trim();
    const v = String(it?.ConfigValue || "").trim();
    if (k) byKey.set(k.toUpperCase(), v);
  }
  for (const c of candidates) {
    const val = byKey.get(c.toUpperCase());
    if (val) return val;
  }
  const firstNonEmpty = (configItems || []).find(
    (x) => String(x?.ConfigValue || "").trim().length > 0
  );
  return firstNonEmpty
    ? String(firstNonEmpty.ConfigValue).trim()
    : undefined;
}

/* =========================================================
   HOOK
========================================================= */

export function useTeamRequestsLogic(
  context: WebPartContext,
  approvalFlowUrl: string,
  searchField: string,
  _permissions?: SecurityPermissions,
  productFlowUrl?: string
) {
  const [rows, setRows] = React.useState<Row[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [tableLoading, setTableLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>("");

  const [searchText, setSearchText] = React.useState<string>("");

  const [statusSelected, setStatusSelected] = React.useState<string[]>([]);
  const [purposeSelected, setPurposeSelected] = React.useState<string[]>([]);
  const [costCenterSelected, setCostCenterSelected] = React.useState<string[]>([]);

  const [expectedFrom, setExpectedFrom] = React.useState<string>("");
  const [expectedTo, setExpectedTo] = React.useState<string>("");

  const [statusOptions, setStatusOptions] = React.useState<string[]>([]);
  const [purposeOptions, setPurposeOptions] = React.useState<string[]>([]);
  const [costCenterOptions, setCostCenterOptions] = React.useState<string[]>([]);

  const [profile, setProfile] = React.useState<EFormUserProfile | null>(null);

  const [flowApiKey, setFlowApiKey] = React.useState<string>("");
  const [hasLoadedOnce, setHasLoadedOnce] = React.useState<boolean>(false);
  const [hasApproverRole, setHasApproverRole] = React.useState<boolean | null>(null);

  const [pageIndex, setPageIndex] = React.useState<number>(0);
  const pageSize: number = 15;

  const siteUrl = context.pageContext.web.absoluteUrl;
  const currentUserEmailRaw = String(
    context.pageContext.user?.email || ""
  ).toLowerCase();
  const currentUserEmailEsc = escOdata(currentUserEmailRaw);
  // ── Storage keys for team request count tracking ──
  const STORAGE_KEY_SUBMITTED = "teamreq_submitted_count";
  const STORAGE_KEY_APPROVED = "teamreq_approved_count";
  const STORAGE_KEY_REJECTED = "teamreq_rejected_count";
  const STORAGE_KEY_IN_APPROVAL = "teamreq_inapproval_count";

  // ── Toggles to force re-render when any count increases ──
  const [countIncreased, setCountIncreased] = React.useState<boolean>(false);

  /* =========================================================
     LOAD PROFILE from EForm Users
  ========================================================= */

  const loadEFormUserProfile =
    React.useCallback(async (): Promise<EFormUserProfile> => {
      const url =
        `${siteUrl}/_api/web/lists/getbytitle('EForm%20Users')/items` +
        `?$select=Id,Title,EmployeeID,Email,ManagerEmail,ExecutiveEmail,C1Email,C2Email,Jobposition,Department/Id,Department/Title` +
        `&$expand=Department` +
        `&$filter=Email eq '${currentUserEmailEsc}'` +
        `&$top=1`;

      const json = await spGetJson(context, url);
      const it = json?.value?.[0];

      return {
        email: currentUserEmailRaw,
        employeeId: it?.EmployeeID ? String(it.EmployeeID) : undefined,
        jobPosition: it?.Jobposition ? String(it.Jobposition) : undefined,
        departmentId: it?.Department?.Id ? Number(it.Department.Id) : undefined,
        departmentTitle: it?.Department?.Title ? String(it.Department.Title) : undefined,
        managerEmail: it?.ManagerEmail ? String(it.ManagerEmail).trim().toLowerCase() : undefined,
        executiveEmail: it?.ExecutiveEmail ? String(it.ExecutiveEmail).trim().toLowerCase() : undefined,
        c1Email: it?.C1Email ? String(it.C1Email).trim().toLowerCase() : undefined,
        c2Email: it?.C2Email ? String(it.C2Email).trim().toLowerCase() : undefined,
      };
    }, [context, siteUrl, currentUserEmailRaw, currentUserEmailEsc]);

  /* =========================================================
     CHECK APPROVER ROLE
     Verifies the current user appears as ManagerEmail, ExecutiveEmail,
     C1Email, or C2Email for at least one entry in EForm Users.
  ========================================================= */

  const checkApproverRole = React.useCallback(async (): Promise<boolean> => {
    const url =
      `${siteUrl}/_api/web/lists/getbytitle('EForm%20Users')/items` +
      `?$select=Id` +
      `&$filter=ManagerEmail eq '${currentUserEmailEsc}' or ExecutiveEmail eq '${currentUserEmailEsc}' or C1Email eq '${currentUserEmailEsc}' or C2Email eq '${currentUserEmailEsc}'` +
      `&$top=1`;
    const json = await spGetJson(context, url);
    return Array.isArray(json?.value) && json.value.length > 0;
  }, [context, siteUrl, currentUserEmailEsc]);

  /* =========================================================
     LOAD CONFIG
  ========================================================= */

  const loadWorkflowConfig = React.useCallback(async (): Promise<string> => {
    const url =
      `${siteUrl}/_api/web/lists/getbytitle('WorkflowConfiguration')/items` +
      `?$select=Id,Title,ConfigKey,ConfigValue,IsActive` +
      `&$filter=IsActive eq 1`;

    const json = await spGetJson(context, url);
    const key = pickApiKey(json?.value || []);
    return key || "";
  }, [context, siteUrl]);

  const loadStatusChoices = React.useCallback(async (): Promise<void> => {
    setStatusOptions(["Approved", "Rejected", "Submitted"]);
  }, []);

  /* =========================================================
     LOAD MASTER OPTIONS
  ========================================================= */

  const loadMasterOptions = React.useCallback(
    async (p: EFormUserProfile): Promise<void> => {
      const purposeUrl =
        `${siteUrl}/_api/web/lists/getbytitle('PurposeMaster')/items` +
        `?$select=Id,Title,IsActive` +
        `&$filter=IsActive eq 1` +
        `&$orderby=Title asc`;

      const hasDepth = !!p.departmentId;
      const ccFilter = hasDepth
        ? `IsActive eq 1 and Department/Id eq ${p.departmentId}`
        : `IsActive eq 1`;

      const costCenterUrl =
        `${siteUrl}/_api/web/lists/getbytitle('CostCenterMaster')/items` +
        `?$select=Id,Title,IsActive,Department/Id,Department/Title` +
        `&$expand=Department` +
        `&$filter=${ccFilter}` +
        `&$orderby=Title asc`;

      const [purposeJson, ccJson] = await Promise.all([
        spGetJson(context, purposeUrl),
        spGetJson(context, costCenterUrl),
      ]);

      const purposeTitles: string[] = (purposeJson?.value || [])
        .map((x: any) => String(x?.Title || "").trim())
        .filter(Boolean);
      setPurposeOptions(dedupeLabels(purposeTitles));

      const ccTitles: string[] = (ccJson?.value || [])
        .map((x: any) => String(x?.Title || "").trim())
        .filter(Boolean);
      setCostCenterOptions(dedupeLabels(ccTitles));
    },
    [context, siteUrl]
  );
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

  /* =========================================================
     LOAD INVENTORY REQUESTS
  ========================================================= */

  const loadInventoryRequests = React.useCallback(
    async (_p: EFormUserProfile): Promise<void> => {
      const clauses: string[] = [
        `WorkflowStatus ne 'Draft' and WorkflowStatus ne 'Pending'`,
      ];

      if (statusSelected.length > 0) {
        const statusTerms: string[] = [];
        statusSelected.forEach((status) => {
          const key = statusFilterKey(status);
          if (
            key === statusFilterKey("Draft") ||
            key === statusFilterKey("Pending")
          )
            return;
          if (isInApprovalStatus(status)) {
            statusTerms.push(`WorkflowStatus eq 'Submitted'`);
            statusTerms.push(`WorkflowStatus eq 'In Approval'`);
            statusTerms.push(`WorkflowStatus eq 'In-Approval'`);
          } else {
            statusTerms.push(`WorkflowStatus eq '${escOdata(status)}'`);
          }
        });
        if (statusTerms.length > 0) {
          clauses.push(`(${Array.from(new Set(statusTerms)).join(" or ")})`);
        }
      }

      if (purposeSelected.length > 0) {
        const purposeTerms = purposeSelected.map(
          (x) => `Purpose/Title eq '${escOdata(x)}'`
        );
        clauses.push(`(${purposeTerms.join(" or ")})`);
      }

      if (costCenterSelected.length > 0) {
        const ccTerms = costCenterSelected.map(
          (x) => `CostCenter/Title eq '${escOdata(x)}'`
        );
        clauses.push(`(${ccTerms.join(" or ")})`);
      }

      if (expectedFrom) {
        clauses.push(`ExpectedDeliveryDate ge datetime'${expectedFrom}T00:00:00Z'`);
      }
      if (expectedTo) {
        clauses.push(`ExpectedDeliveryDate le datetime'${expectedTo}T23:59:59Z'`);
      }

      if (searchText.trim()) {
        const q = escOdata(searchText.trim());
        if (searchField === "RequesterName") {
          clauses.push(`substringof('${q}',Requester/Title)`);
        } else {
          clauses.push(`substringof('${q}',Title)`);
        }
      }

      const filterClause = clauses.join(" and ");

      const url =
        `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items` +
        `?$select=` +
        `Id,Title,RequestNumber,WorkflowStatus,WorkflowStage,` +
        `ApprovalChainJSON,ApprovalType,HasMultipleApprovals,` +
        `CurrentApproverEmail,FinalDecisionEmail,ApproverNames,` +
        `RequestSubmittedOn,FinalDecisionDate,ApprovedDates,ExpectedDeliveryDate,` +
        `TotalQuantity,TotalAmount,Modified,MultipleApproverData,` +
        `Requester/Title,Requester/EMail,` +
        `CurrentApprover/Id,CurrentApprover/Title,CurrentApprover/EMail,` +
        `Purpose/Title,PurposeDetails,CostCenter/Title,` +
        `Department/Id,Department/Title` +
        `&$expand=Requester,CurrentApprover,Purpose,CostCenter,Department` +
        `&$filter=${filterClause}` +
        `&$orderby=Id desc`;

      const json = await spGetJson(context, url);

      // Fallback: Manager → Executive/Executed → C1 → C2
      const deriveCurrentApproverEmail = (
        spEmail: string,
        levels: AllApprovalLevel[]
      ): string => {
        if (spEmail) return spEmail;
        const roleAliases: Array<[string[]]> = [
          [["manager"]],
          [["executive", "executed"]],
          [["c1"]],
          [["c2"]],
        ];
        const sorted = [...levels].sort((a, b) => a.order - b.order);
        for (const [aliases] of roleAliases) {
          const match = sorted.find(
            (l) =>
              aliases.some((a) => l.approvalRole.toLowerCase() === a) &&
              !!l.approverEmail
          );
          if (match?.approverEmail) return match.approverEmail;
        }
        return "";
      };

      const getActedApproverName = (
        raw: unknown,
        workflowStatus: string
      ): string | null => {
        const status = workflowStatus.trim().toLowerCase();
        if (status !== "approved" && status !== "rejected") return null;
        if (!raw) return null;
        let parsed: unknown = raw;
        if (typeof raw === "string") {
          try { parsed = JSON.parse(raw); } catch { return null; }
        }
        if (!parsed || typeof parsed !== "object") return null;
        const p2 = parsed as { stages?: Array<Record<string, unknown>> };
        const stages = Array.isArray(p2.stages) ? p2.stages : [];
        const actedStage = stages.find(
          (s) => !!String(s.Date ?? s.date ?? "").trim()
        );
        if (!actedStage) return null;
        return (
          String(actedStage.approverName ?? actedStage.ApproverName ?? "").trim() || null
        );
      };

      const mapped: Row[] = (json?.value || []).map((it: any) => {
        const approvalStages = parseApprovalStages(it?.ApprovalChainJSON);
        const allApprovalLevels = parseAllApprovalLevels(it?.ApprovalChainJSON);

        const approvalType = String(it?.ApprovalType || "").trim();
        const hasMultipleApprovals = it?.HasMultipleApprovals === true;
        const workflowStatus = String(it?.WorkflowStatus || "");

        const reqNum =
          it?.RequestNumber !== null && it?.RequestNumber !== undefined
            ? String(it.RequestNumber)
            : "";

        const isPending = workflowStatus.trim().toLowerCase() === "pending";
        const titleTxId = String(it?.Title || "").trim();
        const hasFinalTitle =
          !!titleTxId &&
          !/^draft$/i.test(titleTxId) &&
          !/^temp$/i.test(titleTxId);

        const txId = isPending
          ? "-"
          : hasFinalTitle
            ? titleTxId
            : reqNum && reqNum !== "0"
              ? reqNum
              : String(it?.Id || "");

        const currentApproverEmail = deriveCurrentApproverEmail(
          String(it?.CurrentApproverEmail || "").trim().toLowerCase(),
          allApprovalLevels
        );

        const finalDecisionEmail = String(it?.FinalDecisionEmail || "")
          .trim()
          .toLowerCase();

        // ✅ ApproverNames plain text column
        const approverNames = String(it?.ApproverNames || "").trim();

        return {
          id: txId,
          spId: it.Id,

          status: workflowStatus,
          requestor: String(it?.Requester?.Title || ""),
          approver:
            getActedApproverName(it?.MultipleApproverData, workflowStatus) ??
            String(it?.CurrentApprover?.Title || ""),
          approverNames, // ✅ NEW

          purpose: String(it?.Purpose?.Title || ""),
          purposeDetail: it?.PurposeDetails ? String(it.PurposeDetails) : undefined,

          costCenter: String(it?.CostCenter?.Title || ""),

          submitted: formatDateDdMmYyyy(it?.RequestSubmittedOn),
          approved: formatDateDdMmYyyy(it?.FinalDecisionDate),
          approvedDates: it?.ApprovedDates ? it.ApprovedDates.split(",").map((d: string) => { const s = d.trim(); if (!s) return ""; const p = s.split("-"); if (p.length === 3 && p[0].length === 4) return `${p[2]}-${p[1]}-${p[0]}`; return s.replace(/\//g, "-"); }).filter(Boolean).join(", ") : "",
          expected: formatDateDdMmYyyy(it?.ExpectedDeliveryDate),

          qty: Number(it?.TotalQuantity || 0),
          price: Number(it?.TotalAmount || 0),

          currentApproverId: Number(it?.CurrentApprover?.Id || 0),
          currentApproverEmail,
          finalDecisionEmail,

          workflowStage: Number(it?.WorkflowStage || 0),

          approvalType,
          hasMultipleApprovals,
          approvalStages,
          allApprovalLevels,

          departmentId: it?.Department?.Id ? Number(it.Department.Id) : undefined,
          departmentTitle: it?.Department?.Title ? String(it.Department.Title) : undefined,

          modifiedIso: String(it?.Modified || ""),
          lastActivityIso: String(it?.Modified || it?.RequestSubmittedOn || ""),
        };
      });

      const loginEmail = currentUserEmailRaw;

      const visibleRows = mapped.filter((row) => {
        const key = statusFilterKey(row.status);
        const isLifecycleVisible =
          key === statusFilterKey("In Approval") ||
          key === statusFilterKey("Submitted") ||
          key === statusFilterKey("Approved") ||
          key === statusFilterKey("Rejected");

        if (!isLifecycleVisible) return false;

        if (row.currentApproverEmail && row.currentApproverEmail === loginEmail)
          return true;

        return row.allApprovalLevels.some(
          (l) => l.approverEmail === loginEmail && l.canViewRequest === true
        );
      });

      setRows(visibleRows);
    },
    [
      context,
      siteUrl,
      currentUserEmailRaw,
      statusSelected,
      purposeSelected,
      costCenterSelected,
      expectedFrom,
      expectedTo,
      searchText,
      searchField,
    ]
  );

  /* =========================================================
     LOAD ALL
  ========================================================= */

  const loadAll = React.useCallback(async (silent: boolean = false): Promise<void> => {

    try {

      if (!silent) {
        if (!hasLoadedOnce) {
          setLoading(true);
        } else {
          setTableLoading(true);
        }
      }
      setError("");

      let p = profile;
      let apiKey = flowApiKey;

      if (!p && !apiKey) {
        const [loadedProfile, loadedApiKey] = await Promise.all([
          loadEFormUserProfile(),
          loadWorkflowConfig(),
        ]);
        p = loadedProfile;
        apiKey = loadedApiKey;
        setProfile(loadedProfile);
        setFlowApiKey(loadedApiKey);
      } else {
        if (!p) {
          p = await loadEFormUserProfile();
          setProfile(p);
        }
        if (!apiKey) {
          apiKey = await loadWorkflowConfig();
          setFlowApiKey(apiKey);
        }
      }

      // ── Role gate: only users who are a Manager / Executive / C1 / C2
      //    for at least one employee may see Team Request records.
      let isApprover = hasApproverRole;
      if (isApprover === null) {
        isApprover = await checkApproverRole();
        setHasApproverRole(isApprover);
      }

      if (!isApprover) {
        setRows([]);
        return;
      }

      if (
        statusOptions.length === 0 ||
        purposeOptions.length === 0 ||
        costCenterOptions.length === 0
      ) {
        await Promise.all([loadMasterOptions(p), loadStatusChoices()]);
      }

      await loadInventoryRequests(p);
    } catch (e: unknown) {
      setError(toUserMessage(e, "Failed to load team requests"));
    } finally {
      if (!silent) {
        setLoading(false);
        setTableLoading(false);
      }
      setHasLoadedOnce(true);
    }
  }, [
    profile,
    flowApiKey,
    hasLoadedOnce,
    hasApproverRole,
    statusOptions.length,
    purposeOptions.length,
    costCenterOptions.length,
    loadEFormUserProfile,
    loadWorkflowConfig,
    checkApproverRole,
    loadMasterOptions,
    loadStatusChoices,
    loadInventoryRequests,
  ]);

  /* =========================================================
     PAGINATION
  ========================================================= */

  const filteredRows = React.useMemo(() => [...rows], [rows]);

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));

  const safePageIndex =
    pageIndex >= totalPages ? totalPages - 1 : pageIndex < 0 ? 0 : pageIndex;

  const goFirst = () => setPageIndex(0);
  const goPrev = () => setPageIndex((p) => Math.max(0, p - 1));
  const goNext = () => setPageIndex((p) => Math.min(totalPages - 1, p + 1));
  const goLast = () => setPageIndex(totalPages - 1);

  React.useEffect(() => {
    setPageIndex(0);
  }, [
    searchText,
    statusSelected,
    purposeSelected,
    costCenterSelected,
    expectedFrom,
    expectedTo,
  ]);

  /* =========================================================
     ACTION INFO
  ========================================================= */

  const getActionInfo = React.useCallback(
    (row: Row): GetActionInfoResult => {
      const normalizedStatus = String(row.status || "").trim().toLowerCase();

      if (normalizedStatus === "approved" || normalizedStatus === "rejected") {
        return { canAct: false, showReject: false };
      }

      const isActionableState =
        normalizedStatus === "submitted" ||
        normalizedStatus === "in approval" ||
        normalizedStatus === "in-approval";

      if (!isActionableState) {
        return { canAct: false, showReject: false };
      }

      if (row.finalDecisionEmail && row.finalDecisionEmail.trim().length > 0) {
        return { canAct: false, showReject: false };
      }

      const loginEmail = currentUserEmailRaw;

      const isSingleApprover = !row.hasMultipleApprovals;
      const isMultipleApprover = !!row.hasMultipleApprovals;

      if (isSingleApprover) {
        const canApprove = row.allApprovalLevels.some(
          (l) =>
            l.approverEmail === loginEmail &&
            l.canApproveRequest === true
        );
        if (canApprove) return { canAct: true, showReject: true };

        const isInChainWaiting = row.allApprovalLevels.some(
          (l) => l.approverEmail === loginEmail && l.canApproveRequest === false
        );
        return { canAct: false, showReject: false, isApproverWaiting: isInChainWaiting };
      }

      if (isMultipleApprover) {
        const isCurrentApprover =
          !!row.currentApproverEmail &&
          row.currentApproverEmail === loginEmail;
        if (isCurrentApprover) return { canAct: true, showReject: true };

        const isInChainWaiting = row.allApprovalLevels.some(
          (l) => l.approverEmail === loginEmail
        );
        return { canAct: false, showReject: false, isApproverWaiting: isInChainWaiting };
      }

      return { canAct: false, showReject: false };
    },
    [currentUserEmailRaw]
  );

  /* =========================================================
     VALIDATE ITEMS BEFORE APPROVE
     ✅ Now returns skuMismatches[] and quantityShort boolean
        alongside the existing invalidItems for UI display.
  ========================================================= */

  const getActiveRequestItems = React.useCallback(
    async (requestId: number): Promise<RequestItemRow[]> => {
      const url =
        `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items` +
        `?$top=5000` +
        `&$select=Id,SKUCode,Quantity,UnitPrice,RequestIDId,RowStatus` +
        `&$filter=RequestIDId eq ${requestId} and RowStatus eq 'Active'` +
        `&$orderby=LineNumber asc`;
      const json = await spGetJson(context, url);
      return (json?.value || []) as RequestItemRow[];
    },
    [context, siteUrl]
  );


  const validateRequestItemsBeforeApprove = React.useCallback(
    async (row: Row): Promise<{
      ok: boolean;
      invalidItems: ApprovalInvalidItem[];
      skuMismatches: SkuMismatch[];
      quantityShort: boolean;
    }> => {
      const requestId = Number(row.spId || 0);
      if (!requestId) {
        return {
          ok: false,
          invalidItems: [{ itemId: 0, sku: "-", qty: 0, unitPrice: 0, message: "Invalid request id." }],
          skuMismatches: [],
          quantityShort: false,
        };
      }

      const items = await getActiveRequestItems(requestId);
      const invalidItems: ApprovalInvalidItem[] = [];
      const skuMismatches: SkuMismatch[] = [];

      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        const reqQty = Number(it.Quantity || 0);
        const unitPrice = Number(it.UnitPrice || 0);       // tagPricePerUnit
        const sku = String(it.SKUCode || "-").trim() || "-";
        const itemId = Number(it.Id || 0);
        const totalTagPrice = reqQty * unitPrice;           // requestedQty × tagPricePerUnit

        if (productFlowUrl) {
          try {
            const json = await fetchProductFromAws(context.httpClient, sku, productFlowUrl);
            const flow = normalizeFlowToRow(json);
            const liveStock = Number(flow?.Warehouse_SOH_Qty ?? 0);

            if (!flow || liveStock <= 0) {
              // ── OUT OF STOCK ────────────────────────────────────────────
              invalidItems.push({
                itemId,
                sku,
                qty: 0,
                unitPrice,
                message: `SKU "${sku}" is out of stock.`,
              });
              skuMismatches.push({
                itemId,
                sku,
                tagPricePerUnit: unitPrice,
                requestedQty: reqQty,
                totalTagPrice,                 // requestedQty × tagPricePerUnit
                availableQty: 0,
                actualTotalTagPrice: 0,        // 0 stock → 0 actual total
                differenceQty: 0 - reqQty, // always negative
              });

            } else if (reqQty > liveStock) {
              // ── PARTIAL STOCK ───────────────────────────────────────────
              invalidItems.push({
                itemId,
                sku,
                qty: liveStock,
                unitPrice,
                message: `Only ${liveStock} pcs available. Requested: ${reqQty}.`,
              });
              skuMismatches.push({
                itemId,
                sku,
                tagPricePerUnit: unitPrice,
                requestedQty: reqQty,
                totalTagPrice,                       // original total at requested qty
                availableQty: liveStock,
                actualTotalTagPrice: liveStock * unitPrice, // actual total at available qty
                differenceQty: liveStock - reqQty,    // negative = shortfall
              });

            } else {
              // ── QTY OK — sufficient stock, no entry needed ──────────────────
            }

          } catch {
            // ── UNVERIFIABLE (API call failed) ──────────────────────────────
            invalidItems.push({
              itemId,
              sku,
              qty: reqQty,
              unitPrice,
              message: `Unable to verify stock for SKU "${sku}".`,
            });
            skuMismatches.push({
              itemId,
              sku,
              tagPricePerUnit: unitPrice,
              requestedQty: reqQty,
              totalTagPrice,
              availableQty: -1,  // -1 sentinel = unverifiable
              actualTotalTagPrice: -1,
              differenceQty: -1,
            });
          }

        } else {
          // ── NO productFlowUrl — basic qty range check only ──────────────
          if (!Number.isFinite(reqQty) || reqQty < 1 || reqQty > 999) {
            invalidItems.push({
              itemId,
              sku,
              qty: reqQty,
              unitPrice,
              message: "Quantity must be between 1 and 999.",
            });
          }
          // No WMS URL — cannot verify stock, no mismatch entry added.
        }
      }

      // If no SKU had a shortage, send a single zero-row so Power Automate
      // always receives at least one entry in skuMismatches.
      if (skuMismatches.length === 0) {
        skuMismatches.push({
          itemId: 0,
          sku: "0",
          tagPricePerUnit: 0,
          requestedQty: 0,
          totalTagPrice: 0,
          availableQty: 0,
          actualTotalTagPrice: 0,
          differenceQty: 0,
        });
      }

      const quantityShort = true;

      return {
        ok: invalidItems.length === 0,
        invalidItems,
        skuMismatches,
        quantityShort,
      };
    },
    [getActiveRequestItems, context, productFlowUrl]
  );


  /* =========================================================
     UPDATE REQUEST ITEM QUANTITY  (kept — still used elsewhere)
  ========================================================= */

  const updateRequestItemQuantity = React.useCallback(
    async (
      requestId: number,
      itemId: number,
      nextQty: number
    ): Promise<{ totalQty: number; totalAmount: number }> => {
      const qty = Math.trunc(Number(nextQty || 0));
      if (!Number.isFinite(qty) || qty < 1 || qty > 999) {
        throw new Error("Quantity must be between 1 and 999.");
      }

      const itemUrl =
        `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items(${itemId})` +
        `?$select=Id,UnitPrice`;
      const itemJson = (await spGetJson(context, itemUrl)) as { UnitPrice?: number };
      const unitPrice = Number(itemJson?.UnitPrice || 0);
      const lineTotal = qty * unitPrice;

      await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify({ Quantity: qty, LineTotal: lineTotal }),
        }
      );

      const allItems = await getActiveRequestItems(requestId);
      const totalQty = allItems.reduce((sum, x) => sum + Number(x.Quantity || 0), 0);
      const totalAmount = allItems.reduce(
        (sum, x) => sum + Number(x.Quantity || 0) * Number(x.UnitPrice || 0),
        0
      );

      await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${requestId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify({
            TotalQuantity: totalQty,
            TotalAmount: totalAmount,
            LastUpdatedOn: new Date().toISOString(),
          }),
        }
      );

      setRows((prev) =>
        prev.map((r) =>
          r.spId === requestId
            ? { ...r, qty: totalQty, price: totalAmount, lastActivityIso: new Date().toISOString() }
            : r
        )
      );

      return { totalQty, totalAmount };
    },
    [context, getActiveRequestItems, siteUrl]
  );

  /* =========================================================
     SUBMIT DECISION
     ✅ Now validates stock first, then sends quantityShort +
        skuMismatches[] in the Power Automate payload.
     ✅ Proceeds to approve even when quantityShort = true
        (Save-and-Approve edit popup removed).
  ========================================================= */

  const submitDecision = React.useCallback(
    async (
      row: Row,
      action: "Approve" | "Reject",
      comments: string
    ): Promise<{ success: boolean; message: string; newStatus?: string }> => {
      if (!approvalFlowUrl) throw new Error("Approval flow URL is missing");
      if (!flowApiKey)
        throw new Error("Security key (secret) not found in WorkflowConfiguration");

      // ── STOCK VALIDATION (Approve only) ────────────────────────────────
      let quantityShort = false;
      let skuMismatches: SkuMismatch[] = [];
      let finalTotalQty = row.qty;
      let finalTotalAmount = row.price;

      if (action === "Approve") {
        const check = await validateRequestItemsBeforeApprove(row);
        quantityShort = check.quantityShort;
        skuMismatches = check.skuMismatches;

        // ── FINAL APPROVER DETECTION ────────────────────────────────────
        // The final approver is the one at the highest order in the approval
        // chain. Only the final approver actually writes adjusted quantities
        // to SharePoint — intermediate approvers just pass mismatch data along.
        const loginEmail = String(currentUserEmailRaw || "").trim().toLowerCase();
        const activeLevels = row.allApprovalLevels.filter(
          (l) => l.canApproveRequest === true
        );
        let isFinalApprover = false;
        if (activeLevels.length > 0) {
          const maxOrder = Math.max(...activeLevels.map((l) => l.order));
          isFinalApprover = activeLevels.some(
            (l) => l.approverEmail === loginEmail && l.order === maxOrder
          );
        } else if (!row.hasMultipleApprovals) {
          // Single-approver record with no level data — they are always final
          isFinalApprover = true;
        }

        // ── UPDATE SP QUANTITIES (Final Approver only, verifiable mismatches) ─
        // All SharePoint writes must complete before Power Automate is called.
        if (isFinalApprover && quantityShort && skuMismatches.length > 0) {
          const requestId = Number(row.spId || 0);

          for (const mismatch of skuMismatches) {
            // Skip unverifiable items (availableQty === -1 sentinel)
            if (mismatch.availableQty === -1) continue;
            const approvedQty = Math.max(0, mismatch.availableQty);
            await context.spHttpClient.post(
              `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items(${mismatch.itemId})`,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept: "application/json;odata=nometadata",
                  "Content-Type": "application/json;odata=nometadata",
                  "odata-version": "",
                  "IF-MATCH": "*",
                  "X-HTTP-Method": "MERGE",
                },
                body: JSON.stringify({
                  Quantity: approvedQty,
                  LineTotal: approvedQty * mismatch.tagPricePerUnit,
                }),
              }
            );
          }

          // Re-fetch all active items to calculate accurate updated totals
          const updatedItems = await getActiveRequestItems(requestId);
          finalTotalQty = updatedItems.reduce(
            (sum, x) => sum + Number(x.Quantity || 0),
            0
          );
          finalTotalAmount = updatedItems.reduce(
            (sum, x) => sum + Number(x.Quantity || 0) * Number(x.UnitPrice || 0),
            0
          );

          // Persist new totals to InventoryRequests before triggering PA
          await context.spHttpClient.post(
            `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${requestId})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "Content-Type": "application/json;odata=nometadata",
                "odata-version": "",
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE",
              },
              body: JSON.stringify({
                TotalQuantity: finalTotalQty,
                TotalAmount: finalTotalAmount,
              }),
            }
          );
        }
      }

      // ── DETAIL VIEW URL ────────────────────────────────────────────────
      const siteOrigin = String(context.pageContext.web.absoluteUrl || "").replace(/\/$/, "");
      const detailViewUrl = `${siteOrigin}/SitePages/RequestDetail.aspx?requestId=${row.spId}`;

      // ── PAYLOAD ────────────────────────────────────────────────────────
      // totalQuantity / totalTagPrice carry the post-adjustment values so
      // Power Automate emails show the final approved figures, not the
      // originally requested ones.
      const payload: ApprovalFlowPayload = {
        requestId: row.spId,
        actorEmail: currentUserEmailRaw,
        position: String(profile?.jobPosition || "Unknown"),
        action: action as FlowAction,
        comments: comments || "",
        quantityShort,
        skuMismatches,
        transactionId: row.id,
        detailViewUrl,
        totalQuantity: finalTotalQty,
        totalTagPrice: finalTotalAmount,
        purpose: row.purpose,
        status: row.status,
      };

      // ── CALL POWER AUTOMATE ────────────────────────────────────────────
      const resp = await spPostJson(context, approvalFlowUrl, payload, {
        "x-api-key": flowApiKey,
      });

      const ok =
        typeof resp?.status === "boolean"
          ? resp.status
          : String(resp?.status || "").toLowerCase() === "true"
            ? true
            : Boolean(resp?.success);

      const msg = toUserMessage(resp?.message, ok ? "Success" : "Failed");
      const newStatus = resp?.newStatus ? String(resp.newStatus) : undefined;

      // ── UPDATE LOCAL ROW STATE on success ──────────────────────────────
      if (ok && newStatus) {
        setRows((prev) =>
          prev.map((r) =>
            r.spId === row.spId
              ? {
                  ...r,
                  status: newStatus,
                  qty: finalTotalQty,
                  price: finalTotalAmount,
                  approver: "",
                  currentApproverId: 0,
                  currentApproverEmail: "",
                  lastActivityIso: new Date().toISOString(),
                }
              : r
          )
        );
      }

      return { success: ok, message: msg, newStatus };
    },
    [
      context,
      approvalFlowUrl,
      flowApiKey,
      currentUserEmailRaw,
      profile,
      validateRequestItemsBeforeApprove,
      getActiveRequestItems,
      siteUrl,
    ]
  );


  /* =========================================================
     INITIAL LOAD
  ========================================================= */

  React.useEffect(() => {
    loadAll(false).catch(() => undefined);
  }, [loadAll]);


  /* ── Lightweight fetch — counts Submitted, Approved, Rejected, In Approval
      scoped to records visible to this approver ── */
  const getStatusCounts = React.useCallback(async (): Promise<{
    submitted: number;
    approved: number;
    rejected: number;
    inApproval: number;
  }> => {
    const buildUrl = (status: string): string =>
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items` +
      `?$select=Id,CurrentApproverEmail,ApprovalChainJSON` +
      `&$filter=WorkflowStatus eq '${escOdata(status)}'` +
      `&$top=5000`;

    const fetchCount = async (status: string): Promise<number> => {
      const res = await context.spHttpClient.get(
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

      const loginEmail = currentUserEmailRaw;
      const visible = (data?.value || []).filter((it: any) => {
        const currentApproverEmail = String(it?.CurrentApproverEmail || "")
          .trim()
          .toLowerCase();

        if (currentApproverEmail === loginEmail) return true;

        const levels = parseAllApprovalLevels(it?.ApprovalChainJSON);
        return levels.some(
          (l) => l.approverEmail === loginEmail && l.canViewRequest === true
        );
      });

      return visible.length;
    };

    const [submitted, approved, rejected, inApproval] = await Promise.all([
      fetchCount("Submitted"),
      fetchCount("Approved"),
      fetchCount("Rejected"),
      fetchCount("In Approval"),
    ]);

    return { submitted, approved, rejected, inApproval };
  }, [context, siteUrl, currentUserEmailRaw]);

  /* ── Compare counts with localStorage, reload silently if changed ── */
  const checkAndUpdateCounts = React.useCallback(
    async (isInitialLoad: boolean): Promise<void> => {
      try {
        const {
          submitted: currentSubmitted,
          approved: currentApproved,
          rejected: currentRejected,
          inApproval: currentInApproval,
        } = await getStatusCounts();

        const prevSubmitted = Number(localStorage.getItem(STORAGE_KEY_SUBMITTED) ?? 0);
        const prevApproved = Number(localStorage.getItem(STORAGE_KEY_APPROVED) ?? 0);
        const prevRejected = Number(localStorage.getItem(STORAGE_KEY_REJECTED) ?? 0);
        const prevInApproval = Number(localStorage.getItem(STORAGE_KEY_IN_APPROVAL) ?? 0);

        const hasChanged =
          !isInitialLoad &&
          (
            currentSubmitted > prevSubmitted ||
            currentApproved > prevApproved ||
            currentRejected > prevRejected ||
            currentInApproval < prevInApproval  // decrease = request was approved or rejected
          );

        if (hasChanged) {
          setCountIncreased((prev) => !prev);
          await loadAll(true);
        }

        localStorage.setItem(STORAGE_KEY_SUBMITTED, String(currentSubmitted));
        localStorage.setItem(STORAGE_KEY_APPROVED, String(currentApproved));
        localStorage.setItem(STORAGE_KEY_REJECTED, String(currentRejected));
        localStorage.setItem(STORAGE_KEY_IN_APPROVAL, String(currentInApproval));

      } catch (err) {
        console.error("[TeamService] checkAndUpdateCounts error:", err);
      }
    },
    [getStatusCounts, loadAll]
  );
  React.useEffect(() => {
    checkAndUpdateCounts(true).catch(console.error);  // initial save, no re-render

    const intervalId = setInterval(() => {
      checkAndUpdateCounts(false).catch(console.error); // poll — re-renders if count increased
    }, 2 * 60 * 1000);

    return () => clearInterval(intervalId);
  }, [checkAndUpdateCounts]);

  return {
    loading,
    tableLoading,
    error,
    profile,

    filteredRows,
    pageSize,
    safePageIndex,
    totalPages,
    goFirst,
    goPrev,
    goNext,
    goLast,

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

    statusOptions,
    purposeOptions,
    costCenterOptions,

    getActionInfo,
    validateRequestItemsBeforeApprove,
    updateRequestItemQuantity,
    submitDecision,

    reload: loadAll,
    rows,
    countIncreased,
    hasApproverRole,
  };
}