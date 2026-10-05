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
  requesterEmail: string;
  approver: string;
  approverNames: string; // ✅ NEW: ApproverNames column value
  multipleApproverData?: string;

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
  jobTitle?: string;
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
  onApproverRoleResolved?: (hasRole: boolean) => void;
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
  // liveStockQty is internal-only — excluded here because the Power Automate trigger
  // schema rejects unknown properties (additionalProperties: false).
  skuMismatches: Omit<SkuMismatch, "liveStockQty">[]; // empty on Reject or when all stock OK

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
  // Actual live warehouse stock (Warehouse_SOH_Qty) from the API — used for the
  // StockAvailability column. Distinct from availableQty (which drives ApprovalCount).
  liveStockQty: number;
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
        `?$select=Id,Title,Email,ManagerEmail,ExecutiveEmail,C1Email,C2Email` +
        `&$filter=Email eq '${currentUserEmailEsc}'` +
        `&$top=1`;

      const json = await spGetJson(context, url);
      const it = json?.value?.[0];

      let jobTitle = "";
      try {
        const graphClient = await context.msGraphClientFactory.getClient("3");
        const me: { jobTitle?: string } = await graphClient.api("/me").select("jobTitle").get();
        jobTitle = String(me.jobTitle || "");
      } catch { /* ignore — position will be empty if Graph is unavailable */ }

      return {
        email: currentUserEmailRaw,
        jobTitle,
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
    // Check 1: user is listed as a manager/executive/C1/C2 for any employee
    // Wrapped in try/catch — column names may differ across tenants
    try {
      const eformUrl =
        `${siteUrl}/_api/web/lists/getbytitle('EForm%20Users')/items` +
        `?$select=Id` +
        `&$filter=ManagerEmail eq '${currentUserEmailEsc}' or ExecutiveEmail eq '${currentUserEmailEsc}' or C1Email eq '${currentUserEmailEsc}' or C2Email eq '${currentUserEmailEsc}'` +
        `&$top=1`;
      const eformJson = await spGetJson(context, eformUrl);
      if (Array.isArray(eformJson?.value) && eformJson.value.length > 0) return true;
    } catch { /* field names may differ — fall through to check 2 */ }

    // Check 2: user is the current approver of any active request
    try {
      const activeUrl =
        `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items` +
        `?$select=Id` +
        `&$filter=CurrentApproverEmail eq '${currentUserEmailEsc}' and WorkflowStatus ne 'Draft' and WorkflowStatus ne 'Pending'` +
        `&$top=1`;
      const activeJson = await spGetJson(context, activeUrl);
      return Array.isArray(activeJson?.value) && activeJson.value.length > 0;
    } catch {
      return false;
    }
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
    async (_p: EFormUserProfile): Promise<void> => {
      const purposeUrl =
        `${siteUrl}/_api/web/lists/getbytitle('PurposeMaster')/items` +
        `?$select=Id,Title,IsActive` +
        `&$filter=IsActive eq 1` +
        `&$orderby=Title asc`;

      const purposeJson = await spGetJson(context, purposeUrl);

      const purposeTitles: string[] = (purposeJson?.value || [])
        .map((x: any) => String(x?.Title || "").trim())
        .filter(Boolean);
      setPurposeOptions(dedupeLabels(purposeTitles));

      const costCenterUrl =
        `${siteUrl}/_api/web/lists/getbytitle('CostCenterMaster')/items` +
        `?$top=5000` +
        `&$select=Title,IsActive` +
        `&$filter=IsActive eq 1` +
        `&$orderby=Title asc`;

      const costCenterJson = await spGetJson(context, costCenterUrl);

      const costCenterTitles: string[] = (costCenterJson?.value || [])
        .map((x: any) => String(x?.Title || "").trim())
        .filter(Boolean);
      setCostCenterOptions(dedupeLabels(costCenterTitles).sort((a, b) => a.localeCompare(b)));
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
          if (isInApprovalStatus(status) || key === statusFilterKey("Submitted")) {
            // Both "Submitted" and "In Approval" mean "pending approval" — after the
            // first approver approves in a multi-stage chain the status becomes
            // "In Approval", so filtering by "Submitted" must include both values.
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
        `Purpose/Title,PurposeDetails,CostCenter/Title,Departments` +
        `&$expand=Requester,CurrentApprover,Purpose,CostCenter` +
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
          requesterEmail: String(it?.Requester?.EMail || "").trim().toLowerCase(),
          approver:
            getActedApproverName(it?.MultipleApproverData, workflowStatus) ??
            String(it?.CurrentApprover?.Title || ""),
          approverNames, // ✅ NEW
          multipleApproverData: it?.MultipleApproverData ? String(it.MultipleApproverData) : undefined,

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

          departmentTitle: String(it?.Departments || ""),

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

        // Never show records created by the login user in Team Requests
        if (row.requesterEmail && row.requesterEmail === loginEmail) return false;

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

  const [costCenterOptions, setCostCenterOptions] = React.useState<string[]>([]);

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
        // FinalDecisionEmail is set when C1 approves in multi-approver flows,
        // not just when the overall request is finalised. Only block if the current
        // user is NOT the current approver — if they are, let the path below decide.
        const _email = currentUserEmailRaw;
        const _isCurrentApprover =
          !!row.currentApproverEmail && row.currentApproverEmail === _email;
        if (!_isCurrentApprover) {
          return { canAct: false, showReject: false };
        }
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

        // Fallback: SP's CurrentApproverEmail is authoritative — covers C2
        // when allApprovalLevels hasn't been updated by PA yet.
        if (!!row.currentApproverEmail && row.currentApproverEmail === loginEmail) {
          return { canAct: true, showReject: true };
        }

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
      allItemResults: SkuMismatch[];
      hasZeroStockItem: boolean;
    }> => {
      const requestId = Number(row.spId || 0);
      if (!requestId) {
        return {
          ok: false,
          invalidItems: [{ itemId: 0, sku: "-", qty: 0, unitPrice: 0, message: "Invalid request id." }],
          skuMismatches: [],
          quantityShort: false,
          allItemResults: [],
          hasZeroStockItem: false,
        };
      }

      const items = await getActiveRequestItems(requestId);
      const invalidItems: ApprovalInvalidItem[] = [];
      const skuMismatches: SkuMismatch[] = [];
      // Every checked item (mismatch or not) — used to persist RaiseCount/ApprovalCount
      // for every SKU on final approval, not just mismatched ones.
      const allItemResults: SkuMismatch[] = [];
      // True when any SKU's live stock came back as zero/unavailable — blocks approval entirely.
      let hasZeroStockItem = false;

      const requestTotalTagPrice = items.reduce((sum, it) => {
        return sum + (Number(it.Quantity || 0) * Number(it.UnitPrice || 0));
      }, 0);

      // Pre-scan: items with qty=0 are always a shortage (no quantity was specified)
      for (const it of items) {
        const reqQty = Number(it.Quantity || 0);
        if (reqQty === 0) {
          const entry: SkuMismatch = {
            itemId: Number(it.Id || 0),
            sku: String(it.SKUCode || "-").trim() || "-",
            tagPricePerUnit: Number(it.UnitPrice || 0),
            requestedQty: 0,
            totalTagPrice: 0,
            availableQty: 0,
            actualTotalTagPrice: 0,
            differenceQty: 0,
            liveStockQty: 0,
          };
          skuMismatches.push(entry);
          allItemResults.push(entry);
        }
      }

      // If total tag price is still 0 after accounting for zero-qty items, skip WMS checks
      if (requestTotalTagPrice === 0) {
        const quantityShort = skuMismatches.length > 0;
        return {
          ok: true,
          invalidItems: [],
          skuMismatches,
          quantityShort,
          allItemResults,
          hasZeroStockItem,
        };
      }

      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        const reqQty = Number(it.Quantity || 0);
        if (reqQty === 0) continue; // already captured above
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
              hasZeroStockItem = true;
              invalidItems.push({
                itemId,
                sku,
                qty: 0,
                unitPrice,
                message: `SKU "${sku}" is out of stock.`,
              });
              {
                const entry: SkuMismatch = {
                  itemId,
                  sku,
                  tagPricePerUnit: unitPrice,
                  requestedQty: reqQty,
                  totalTagPrice,                 // requestedQty × tagPricePerUnit
                  availableQty: 0,
                  actualTotalTagPrice: 0,        // 0 stock → 0 actual total
                  differenceQty: 0 - reqQty, // always negative
                  liveStockQty: 0,
                };
                skuMismatches.push(entry);
                allItemResults.push(entry);
              }

            } else if (reqQty > liveStock) {
              // ── PARTIAL STOCK ───────────────────────────────────────────
              invalidItems.push({
                itemId,
                sku,
                qty: liveStock,
                unitPrice,
                message: `Only ${liveStock} pcs available. Requested: ${reqQty}.`,
              });
              {
                const entry: SkuMismatch = {
                  itemId,
                  sku,
                  tagPricePerUnit: unitPrice,
                  requestedQty: reqQty,
                  totalTagPrice,                       // original total at requested qty
                  availableQty: liveStock,
                  actualTotalTagPrice: liveStock * unitPrice, // actual total at available qty
                  differenceQty: liveStock - reqQty,    // negative = shortfall
                  liveStockQty: liveStock,
                };
                skuMismatches.push(entry);
                allItemResults.push(entry);
              }

            } else {
              // ── QTY OK — sufficient stock, no mismatch, but still record it ──
              // ApprovalCount = the requested qty itself (fully available);
              // StockAvailability = actual live stock.
              allItemResults.push({
                itemId,
                sku,
                tagPricePerUnit: unitPrice,
                requestedQty: reqQty,
                totalTagPrice,
                availableQty: reqQty,
                actualTotalTagPrice: totalTagPrice,
                differenceQty: 0,
                liveStockQty: liveStock,
              });
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
            {
              const entry: SkuMismatch = {
                itemId,
                sku,
                tagPricePerUnit: unitPrice,
                requestedQty: reqQty,
                totalTagPrice,
                availableQty: -1,  // -1 sentinel = unverifiable
                actualTotalTagPrice: -1,
                differenceQty: -1,
                liveStockQty: -1,
              };
              skuMismatches.push(entry);
              allItemResults.push(entry);
            }
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

      const quantityShort = skuMismatches.length > 0;

      return {
        ok: invalidItems.length === 0,
        invalidItems,
        skuMismatches,
        quantityShort,
        allItemResults,
        hasZeroStockItem,
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
      // Every checked item (mismatch or not) — used to persist RaiseCount/ApprovalCount
      // for every SKU on final approval, not just mismatched ones.
      let allItemResults: SkuMismatch[] = [];
      let finalTotalQty = row.qty;
      let finalTotalAmount = row.price;
      // Final approver = the one whose action should persist mismatch data to
      // SharePoint. Single-approver records: any approver is final. Multi-stage
      // records: only the approver at the highest order is final.
      let isFinalApprover = false;

      if (action === "Approve") {
        const check = await validateRequestItemsBeforeApprove(row);
        quantityShort = check.quantityShort;
        skuMismatches = check.skuMismatches;
        allItemResults = check.allItemResults;

        // Hard block: any SKU with zero live stock cannot be approved at all.
        if (check.hasZeroStockItem) {
          throw new Error("Stock availability for the particular record is zero, so you can't approve this record.");
        }

        // ── FINAL APPROVER DETECTION ────────────────────────────────────
        // The final approver is the one at the highest order in the approval
        // chain. Only the final approver actually writes adjusted quantities
        // to SharePoint — intermediate approvers just pass mismatch data along.
        const loginEmail = String(currentUserEmailRaw || "").trim().toLowerCase();
        const activeLevels = row.allApprovalLevels.filter(
          (l) => l.canApproveRequest === true
        );
        if (!row.hasMultipleApprovals) {
          // Single-approver record — always the final approver
          isFinalApprover = true;
        } else if (activeLevels.length > 0) {
          // Multi-approver: final = the approver at the highest order
          const maxOrder = Math.max(...activeLevels.map((l) => l.order));
          isFinalApprover = activeLevels.some(
            (l) => l.approverEmail === loginEmail && l.order === maxOrder
          );
        }

        // The final approver adjusts short-stock quantities. IMPORTANT: nothing is
        // written to SharePoint before the Power Automate call — if PA errors, no list
        // change occurs. Here we only COMPUTE the adjusted totals (in memory) so the PA
        // payload carries the final figures; the actual writes happen AFTER PA succeeds.
        // Totals are SUMMED from the approved quantity of every line item — never
        // derived by subtracting from a cached header total, which could be stale and
        // drive the result negative. Summing non-negative values can never go below 0.
        if (isFinalApprover && quantityShort && skuMismatches.length > 0 && allItemResults.length > 0) {
          let qtySum = 0;
          let priceSum = 0;
          for (const m of allItemResults) {
            const requested = Number.isFinite(Number(m.requestedQty)) ? Number(m.requestedQty) : 0;
            const availRaw = Number(m.availableQty);
            // Unverifiable stock (-1) keeps the requested quantity.
            const adjQty =
              !Number.isFinite(availRaw) || availRaw === -1 ? requested : Math.max(0, availRaw);
            const unitPrice = Number.isFinite(Number(m.tagPricePerUnit)) ? Number(m.tagPricePerUnit) : 0;
            qtySum += adjQty;
            priceSum += adjQty * unitPrice;
          }
          finalTotalQty = Math.max(0, qtySum);
          finalTotalAmount = Math.max(0, priceSum);
        }
      }

      // ── DETAIL VIEW URL ────────────────────────────────────────────────
      const siteOrigin = String(context.pageContext.web.absoluteUrl || "").replace(/\/$/, "");
      const detailViewUrl = `${siteOrigin}/SitePages/RequestDetail.aspx?requestId=${row.spId}`;

      // Strip the internal-only liveStockQty field — the Power Automate trigger schema
      // rejects unknown properties (additionalProperties: false).
      const payloadSkuMismatches = skuMismatches.map((m) => ({
        itemId: m.itemId,
        sku: m.sku,
        tagPricePerUnit: m.tagPricePerUnit,
        requestedQty: m.requestedQty,
        totalTagPrice: m.totalTagPrice,
        availableQty: m.availableQty,
        actualTotalTagPrice: m.actualTotalTagPrice,
        differenceQty: m.differenceQty,
      }));

      // ── PAYLOAD ────────────────────────────────────────────────────────
      // totalQuantity / totalTagPrice carry the post-adjustment values so
      // Power Automate emails show the final approved figures, not the
      // originally requested ones.
      const payload: ApprovalFlowPayload = {
        requestId: row.spId,
        actorEmail: currentUserEmailRaw,
        position: profile?.jobTitle || "",
        action: action as FlowAction,
        comments: comments || "",
        quantityShort,
        skuMismatches: payloadSkuMismatches,
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

      // PA succeeded (spPostJson throws on HTTP error) — persist the adjusted header
      // totals. NOTE: the line-item Quantity/LineTotal are intentionally NOT overwritten
      // — the originally requested quantity stays intact and the approved figure is
      // recorded in ApprovalCount instead.
      if (isFinalApprover && action === "Approve" && quantityShort && skuMismatches.length > 0) {
        await context.spHttpClient.post(
          `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${Number(row.spId || 0)})`,
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
              // Hard guard: never store a negative total in SharePoint.
              TotalQuantity: Math.max(0, Number(finalTotalQty) || 0),
              TotalAmount: Math.max(0, Number(finalTotalAmount) || 0),
            }),
          }
        );
      }

      // Persist RaiseCount / ApprovalCount / StockAvailability for every checked item —
      // mismatched or not (best-effort, non-fatal). Same final-approver rule as the
      // Quantity/Total update above — intermediate approvers in a multi-stage chain
      // must not write.
      console.log("[EForm Approval] column-write gate:", {
        isFinalApprover,
        action,
        allItemResults: allItemResults.length,
        willWrite: isFinalApprover && action === "Approve" && allItemResults.length > 0,
      });
      if (isFinalApprover && action === "Approve" && allItemResults.length > 0) {
        try {
          const CHUNK = 8;
          for (let i = 0; i < allItemResults.length; i += CHUNK) {
            const chunk = allItemResults.slice(i, i + CHUNK);
            await Promise.all(
              chunk.map((m) => {
                // Coerce to finite numbers — an undefined/NaN value would be dropped or
                // sent as null by JSON.stringify, leaving the column blank.
                const requested = Number.isFinite(Number(m.requestedQty)) ? Number(m.requestedQty) : 0;
                const availRaw = Number(m.availableQty);
                // Items whose live stock could not be verified (-1) still record
                // RaiseCount/ApprovalCount from the requested quantity; only
                // StockAvailability is omitted, since the live figure is unknown.
                const approved =
                  !Number.isFinite(availRaw) || availRaw === -1 ? requested : availRaw;
                const unitPrice = Number.isFinite(Number(m.tagPricePerUnit))
                  ? Number(m.tagPricePerUnit)
                  : 0;
                const fields: Record<string, number> = {
                  RaiseCount: requested,
                  ApprovalCount: approved,
                  // LineTotal reflects the APPROVED quantity (Quantity itself is left as
                  // originally requested), so it matches the header TotalAmount.
                  LineTotal: Math.max(0, approved * unitPrice),
                };
                const liveRaw = Number(m.liveStockQty);
                if (Number.isFinite(liveRaw) && liveRaw >= 0) fields.StockAvailability = liveRaw;
                return context.spHttpClient.post(
                  `${siteUrl}/_api/web/lists/getbytitle('InventoryRequestItems')/items(${m.itemId})`,
                  SPHttpClient.configurations.v1,
                  {
                    headers: {
                      Accept: "application/json;odata=nometadata",
                      "Content-Type": "application/json;odata=nometadata",
                      "odata-version": "",
                      "IF-MATCH": "*",
                      "X-HTTP-Method": "MERGE",
                    },
                    body: JSON.stringify(fields),
                  }
                );
              })
            );
          }
          console.log(
            `[EForm Approval] Wrote RaiseCount/ApprovalCount/StockAvailability for ${allItemResults.length} item(s).`
          );
        } catch (err) {
          console.error("[TeamService] RaiseCount/ApprovalCount update failed:", err);
        }
      }

      // ── INTERMEDIATE MULTI-STAGE APPROVAL ─────────────────────────────
      // When A1 approves but there are still more stages (e.g. C1 approves,
      // C2 still needs to), Power Automate does not reliably set
      // WorkflowStatus → "In Approval" or flip CurrentApproverEmail to the
      // next approver. The frontend does it directly so A2 sees the record
      // immediately with the correct status and can act on it.
      if (ok && action === "Approve" && row.hasMultipleApprovals) {
        const loginEmail = String(currentUserEmailRaw || "").trim().toLowerCase();

        const approveableLevels = row.allApprovalLevels
          .filter((l) => l.canApproveRequest)
          .sort((a, b) => a.order - b.order);

        const currentLevelIdx = approveableLevels.findIndex(
          (l) => l.approverEmail === loginEmail
        );

        const isIntermediateApproval =
          currentLevelIdx >= 0 &&
          currentLevelIdx < approveableLevels.length - 1;

        if (isIntermediateApproval) {
          const nextLevel = approveableLevels[currentLevelIdx + 1];
          const nextStage = row.workflowStage + 1;

          // Resolve next approver's SharePoint user ID (non-blocking on error)
          let nextApproverId: number | undefined;
          try {
            const ensureUrl =
              `${siteUrl}/_api/web/ensureUser('${nextLevel.approverEmail.replace(/'/g, "''")}')`;
            const ensureRes = await context.spHttpClient.post(
              ensureUrl,
              SPHttpClient.configurations.v1,
              {
                headers: { "Content-Type": "application/json;odata=verbose" },
                body: JSON.stringify({}),
              }
            );
            if (ensureRes.ok) {
              const ensureJson = (await ensureRes.json()) as { Id?: number };
              nextApproverId = ensureJson.Id;
            }
          } catch { /* non-blocking */ }

          // Write the intermediate state to SharePoint
          await context.spHttpClient.post(
            `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${row.spId})`,
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
                WorkflowStatus: "In Approval",
                WorkflowStage: nextStage,
                CurrentApproverEmail: nextLevel.approverEmail,
                ...(nextApproverId ? { CurrentApproverId: nextApproverId } : {}),
                LastUpdatedOn: new Date().toISOString(),
              }),
            }
          );

          // Update local row so the table reflects the change immediately
          setRows((prev) =>
            prev.map((r) =>
              r.spId === row.spId
                ? {
                    ...r,
                    status: "In Approval",
                    workflowStage: nextStage,
                    currentApproverEmail: nextLevel.approverEmail,
                    currentApproverId: nextApproverId || 0,
                    qty: finalTotalQty,
                    price: finalTotalAmount,
                    lastActivityIso: new Date().toISOString(),
                  }
                : r
            )
          );

          return { success: true, message: msg, newStatus: "In Approval" };
        }
      }

      // ── UPDATE LOCAL ROW STATE on success (final approval / rejection) ──
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