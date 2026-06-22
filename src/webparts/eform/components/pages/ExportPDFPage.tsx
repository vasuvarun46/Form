import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import { SPHttpClient } from "@microsoft/sp-http";
import LogoPng from "../../assets/LOGO.png";
import { Route } from '../Eform';
import { ExportActionInfo } from './ExportActionInfo';
import ExportPDFHeader from '../Common/PageHeader/ExportPDFHeader';
import { RouteState, SecurityPermissions } from '../spHelpers';
import { useIsDesktop } from "../hooks/useIsDesktop";
import NotificationBar from "../../components/pages/TeamRequest/NotificationBar";

type PendingAction = "Approve" | "Reject" | "Bypass";

type InvalidItem = {
  itemId: number;
  sku: string;
  qty: number;
  unitPrice: number;
  message: string;
};

type ExportActionHandlers = {
  submitDecision: (
    action: PendingAction,
    comments: string
  ) => Promise<{ success: boolean; message: string; newStatus?: string }>;

  validateBeforeApprove: () => Promise<{
    ok: boolean;
    invalidItems: InvalidItem[];
  }>;

  updateInvalidQuantity: (
    itemId: number,
    nextQty: number
  ) => Promise<void>;

  reloadAfterAction?: () => Promise<void>;
};

type Props = {
  context: WebPartContext;
  spId: number | undefined;
  exportActionInfo?: ExportActionInfo;
  onNavigate: (route: Route, state?: RouteState) => void;
  onApprove?: (comments?: string) => Promise<string>;  // ✅ returns message
  onReject?: (comments?: string) => Promise<string>;   // ✅ returns message
  onBypass?: (comments?: string) => Promise<string>;   // ✅ returns message
  exportSource?: Route;
  permissions: SecurityPermissions;
  autoExport?: boolean;
  onAutoExportDone?: () => void;
  actionHandlers?: ExportActionHandlers;
};

type FieldValuesAsText = Record<string, string>;
type SpLookupTitle = { Title?: string };
type SpPersonTitle = { Title?: string; EMail?: string };

type InventoryRequestItemJson = Record<string, unknown> & {
  Id?: number;
  Requester?: SpPersonTitle;
  Department?: SpLookupTitle;
  CostCenter?: SpLookupTitle;
  Purpose?: SpLookupTitle;
  RequestSubmittedOn?: string;
  CurrentApproverEmail?: string;
};

type ExportData = {
  Title: string;
  requester: string;
  jobPosition: string;
  department: string;
  costCenter: string;
  expectedDeliveryDate: string;
  placeOfDelivery: string;
  purpose: string;
  purposeDetails: string;
  submittedDate: string;
  approverName: string;
  approverNames?: string;
  approverPosition: string;
  approvedDate: string;
  rejectReason: string;
  approvalStatus: string;
  workflowStatus: string;
  currentApproverEmail?: string;
};

type LineItem = {
  id: number;
  title: string;
  lineNumber: number;
  skuCode: string;
  productDescription: string;
  color: string;
  size: string;
  unitPrice: number;
  quantity: number;
  lineTotal: number;
  lookupSource: string;
  rowStatus: string;
};

// ✅ UPDATED: approvalRole and approverEmail are required fields
type ParsedMultipleApproverData = {
  approvalType?: string;
  stageCount: number;
  approvalUsers: Array<{
    name: string;
    position: string;
    order: number;
  }>;
  stages: Array<{
    stage: number;
    approverName: string;
    approverJobPosition: string;
    approvalRole: string;   // ✅ required — e.g. "C1 Level", "C2 Level"
    approverEmail: string;  // ✅ required — used as name fallback
    action: string;
    date: string;
    reason: string;
    actionBy: string;
    Status?: string;
    Date?: string;
  }>;
};

// ✅ UPDATED: approverEmail added, status allows undefined
type FooterApproverData = {
  slot: number;
  name: string;
  position: string;
  status: "Approved" | "Rejected" | undefined;
  date: string;
  reason: string;
  bypassBy: string;
  approverEmail?: string;  // ← optional, fixes all missing-property errors
  approvalRole?: string;
  comments?: string;
};

// ✅ ADD this new type at the top with other types
type AllApprovalLevel = {
  order: number;
  approvalRole: string;
  approverEmail: string;
  approverId: number;
  canApproveRequest: boolean;
  canViewRequest: boolean;
  isCurrentApprover: boolean;
  isEnabledByLOA: boolean;
  emailMissing: boolean;
};

function getStatusDisplay(status?: string): { label: string; className: string } {
  const s = (status || "").toLowerCase().trim();
  if (s === "submitted") return { label: "ส่งแล้ว / SUBMITTED", className: "text-blue-600" };
  if (s === "approved") return { label: "อนุมัติ / APPROVED", className: "text-green-600" };
  if (s === "rejected") return { label: "ไม่อนุมัติ / REJECTED", className: "text-red-600" };
  if (s === "in approval" || s === "pending") return { label: "อยู่ระหว่างการอนุมัติ / IN APPROVAL", className: "text-amber-600" };
  if (s === "draft") return { label: "แบบร่าง / DRAFT", className: "text-gray-500" };
  return { label: status || "-", className: "text-gray-600" };
}

function getSubmitDisplay(status?: string): { label: string; className: string } {
  const s = (status || "").toLowerCase().trim();
  if (s === "submitted") return { label: "Submitted", className: "text-blue-600" };
  if (s === "approved") return { label: "Approved", className: "text-green-600" };
  if (s === "rejected") return { label: "Rejected", className: "text-red-600" };
  if (s === "in approval" || s === "pending") return { label: "In Approval", className: "text-amber-600" };
  if (s === "draft") return { label: "Draft", className: "text-gray-500" };
  return { label: status || "-", className: "text-gray-600" };
}

function normalizeApprovalStatus(raw?: string): "Approved" | "Rejected" | undefined {
  const s = (raw || "").toLowerCase();
  if (s === "approved" || s === "approve" || s === "bypassed" || s === "bypass") return "Approved";
  if (s === "rejected" || s === "reject") return "Rejected";
  return undefined;
}

function cleanDisplayText(value?: string): string {
  const s = (value || "").trim();
  if (!s) return "";
  const lower = s.toLowerCase();
  if (lower === "undefined" || lower === "null" || lower === "nan") return "";
  return s;
}

function cleanPositionText(value?: string): string {
  const s = cleanDisplayText(value).replace(/\s+/g, " ").trim();
  if (!s) return "";
  const exactRepeat = /^(.{4,}?)\1+$/.exec(s);
  if (exactRepeat && exactRepeat[1]) return exactRepeat[1].trim();
  const spacedRepeat = /^(.{4,}?)(?:\s+\1)+$/.exec(s);
  if (spacedRepeat && spacedRepeat[1]) return spacedRepeat[1].trim();
  return s;
}

/* =========================================================
   fetchSPUserProfile
   Queries the SharePoint People Search index for a user by
   WorkEmail and returns their JobTitle and Department.
   No Graph API or Entra ID required.
========================================================= */
type PeopleSearchResponse = {
  PrimaryQueryResult?: {
    RelevantResults?: {
      Table?: {
        Rows?: Array<{ Cells?: Array<{ Key: string; Value: string }> }>;
      };
    };
  };
};

async function fetchSPUserProfile(
  context: WebPartContext,
  email: string
): Promise<{ jobTitle: string; department: string }> {
  const empty = { jobTitle: "", department: "" };
  try {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const safeEmail = encodeURIComponent(email.toLowerCase().trim());

    const url =
      `${siteUrl}/_api/search/query` +
      `?querytext='WorkEmail:${safeEmail}'` +
      `&sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'` +
      `&selectproperties='JobTitle,Department'` +
      `&rowlimit=1`;

    const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: "application/json;odata=nometadata", "odata-version": "" },
    });
    if (!res.ok) return empty;

    const data = (await res.json()) as PeopleSearchResponse;
    const cells = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows?.[0]?.Cells ?? [];
    const get = (key: string): string =>
      String(cells.find((c) => c.Key === key)?.Value ?? "").trim();

    return {
      jobTitle: cleanDisplayText(get("JobTitle")),
      department: cleanDisplayText(get("Department")),
    };
  } catch {
    return empty;
  }
}

async function fetchEFormUserJobPosition(
  context: WebPartContext,
  email: string
): Promise<string> {
  try {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const safeEmail = email.replace(/'/g, "''");
    const url =
      `${siteUrl}/_api/web/lists/getbytitle('EForm%20Users')/items` +
      `?$select=Jobposition` +
      `&$filter=Email eq '${safeEmail}'` +
      `&$top=1`;
    const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) return "";
    const json = await res.json();
    return cleanDisplayText(String(json?.value?.[0]?.Jobposition || ""));
  } catch {
    return "";
  }
}

function parseMultipleApproverData(raw: unknown): ParsedMultipleApproverData | null {
  if (raw === null || raw === undefined) return null;

  let parsed: unknown = raw;
  if (typeof raw === "string") {
    const s = raw.trim();
    if (!s) return null;
    try {
      parsed = JSON.parse(s);
    } catch {
      return null;
    }
  }

  if (!parsed || typeof parsed !== "object") return null;
  const p = parsed as Record<string, unknown>;

  const usersRaw = Array.isArray(p.approvalUsers) ? p.approvalUsers : [];
  const approvalUsers = usersRaw
    .map((u) => {
      const user = (u || {}) as Record<string, unknown>;
      const orderValue = Number(user.Order ?? user.order ?? user.stage ?? 0);
      return {
        name: cleanDisplayText(
          String(user.UserName ?? user.userName ?? user.approverName ?? user.name ?? "")
        ),
        position: cleanPositionText(
          String(user.Jobposition ?? user.JobPosition ?? user.position ?? user.approverJobPosition ?? "")
        ),
        order: Number.isFinite(orderValue) ? orderValue : 0,
      };
    })
    .filter((u) => !!u.name)
    .sort((a, b) => {
      if (a.order > 0 && b.order > 0) return a.order - b.order;
      if (a.order > 0) return -1;
      if (b.order > 0) return 1;
      return a.name.localeCompare(b.name);
    });

  // ✅ Parse allApprovalLevels to resolve approver names from email
  const allApprovalLevelsRaw = Array.isArray(p.allApprovalLevels)
    ? p.allApprovalLevels
    : [];
  const allApprovalLevels: AllApprovalLevel[] = allApprovalLevelsRaw.map((l) => {
    const level = (l || {}) as Record<string, unknown>;
    return {
      order: Number(level.order ?? 0),
      approvalRole: cleanDisplayText(String(level.approvalRole ?? "")),
      approverEmail: cleanDisplayText(String(level.approverEmail ?? "")).toLowerCase(),
      approverId: Number(level.approverId ?? 0),
      canApproveRequest: level.canApproveRequest === true,
      canViewRequest: level.canViewRequest === true,
      isCurrentApprover: level.isCurrentApprover === true,
      isEnabledByLOA: level.isEnabledByLOA === true,
      emailMissing: level.emailMissing === true,
    };
  });

  const stagesRaw = Array.isArray(p.stages) ? p.stages : [];
  const stages = stagesRaw
    .map((s) => {
      const stage = (s || {}) as Record<string, unknown>;
      const stageValue = Number(stage.stage ?? stage.Stage ?? stage.order ?? stage.Order ?? 0);

      const actionByEmail = cleanDisplayText(
        String(stage.ActionBy ?? stage.actionBy ?? stage.actionByEmail ?? "")
      ).toLowerCase();

      const approverEmailRaw = cleanDisplayText(
        String(stage.approverEmail ?? stage.ApproverEmail ?? "")
      ).toLowerCase();

      const emailToResolve = approverEmailRaw || actionByEmail;
      const matchedLevel = emailToResolve
        ? allApprovalLevels.find((l) => l.approverEmail === emailToResolve)
        : undefined;

      const resolvedApproverName = cleanDisplayText(
        String(stage.approverName ?? stage.ApproverName ?? stage.UserName ?? "")
      );

      const resolvedPosition = cleanPositionText(
        String(stage.approverJobPosition ?? stage.Jobposition ?? stage.JobPosition ?? "")
      );

      return {
        stage: Number.isFinite(stageValue) ? stageValue : 0,
        approverName: resolvedApproverName,
        approverJobPosition: resolvedPosition,
        approvalRole: cleanDisplayText(
          String(stage.approvalRole ?? stage.ApprovalRole ?? matchedLevel?.approvalRole ?? "")
        ),
        action: cleanDisplayText(
          String(stage.Status ?? stage.status ?? stage.Action ?? stage.action ?? stage.ApprovalStatus ?? stage.approvalStatus ?? "")
        ),
        date: cleanDisplayText(
          String(stage.Date ?? stage.date ?? stage.ActionDate ?? stage.actionDate ?? "")
        ),
        // ✅ FIXED: added stage.reason and stage.Reason at the TOP of the chain
        reason: cleanDisplayText(
          String(
            stage.reason ??             // ✅ lowercase "reason" — your PA flow writes this
            stage.Reason ??             // ✅ uppercase variant
            stage.RejectReason ??
            stage.rejectReason ??
            stage.RejectionReason ??
            stage.rejectionReason ??
            stage.Comments ??
            stage.comments ??
            ""
          )
        ),
        actionBy: cleanDisplayText(
          String(stage.ActionBy ?? stage.actionBy ?? stage.actionByEmail ?? "")
        ),
        approverEmail: approverEmailRaw,
        Status: cleanDisplayText(String(stage.Status ?? stage.status ?? "")),
        Date: cleanDisplayText(String(stage.Date ?? stage.date ?? "")),
      };
    })
    .filter((s) => s.stage > 0 || !!s.approverName || !!s.approverEmail)
    .sort((a, b) => {
      if (a.stage > 0 && b.stage > 0) return a.stage - b.stage;
      if (a.stage > 0) return -1;
      if (b.stage > 0) return 1;
      return a.approverName.localeCompare(b.approverName);
    });

  const stageCountCandidate = Number(p.stageCount ?? p.StageCount ?? 0);
  const stageCount = Number.isFinite(stageCountCandidate) && stageCountCandidate > 0
    ? stageCountCandidate
    : (approvalUsers.length || stages.length || 0);

  return {
    approvalType: String(p.approvalType || ""),
    stageCount,
    approvalUsers,
    stages,
  };
}

function extractRejectReasonText(text?: Record<string, string>): string {
  if (!text) return "";
  const explicitKeys = [
    "RejectReason", "RejectedReason", "RejectionReason", "RejectComments",
    "RejectComment", "FinalDecisionReason", "FinalDecisionComment",
    "ApprovalComments", "ApproverComments", "Comments", "Comment",
    "Reason", "Remarks", "Remark",
  ];
  for (const key of explicitKeys) {
    const value = cleanDisplayText(text[key]);
    if (value) return value;
  }
  const fallback = Object.entries(text).find(([key, value]) => {
    const k = key.toLowerCase();
    if (!cleanDisplayText(value)) return false;
    if (!/(reject|reason|comment|remark)/i.test(k)) return false;
    if (/(editor|author|created|modified|version|workflow|status|approver|request)/i.test(k)) return false;
    return true;
  });
  return fallback ? cleanDisplayText(fallback[1]) : "";
}

function extractDateFromMultipleApproverData(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  let parsed: unknown = raw;
  if (typeof raw === "string") {
    const s = raw.trim();
    if (!s) return "";
    try { parsed = JSON.parse(s); } catch { return ""; }
  }
  const p = parsed as { stages?: Array<{ Date?: string; date?: string }> };
  const stages = Array.isArray(p?.stages) ? p.stages : [];
  if (!stages.length) return "";

  let pickedRaw = "";
  let pickedTs = Number.NEGATIVE_INFINITY;
  for (const st of stages) {
    const candidateRaw = cleanDisplayText(st?.Date || st?.date || "");
    if (!candidateRaw) continue;
    const ts = new Date(candidateRaw).getTime();
    if (Number.isFinite(ts) && ts > pickedTs) {
      pickedTs = ts;
      pickedRaw = candidateRaw;
    } else if (!pickedRaw) {
      pickedRaw = candidateRaw;
    }
  }
  return pickedRaw;
}

function formatDateDdMmYyyy(value?: string): string {
  const raw = cleanDisplayText(value);
  if (!raw) return "";
  const d = new Date(raw);
  if (!isNaN(d.getTime())) {
    const bangkokOffset = 7 * 60 * 60 * 1000;
    const bangkokDate = new Date(d.getTime() + bangkokOffset);
    const dd = String(bangkokDate.getUTCDate()).padStart(2, "0");
    const mm = String(bangkokDate.getUTCMonth() + 1).padStart(2, "0");
    const yyyy = String(bangkokDate.getUTCFullYear());
    return `${dd}-${mm}-${yyyy}`;
  }
  return raw;
}

export default function ExportPDFPage({
  context,
  spId,
  exportActionInfo,
  onNavigate,
  onApprove,
  onReject,
  onBypass,
  exportSource,
  permissions,
  autoExport = false,
  onAutoExportDone,
  actionHandlers,
}: Props): JSX.Element {
  const listTitle = "InventoryRequests";
  const [exportData, setExportData] = React.useState<ExportData | null>(null);
  const [error, setError] = React.useState<string>("");
  const [lineItems, setLineItems] = React.useState<LineItem[]>([]);
  const [lineError, setLineError] = React.useState<string>("");
  const [lineLoading, setLineLoading] = React.useState<boolean>(false);
  const isDesktopTable = useIsDesktop(700);
  const pdfRef = React.useRef<HTMLDivElement | null>(null);
  const pdfPagesRef = React.useRef<HTMLDivElement | null>(null);
  const autoExportTriggeredRef = React.useRef<boolean>(false);
  const [multipleApproverRaw, setMultipleApproverRaw] = React.useState<unknown>(null);
  const [resolvedApproverNames, setResolvedApproverNames] = React.useState<Record<string, string>>({});
  const [requestorAadProfile, setRequestorAadProfile] = React.useState<{ jobTitle: string; department: string } | null>(null);
  const [approverAadPositions, setApproverAadPositions] = React.useState<Record<string, string>>({});
  const fetchedApproverEmails = React.useRef<Set<string>>(new Set());
  const [pdfBusy, setPdfBusy] = React.useState(false);
  const canAct = exportActionInfo?.canAct ?? false;
  const showReject = exportActionInfo?.showReject ?? false;
  const isBypassMode = exportActionInfo?.mode === "bypass";
  const isApproveMode = exportActionInfo?.mode === "approve" || (canAct && !isBypassMode);
  const isApproverWaiting = exportActionInfo?.isApproverWaiting ?? false;
  const canExportPdf =
    exportSource === "my" ? permissions?.exportOwnReport
      : exportSource === "team" ? permissions?.exportTeamReport
        : exportSource === "all" ? permissions?.exportAllReport
          : false;

  const [toast, setToast] = React.useState<null | { type: "success" | "error"; message: string }>(null);
  React.useEffect(() => {
    if (!toast) return;
  }, [toast]);

  const [actionModalOpen, setActionModalOpen] = React.useState(false);
  const [pendingAction, setPendingAction] = React.useState<PendingAction | null>(null);
  const [actionComments, setActionComments] = React.useState("");
  const [actionSubmitting, setActionSubmitting] = React.useState(false);
  const [editQtyPopupOpen, setEditQtyPopupOpen] = React.useState(false);
  const [editRows, setEditRows] = React.useState<InvalidItem[]>([]);

  React.useEffect(() => {
    if (!toast) return;
    setActionModalOpen(false);
    setPendingAction(null);
    setActionComments("");
    setEditQtyPopupOpen(false);
    setEditRows([]);
  }, [toast]);

  React.useEffect(() => {
    const safeText = (v: unknown): string =>
      cleanDisplayText(v === null || v === undefined ? "" : String(v));

    const load = async (): Promise<void> => {
      try {
        setError("");
        setExportData(null);
        setMultipleApproverRaw(null);
        if (!spId) return;

        const webUrl = context.pageContext.web.absoluteUrl;
        const requesterInternal = "Requester";
        const jobInternal = "JobPosition";
        const deptInternal = "Department";
        const costInternal = "CostCenter";
        const expectedInternal = "ExpectedDeliveryDate";
        const placeInternal = "PlaceOfDelivery";
        const purposeInternal = "Purpose";
        const purposeDetailsInternal = "PurposeDetails";
        const multipleApproverDataInternal = "MultipleApproverData";
        const workflowStatusInternal = "WorkflowStatus";
        const Title = "Title";
        const ApproverNames = "ApproverNames"

        const url =
          `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${spId})` +
          `?$select=Id,RequestSubmittedOn,${Title},${ApproverNames},CurrentApproverEmail,` +
          `${requesterInternal}/Title,${requesterInternal}/EMail,${jobInternal},${deptInternal}/Title,` +
          `${expectedInternal},${placeInternal},${purposeDetailsInternal},` +
          `${multipleApproverDataInternal},${costInternal}/Title,` +
          `${purposeInternal}/Title,${workflowStatusInternal}` +
          `&$expand=${requesterInternal},${deptInternal},${costInternal},${purposeInternal}`;

        const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!res.ok) {
          const body = await res.text();
          throw new Error(`SharePoint ${res.status}: ${body}`);
        }

        const d = (await res.json()) as InventoryRequestItemJson;
        setMultipleApproverRaw(d[multipleApproverDataInternal]);

        const textUrl =
          `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${spId})/FieldValuesAsText`;
        const textRes = await context.spHttpClient.get(textUrl, SPHttpClient.configurations.v1);
        if (!textRes.ok) {
          const body = await textRes.text();
          throw new Error(`FieldValuesAsText ${textRes.status}: ${body}`);
        }

        const t = (await textRes.json()) as FieldValuesAsText;
        let submittedDate = "";
        const rawSubmitted = d.RequestSubmittedOn;
        if (typeof rawSubmitted === "string" && rawSubmitted) {
          submittedDate = formatDateDdMmYyyy(rawSubmitted);
        }

        const approvalStatus = safeText(t.WorkflowStatus) || safeText(t.Status);
        const rejectedByStatusText = /\breject(ed)?\b/i.test(approvalStatus);
        const approverName =
          safeText(t.CurrentApprover) || safeText(t.Approver) || safeText(t.FinalDecisionBy);
        const approverPosition = cleanDisplayText(safeText(t.ApproverPosition));
        const rawApprovedDate =
          safeText(t.FinalDecisionDate) || safeText(t.ApprovedDate) || safeText(t.Modified);
        const approvedDateFromFields = cleanDisplayText(rawApprovedDate);
        const approvedDateFromMultipleApproverData =
          extractDateFromMultipleApproverData(d[multipleApproverDataInternal]);
        const approvedDate = rejectedByStatusText
          ? (approvedDateFromMultipleApproverData || approvedDateFromFields)
          : (approvedDateFromFields || approvedDateFromMultipleApproverData);
        const rejectReason = extractRejectReasonText(t);
        const workflowStatus =
          safeText(t.WorkflowStatus) || safeText(d[workflowStatusInternal]);
        const expectedRaw = d[expectedInternal];
        const expectedDate =
          typeof expectedRaw === "string" && expectedRaw
            ? formatDateDdMmYyyy(expectedRaw) : "";

        const requesterEmail = safeText(
          (d[requesterInternal] as SpPersonTitle | undefined)?.EMail || ""
        );

        setExportData({
          requester: safeText((d[requesterInternal] as SpPersonTitle | undefined)?.Title),
          jobPosition: safeText(d[jobInternal]),
          department: safeText((d[deptInternal] as SpLookupTitle | undefined)?.Title),
          costCenter: safeText((d[costInternal] as SpLookupTitle | undefined)?.Title),
          expectedDeliveryDate: expectedDate,
          placeOfDelivery: safeText(d[placeInternal]),
          purpose: safeText((d[purposeInternal] as SpLookupTitle | undefined)?.Title),
          purposeDetails: safeText(d[purposeDetailsInternal]),
          Title: safeText(d[Title]),
          approverNames: safeText(d[ApproverNames]),
          submittedDate,
          approverName,
          approverPosition,
          approvedDate,
          rejectReason,
          approvalStatus,
          workflowStatus,
          currentApproverEmail: safeText(d.CurrentApproverEmail || "").toLowerCase(),
        });

        // Fetch requestor's job title and department from Azure AD (via SP User Profile service)
        if (requesterEmail) {
          try {
            const aadProfile = await fetchSPUserProfile(context, requesterEmail);
            setRequestorAadProfile(aadProfile);
          } catch {
            // Silently fall back to SharePoint field data
          }
        }
      } catch (e: unknown) {
        setError(e instanceof Error ? e.message : String(e));
        setMultipleApproverRaw(null);
      }
    };

    load().catch(() => undefined);
  }, [context, spId]);




  const parsedMultipleApproverData = React.useMemo(
    () => parseMultipleApproverData(multipleApproverRaw),
    [multipleApproverRaw]
  );



  const resolvedRequestorName = cleanDisplayText(exportData?.requester);
  // Prefer Azure AD job title; fall back to SharePoint JobPosition field
  const resolvedRequestorPosition =
    cleanPositionText(requestorAadProfile?.jobTitle) ||
    cleanPositionText(exportData?.jobPosition) ||
    "-";
  // Prefer Azure AD department; fall back to SharePoint Department lookup
  const resolvedRequestorDepartment =
    cleanDisplayText(exportData?.department) ||
    cleanDisplayText(requestorAadProfile?.department) ||
    "";

  // =========================================================
  // ✅ SIMPLIFIED footerApprovers — no SingleApprover/MultipleApprover branching.
  //    Logic: iterate stages from JSON. If a stage has a date → it acted → show data.
  //    If no date → show empty placeholder.
  //    If no stages at all → fall back to exportData fields (single approver legacy).
  // =========================================================
  const footerApprovers = React.useMemo<FooterApproverData[]>(() => {
    const workflowType = parsedMultipleApproverData?.approvalType || "SingleApprover";
    const stages = parsedMultipleApproverData?.stages ?? [];

    // -----------------------------
    // SINGLE APPROVER WORKFLOW
    // -----------------------------
    if (workflowType === "SingleApprover") {
      const actedStage = stages.find((s) => !!cleanDisplayText(s.date || s.Date || ""));

      if (actedStage) {
        return [{
          slot: 1,
          name: cleanDisplayText(exportData?.approverNames || exportData?.approverName),
          position: cleanDisplayText(actedStage.approverJobPosition || ""),
          status: normalizeApprovalStatus(actedStage.action || actedStage.Status),
          date: formatDateDdMmYyyy(actedStage.date || actedStage.Date),
          reason: cleanDisplayText(actedStage.reason),
          bypassBy: "",
          // Use stage email if present; fall back to CurrentApproverEmail from SP
          approverEmail: actedStage.approverEmail || exportData?.currentApproverEmail || "",
          approvalRole: cleanDisplayText(actedStage.approvalRole),
        }];
      }

      // No one acted yet — use CurrentApproverEmail so Search API can fetch their position
      return [{
        slot: 1,
        name: cleanDisplayText(exportData?.approverNames || exportData?.approverName),
        position: cleanPositionText(
          exportData?.approverPosition ||
          parsedMultipleApproverData?.stages?.[0]?.approverJobPosition ||
          parsedMultipleApproverData?.approvalUsers?.[0]?.position
        ),
        status: normalizeApprovalStatus(
          exportData?.workflowStatus || exportData?.approvalStatus
        ),
        date: formatDateDdMmYyyy(cleanDisplayText(exportData?.approvedDate)),
        reason: cleanDisplayText(exportData?.rejectReason),
        bypassBy: "",
        approverEmail: exportData?.currentApproverEmail || "",
      }];
    }

    // -----------------------------
    // MULTIPLE APPROVER WORKFLOW
    // -----------------------------
    if (!stages.length) {
      return [{
        slot: 1,
        name: "",
        position: "",
        status: undefined,
        date: "",
        reason: "",
        bypassBy: "",
        approverEmail: "",  // ✅
      }];
    }

    return stages.map((s, index) => {
      const slotNumber = s.stage > 0 ? s.stage : index + 1;
      const hasActed = !!cleanDisplayText(s.date || s.Date || "");

      if (!hasActed) {
        return {
          slot: slotNumber,
          name: "",
          position: "",
          status: undefined as "Approved" | "Rejected" | undefined,
          date: "",
          reason: "",
          bypassBy: "",
          approverEmail: s.approverEmail,  // ✅
          approvalRole: cleanDisplayText(s.approvalRole),
        };
      }

      return {
        slot: slotNumber,
        name: cleanDisplayText(s.approverName),
        position: cleanDisplayText(s.approverJobPosition || ""),
        status: normalizeApprovalStatus(s.action || s.Status),
        date: formatDateDdMmYyyy(s.date || s.Date),
        reason: cleanDisplayText(s.reason),
        bypassBy: "",
        approverEmail: s.approverEmail,  // ✅
        approvalRole: cleanDisplayText(s.approvalRole),
        comments: cleanDisplayText(s.reason),
      };
    });

  }, [
    parsedMultipleApproverData,
    exportData?.approverName,
    exportData?.approverNames,
    exportData?.approverPosition,
    exportData?.workflowStatus,
    exportData?.approvalStatus,
    exportData?.approvedDate,
    exportData?.rejectReason,
    exportData?.currentApproverEmail,
  ]);


  React.useEffect(() => {
    const resolveNames = async (): Promise<void> => {
      if (!footerApprovers.length) return;

      const webUrl = context.pageContext.web.absoluteUrl;

      const emailsToResolve = footerApprovers
        .filter((a) => {
          const hasEmail = !!cleanDisplayText(a.approverEmail || "");
          if (!hasEmail) return false;

          const emailKey = cleanDisplayText(a.approverEmail || "").toLowerCase();
          const alreadyResolved = !!resolvedApproverNames[emailKey];
          if (alreadyResolved) return false;

          // ✅ Resolve if name is empty
          const nameEmpty = !cleanDisplayText(a.name);
          if (nameEmpty) return true;

          // ✅ Also resolve if name looks like a role label (equals approvalRole)
          // This catches the "C1 Level" / "C2 Level" case
          const nameIsRoleLabel =
            cleanDisplayText(a.name).toLowerCase() ===
            cleanDisplayText(a.approvalRole || "").toLowerCase();
          if (nameIsRoleLabel) return true;

          return false;
        })
        .map((a) => cleanDisplayText(a.approverEmail || "").toLowerCase())
        .filter(Boolean);

      const uniqueEmails = Array.from(new Set(emailsToResolve));
      if (!uniqueEmails.length) return;

      const newNames: Record<string, string> = {};

      for (const email of uniqueEmails) {
        try {
          const safeEmail = email.replace(/'/g, "''");

          // ✅ Try ensureUser first
          const ensureUrl = `${webUrl}/_api/web/ensureUser('${safeEmail}')`;
          const res = await context.spHttpClient.post(
            ensureUrl,
            SPHttpClient.configurations.v1,
            {
              headers: {
                "Content-Type": "application/json;odata=verbose",
                Accept: "application/json;odata=nometadata",
                "odata-version": "",
              },
              body: JSON.stringify({}),
            }
          );

          if (res.ok) {
            const json = await res.json() as { Title?: string };
            const displayName = cleanDisplayText(json.Title || "");
            if (displayName) {
              newNames[email] = displayName;
              continue;
            }
          }

          // ✅ Fallback: siteusers
          const siteUserUrl =
            `${webUrl}/_api/web/siteusers` +
            `?$select=Title,Email` +
            `&$filter=Email eq '${safeEmail}'` +
            `&$top=1`;

          const fallbackRes = await context.spHttpClient.get(
            siteUserUrl,
            SPHttpClient.configurations.v1
          );

          if (fallbackRes.ok) {
            const fallbackJson = await fallbackRes.json() as {
              value?: Array<{ Title?: string }>;
            };
            const fallbackName = cleanDisplayText(fallbackJson.value?.[0]?.Title || "");
            if (fallbackName) {
              newNames[email] = fallbackName;
            }
          }
        } catch {
          // silently skip
        }
      }

      if (Object.keys(newNames).length > 0) {
        setResolvedApproverNames((prev) => ({ ...prev, ...newNames }));
      }
    };

    resolveNames().catch(() => undefined);
  }, [footerApprovers, context, resolvedApproverNames]);

  // Fetch approver job positions from Azure AD / Entra ID (via SP User Profile service)
  React.useEffect(() => {
    const fetchApproverPositions = async (): Promise<void> => {
      const uniqueEmails = Array.from(
        new Set(
          footerApprovers
            .map((a) => cleanDisplayText(a.approverEmail || "").toLowerCase())
            .filter((e) => !!e && !fetchedApproverEmails.current.has(e))
        )
      );
      if (!uniqueEmails.length) return;

      // Mark as in-flight before async calls to prevent duplicate fetches
      uniqueEmails.forEach((e) => fetchedApproverEmails.current.add(e));

      const newPositions: Record<string, string> = {};
      for (const email of uniqueEmails) {
        try {
          const profile = await fetchSPUserProfile(context, email);
          if (profile.jobTitle) {
            newPositions[email] = profile.jobTitle;
          } else {
            // Fallback: query EForm Users list for actual Jobposition
            const eformPos = await fetchEFormUserJobPosition(context, email);
            if (eformPos) newPositions[email] = eformPos;
          }
        } catch {
          // silently skip
        }
      }

      if (Object.keys(newPositions).length > 0) {
        setApproverAadPositions((prev) => ({ ...prev, ...newPositions }));
      }
    };

    fetchApproverPositions().catch(() => undefined);
  }, [footerApprovers, context]);

  const runActionAndNavigate = React.useCallback(
    async (action: PendingAction, comments = ""): Promise<void> => {
      setActionSubmitting(true);
      try {
        let successMessage = "";

        if (action === "Approve") {
          if (!onApprove) throw new Error("Approve action is unavailable.");
          successMessage = await onApprove(comments); // ✅ use PA message
        }
        if (action === "Reject") {
          if (!onReject) throw new Error("Reject action is unavailable.");
          successMessage = await onReject(comments);  // ✅ use PA message
        }
        if (action === "Bypass") {
          if (!onBypass) throw new Error("Bypass action is unavailable.");
          successMessage = await onBypass(comments);  // ✅ use PA message
        }

        onNavigate("team", { toastType: "success", toastMessage: successMessage });
      } catch (err: unknown) {
        const message =
          err instanceof Error && err.message.trim() ? err.message : "Action failed";
        onNavigate("team", { toastType: "error", toastMessage: message });
      } finally {
        setActionSubmitting(false);
        setActionModalOpen(false);
        setPendingAction(null);
        setActionComments("");
      }
    },
    [onApprove, onReject, onBypass, onNavigate]
  );

  const openActionModal = (action: PendingAction): void => {
    setPendingAction(action);
    setActionComments("");
    setActionModalOpen(true);
  };

  const closeActionModal = (): void => {
    if (actionSubmitting) return;
    setActionModalOpen(false);
    setPendingAction(null);
    setActionComments("");
  };

  const confirmAction = (): void => {
    if (!pendingAction) return;
    if (pendingAction === "Reject" && !actionComments.trim()) {
      setToast({ type: "error", message: "Rejection reason is required." });
      return;
    }
    runActionAndNavigate(pendingAction, actionComments.trim()).catch((e) => { console.error(e); });
  };

  const handleApprove = async (): Promise<void> => {
    if (actionHandlers?.validateBeforeApprove) {
      const { ok, invalidItems } = await actionHandlers.validateBeforeApprove();
      if (!ok && invalidItems.length > 0) {
        setEditRows(invalidItems);
        setEditQtyPopupOpen(true);
        return;
      }
    }
    openActionModal("Approve");
  };

  const handleReject = (): void => { openActionModal("Reject"); };
  const handleBypass = (): void => { openActionModal("Bypass"); };

  const applyQtyAndApprove = async (): Promise<void> => {
    try {
      for (const row of editRows) {
        await actionHandlers?.updateInvalidQuantity(row.itemId, row.qty);
      }
      const res = await actionHandlers?.submitDecision("Approve", "");
      if (res?.success) {
        setToast({ type: "success", message: res.message });
        setEditQtyPopupOpen(false);
        await actionHandlers?.reloadAfterAction?.();
      } else {
        setToast({ type: "error", message: res?.message || "Approve failed" });
      }
    } catch (err) {
      setToast({ type: "error", message: "Error updating quantities" });
    }
  };

  React.useEffect(() => {
    const loadLines = async (): Promise<void> => {
      try {
        setLineError("");
        setLineLoading(true);
        setLineItems([]);
        if (!spId) return;

        const webUrl = context.pageContext.web.absoluteUrl;
        const childListTitle = "InventoryRequestItems";
        const requestLookupInternal = "RequestID";
        const colTitle = "Title";
        const colLineNumber = "LineNumber";
        const colSku = "SKUCode";
        const colDesc = "ProductDescription";
        const colColor = "Color";
        const colSize = "Size";
        const colUnitPrice = "UnitPrice";
        const colQty = "Quantity";
        const colLineTotal = "LineTotal";
        const colLookupSource = "LookupSource";
        const colRowStatus = "RowStatus";

        type ChildRow = Record<string, unknown> & { Id?: number };
        type SpListResponse<T> = { value?: T[]; "@odata.nextLink"?: string };

        const tryFetch = async (url: string): Promise<SpListResponse<ChildRow>> => {
          const r = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
          if (!r.ok) {
            const body = await r.text();
            throw new Error(`SharePoint ${r.status}: ${body}`);
          }
          return (await r.json()) as SpListResponse<ChildRow>;
        };

        // ✅ Paginated fetch — keeps following @odata.nextLink until exhausted
        const fetchAllPages = async (firstUrl: string): Promise<ChildRow[]> => {
          const allRows: ChildRow[] = [];
          let nextUrl: string | undefined = firstUrl;

          while (nextUrl) {
            const page: SpListResponse<ChildRow> = await tryFetch(nextUrl);
            const pageRows = Array.isArray(page.value) ? page.value : [];
            allRows.push(...pageRows);
            nextUrl = page["@odata.nextLink"] || undefined;
          }

          return allRows;
        };

        const selectFields = [
          "Id",
          colTitle,
          colLineNumber,
          colSku,
          colDesc,
          colColor,
          colSize,
          colUnitPrice,
          colQty,
          colLineTotal,
          colLookupSource,
          colRowStatus,
          `${requestLookupInternal}/Id`,
        ].join(",");

        // ✅ $top=5000 is the SharePoint max per page — nextLink handles the rest
        const urlLookup =
          `${webUrl}/_api/web/lists/getbytitle('${childListTitle}')/items` +
          `?$select=${selectFields}` +
          `&$expand=${requestLookupInternal}` +
          `&$filter=${requestLookupInternal}/Id eq ${spId} and ${colRowStatus} eq 'Active'` +
          `&$orderby=${colLineNumber} asc` +
          `&$top=5000`;

        let rowsRaw: ChildRow[];

        try {
          rowsRaw = await fetchAllPages(urlLookup);
        } catch {
          // Fallback: try without expand (use numeric lookup field directly)
          const urlNumber =
            `${webUrl}/_api/web/lists/getbytitle('${childListTitle}')/items` +
            `?$select=Id,${colTitle},${colLineNumber},${colSku},${colDesc},${colColor},${colSize},${colUnitPrice},${colQty},${colLineTotal},${colLookupSource},${colRowStatus},${requestLookupInternal}` +
            `&$filter=${requestLookupInternal} eq ${spId} and ${colRowStatus} eq 'Active'` +
            `&$orderby=${colLineNumber} asc` +
            `&$top=5000`;

          rowsRaw = await fetchAllPages(urlNumber);
        }

        const rows: LineItem[] = rowsRaw.map((x): LineItem => {
          const id = Number(x.Id ?? 0);
          const title = String(x[colTitle] ?? "");
          const lineNumber = Number(x[colLineNumber] ?? 0);
          const skuCode = String(x[colSku] ?? "");
          const productDescription = String(x[colDesc] ?? "");
          const color = String(x[colColor] ?? "");
          const size = String(x[colSize] ?? "");
          const unitPrice = Number(x[colUnitPrice] ?? 0);
          const quantity = Number(x[colQty] ?? 0);
          const lineTotalRaw = x[colLineTotal];
          const lineTotal =
            typeof lineTotalRaw === "number"
              ? lineTotalRaw
              : Number(lineTotalRaw ?? unitPrice * quantity);
          const lookupSource = String(x[colLookupSource] ?? "");
          const rowStatus = String(x[colRowStatus] ?? "");

          return {
            id,
            title,
            lineNumber,
            skuCode,
            productDescription,
            color,
            size,
            unitPrice,
            quantity,
            lineTotal,
            lookupSource,
            rowStatus,
          };
        });

        setLineItems(rows);
      } catch (e: unknown) {
        setLineError(e instanceof Error ? e.message : String(e));
      } finally {
        setLineLoading(false);
      }
    };

    loadLines().catch(() => undefined);
  }, [context, spId]);

  type JsPdfCtor = new (opts: { orientation: "p" | "l"; unit: "mm"; format: "a4" }) => {
    internal: { pageSize: { getWidth: () => number; getHeight: () => number } };
    addImage: (data: string, format: "PNG", x: number, y: number, w: number, h: number) => void;
    addPage: () => void;
    save: (name: string) => void;
  };

  const handleExportPdf = async (): Promise<void> => {
    if (!canExportPdf) {
      alert("You are not allowed to export this report.");
      if (autoExport) onAutoExportDone?.();
      return;
    }
    const el = pdfRef.current;
    if (!el) { if (autoExport) onAutoExportDone?.(); return; }

    try {
      setPdfBusy(true);
      const htmlToImage = await import(/* webpackChunkName: "html-to-image" */ "html-to-image");
      const jspdfMod = await import(/* webpackChunkName: "jspdf" */ "jspdf");
      const JsPDF: JsPdfCtor =
        ((jspdfMod as unknown as { jsPDF?: JsPdfCtor }).jsPDF ??
          (jspdfMod as unknown as { default?: JsPdfCtor }).default) as JsPdfCtor;

      const prevOverflow = el.style.overflow;
      const prevHeight = el.style.height;
      const prevMaxHeight = el.style.maxHeight;
      el.style.overflow = "visible";
      el.style.height = "auto";
      el.style.maxHeight = "none";
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));

      const orientation: "p" | "l" = "p";
      const pdf = new JsPDF({ orientation, unit: "mm", format: "a4" });
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();

      try {
        const exportRoot = pdfPagesRef.current ?? el;
        const pageNodes = Array.from(exportRoot.querySelectorAll<HTMLElement>('[data-pdf-page="true"]'));
        const exportNodes = pageNodes.length > 0 ? pageNodes : [el];

        for (let i = 0; i < exportNodes.length; i += 1) {
          const node = exportNodes[i];
          const canvas = await htmlToImage.toCanvas(node, {
            pixelRatio: 2, cacheBust: true, backgroundColor: "#ffffff",
          });
          const imgData = canvas.toDataURL("image/png");
          const scale = Math.min(pageWidth / canvas.width, pageHeight / canvas.height);
          const imgWidth = canvas.width * scale;
          const imgHeight = canvas.height * scale;
          const x = (pageWidth - imgWidth) / 2;
          const y = 0;
          if (i > 0) pdf.addPage();
          pdf.addImage(imgData, "PNG", x, y, imgWidth, imgHeight);
        }
      } finally {
        el.style.overflow = prevOverflow;
        el.style.height = prevHeight;
        el.style.maxHeight = prevMaxHeight;
      }

      const safeTxnId = (exportData?.Title || String(spId ?? "NA")).replace(/\//g, "-");
      pdf.save(`${safeTxnId}.pdf`);
    } catch (e: unknown) {
      console.error("PDF export failed:", e);
      alert("PDF export failed. Check console.");
    } finally {
      setPdfBusy(false);
      if (autoExport) onAutoExportDone?.();
    }
  };

  React.useEffect(() => {
    if (!autoExport) return;
    if (autoExportTriggeredRef.current) return;
    if (!spId) return;
    if (!exportData) return;
    if (lineLoading) return;
    autoExportTriggeredRef.current = true;
    handleExportPdf().catch(() => undefined);
  }, [autoExport, spId, exportData, lineLoading]);

  const placeToDelivery = exportData?.placeOfDelivery?.trim() || "";
  const FIRST_PAGE_ROWS = 15;
  const NEXT_PAGE_ROWS = 15;
  const PDF_PREVIEW_PAGE_MIN_HEIGHT = 980;

  const lineItemPages = React.useMemo<LineItem[][]>(() => {
    if (lineItems.length === 0) return [[]];
    const pages: LineItem[][] = [];
    pages.push(lineItems.slice(0, FIRST_PAGE_ROWS));
    for (let i = FIRST_PAGE_ROWS; i < lineItems.length; i += NEXT_PAGE_ROWS) {
      pages.push(lineItems.slice(i, i + NEXT_PAGE_ROWS));
    }
    return pages;
  }, [lineItems, FIRST_PAGE_ROWS, NEXT_PAGE_ROWS]);

  const pageRowOffsets = React.useMemo<number[]>(() => {
    const offsets: number[] = [];
    let offset = 0;
    for (const pageRows of lineItemPages) {
      offsets.push(offset);
      offset += pageRows.length;
    }
    return offsets;
  }, [lineItemPages]);

  const totalPdfPages = lineItemPages.length;
  const isSinglePdfPage = totalPdfPages === 1;
  const isRejectModal = pendingAction === "Reject";
  const rejectReasonEmpty = isRejectModal && !actionComments.trim();
  const confirmDisabled = actionSubmitting || rejectReasonEmpty;
  const totalQty = lineItems.reduce((s, x) => s + x.quantity, 0);
  const totalPrice = lineItems.reduce((s, x) => s + x.lineTotal, 0);

  const renderItemsTable = (
    rows: LineItem[],
    rowOffset: number,
    showGrandTotal: boolean
  ): JSX.Element => (
    <div className="spfx-table-wrapper">
      <table className={isDesktopTable ? "w-full border-collapse table-auto text-[13px]" : "w-full border-collapse table-auto text-[8px] sm:text-[12px]"}>
        <thead className="bg-gray-100 text-gray-700">
          <tr className="border border-gray-300 border-collapse">
            <th className={isDesktopTable ? "px-3 py-4 w-[44px] text-left align-top" : "px-1 sm:px-2 py-2 sm:py-3 w-[36px] sm:w-[44px] text-left align-top"}>
              <div className={isDesktopTable ? "text-[12px] font-bold text-black" : "text-[8px] sm:text-[11px] font-bold text-black"}>ลำดับ</div>
              <div className={isDesktopTable ? "text-sm font-bold" : "text-[8px] sm:text-xs font-bold"}>NO.</div>
            </th>
            <th className={isDesktopTable ? "px-3 py-3 text-left align-top break-words" : "px-1 sm:px-2 py-2 text-left align-top break-words"}>
              <div className={isDesktopTable ? "text-[12px] font-bold text-black" : "text-[8px] sm:text-[11px] font-bold text-black"}>รหัสสินค้า</div>
              <div className={isDesktopTable ? "text-sm font-bold" : "text-[8px] sm:text-xs font-bold"}>SKU Code</div>
            </th>
            <th className={isDesktopTable ? "px-3 py-3 text-left align-top break-words" : "px-1 sm:px-2 py-2 text-left align-top break-words"}>
              <div className={isDesktopTable ? "text-[12px] font-bold text-black" : "text-[8px] sm:text-[11px] font-bold text-black"}>กลุ่มสินค้า</div>
              <div className={isDesktopTable ? "text-sm font-bold" : "text-[8px] sm:text-xs font-bold"}>Product Code</div>
            </th>
            <th className={isDesktopTable ? "px-3 py-3 text-left align-top break-words" : "px-1 sm:px-2 py-2 text-left align-top break-words"}>
              <div className={isDesktopTable ? "text-[12px] font-bold text-black" : "text-[8px] sm:text-[11px] font-bold text-black"}>สี</div>
              <div className={isDesktopTable ? "text-sm font-bold" : "text-[8px] sm:text-xs font-bold"}>Color</div>
            </th>
            <th className={isDesktopTable ? "px-3 py-3 text-left align-top break-words" : "px-1 sm:px-2 py-2 text-left align-top break-words"}>
              <div className={isDesktopTable ? "text-[12px] font-bold text-black" : "text-[8px] sm:text-[11px] font-bold text-black"}>ขนาด</div>
              <div className={isDesktopTable ? "text-sm font-bold" : "text-[8px] sm:text-xs font-bold"}>Size</div>
            </th>
            <th className={isDesktopTable ? "px-3 py-3 text-right align-top break-words" : "px-1 sm:px-2 py-2 text-right align-top break-words"}>
              <div className={isDesktopTable ? "text-[12px] font-bold text-black" : "text-[8px] sm:text-[11px] font-bold text-black"}>ราคาป้ายต่อหน่วย</div>
              <div className={isDesktopTable ? "text-sm font-bold" : "text-[8px] sm:text-xs font-bold"}>Tag Price Per Unit</div>
            </th>
            <th className={isDesktopTable ? "px-3 py-3 text-right align-top break-words" : "px-1 sm:px-2 py-2 text-right align-top break-words"}>
              <div className={isDesktopTable ? "text-[12px] font-bold text-black" : "text-[8px] sm:text-[11px] font-bold text-black"}>จำนวน</div>
              <div className={isDesktopTable ? "text-sm font-bold" : "text-[8px] sm:text-xs font-bold"}>Quantity</div>
            </th>
            <th className={isDesktopTable ? "px-3 py-3 text-right align-top break-words" : "px-1 sm:px-2 py-2 text-right align-top break-words"}>
              <div className={isDesktopTable ? "text-[12px] font-bold text-black" : "text-[8px] sm:text-[11px] font-bold text-black"}>ราคาป้ายทั้งหมด</div>
              <div className={isDesktopTable ? "text-sm font-bold" : "text-[8px] sm:text-xs font-bold"}>Total Tag Price</div>
            </th>
          </tr>
        </thead>
        <tbody className="text-gray-800 font-medium">
          {rowOffset === 0 && lineError && (
            <tr><td colSpan={8} className="px-3 py-4 text-sm text-red-600">{lineError}</td></tr>
          )}
          {rowOffset === 0 && lineLoading && (
            <tr><td colSpan={8} className="px-3 py-4 text-sm">Loading items...</td></tr>
          )}
          {rowOffset === 0 && !lineLoading && !lineError && lineItems.length === 0 && (
            <tr><td colSpan={8} className="px-3 py-4 text-sm text-gray-500">No items found for Request ID: {spId}</td></tr>
          )}
          {rows.map((r, idx) => (
            <tr key={r.id} className="border border-gray-300 border-collapse">
              <td className={isDesktopTable ? "px-3 py-4 align-top" : "px-1 sm:px-2 py-2 sm:py-3 align-top"}>{rowOffset + idx + 1}.</td>
              <td className={isDesktopTable ? "px-3 py-3 align-top break-all" : "px-1 sm:px-2 py-2 align-top break-all"}>{r.skuCode || r.title}</td>
              <td className={isDesktopTable ? "px-3 py-3 align-top break-all" : "px-1 sm:px-2 py-2 align-top break-all"}>{r.productDescription}</td>
              <td className={isDesktopTable ? "px-3 py-3 align-top break-words" : "px-1 sm:px-2 py-2 align-top break-words"}>{r.color}</td>
              <td className={isDesktopTable ? "px-3 py-3 align-top break-words" : "px-1 sm:px-2 py-2 align-top break-words"}>{r.size}</td>
              <td className={isDesktopTable ? "px-3 py-3 text-right align-top" : "px-1 sm:px-2 py-2 text-right align-top"}>
                {r.unitPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </td>
              <td className={isDesktopTable ? "px-3 py-3 text-right align-top" : "px-1 sm:px-2 py-2 text-right align-top"}>{r.quantity}</td>
              <td className={isDesktopTable ? "px-3 py-3 text-right align-top" : "px-1 sm:px-2 py-2 text-right align-top"}>
                {r.lineTotal.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </td>
            </tr>
          ))}
          {showGrandTotal && (
            <tr className="border-t border-gray-300">
              <td colSpan={6} className="px-3 py-3 text-center font-semibold text-gray-900 bg-[#BCC8D6]">
                ยอดรวมทั้งหมด / Grand Total
              </td>
              <td className="px-3 py-3 text-right font-bold text-gray-900 bg-[#CAD6E6]">{totalQty}</td>
              <td className="px-3 py-3 text-right font-semibold text-gray-900 bg-[#E0EEFF]">
                {totalPrice.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );

  // =========================================================
  // ✅ UPDATED renderFooterSection
  //    - Grid columns: 1 requestor + N approver slots
  //    - Name: approverName if set, else approverEmail as fallback
  //    - Position: approvalRole ("C1 Level") preferred over job position
  //    - Shows acted slots with data, unacted slots with empty placeholders
  // =========================================================
  const renderFooterSection = (): JSX.Element => {
    const totalColumns = 1 + footerApprovers.length;
    const gridClass =
      totalColumns === 2 ? "grid-cols-2"
        : totalColumns === 3 ? "grid-cols-3"
          : totalColumns >= 4 ? "grid-cols-4"
            : "grid-cols-2";

    return (
      <div className="mt-6 text-[13px] text-gray-700">
        <div className={`grid gap-10 ${gridClass}`}>

          {/* ── Requestor column ── */}
          <div>
            <p className="font-semibold mb-3">ผู้เบิก / Requestor</p>
            <p className="mb-1">
              <span className="text-gray-400">ชื่อ / Name :</span>{" "}
              <span className="text-gray-700">{resolvedRequestorName || "-"}</span>
            </p>
            <p className="mb-1">
              <span className="text-gray-400">ตำแหน่ง / Position :</span>{" "}
              <span className="text-gray-700">{resolvedRequestorPosition}</span>
            </p>
            <p>
              <span className="text-gray-400">Submitted Date :</span>{" "}
              <span className="text-gray-700">{exportData?.submittedDate || "-"}</span>
            </p>
          </div>

          {/* ── Approver columns ── */}
          {footerApprovers.map((approver) => {
            const approverEmailKey = cleanDisplayText(
              approver.approverEmail || ""
            ).toLowerCase();

            // ✅ Priority: approverName from JSON → resolved from SP → raw email
            const displayName =
              cleanDisplayText(approver.name) ||
              cleanDisplayText(resolvedApproverNames[approverEmailKey] || "") ||
              cleanDisplayText(approver.approverEmail || "") ||
              "-";

            // AAD job title takes priority; fall back to MultipleApproverData position field
            const aadPosition = cleanDisplayText(approverAadPositions[approverEmailKey] || "");
            const displayPosition =
              aadPosition ||
              cleanDisplayText(approver.position) ||
              "-";

            return (
              <div key={`footer-approver-${approver.slot}`}>
                <p className="font-semibold mb-3">
                  {approver.slot}) ผู้อนุมัติ / Approver
                </p>

                {approver.status === "Approved" || approver.status === "Rejected" ? (
                  <>
                    <p className="mb-1">
                      <span className="text-gray-400">ชื่อ / Name :</span>{" "}
                      <span>{displayName}</span>
                    </p>

                    <p className="mb-1">
                      <span className="text-gray-400">ตำแหน่ง / Position :</span>{" "}
                      <span>{displayPosition}</span>
                    </p>

                    <p className="mb-1">
                      <span className="text-gray-400">Approval :</span>{" "}
                      <span className={`${getSubmitDisplay(approver.status).className} font-bold`}>
                        {getSubmitDisplay(approver.status).label}
                      </span>
                    </p>

                    {approver.status === "Rejected" && (
                      <p className="mb-1">
                        <span className="text-gray-400">Reject Reason :</span>{" "}
                        <span>{cleanDisplayText(approver.reason) || "-"}</span>
                      </p>
                    )}

                    <p>
                      <span className="text-gray-400">Approved Date :</span>{" "}
                      <span>{approver.date}</span>
                    </p>

                    {!!approver.bypassBy && (
                      <p className="text-xs text-gray-500 mt-1">
                        (ดำเนินการแทนโดย {approver.bypassBy})
                      </p>
                    )}
                  </>
                ) : (
                  <div className="text-gray-400 space-y-1">
                    <p>ชื่อ / Name :</p>
                    <p>ตำแหน่ง / Position :</p>
                    <p>Approval :</p>
                    <p>Approved Date :</p>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  return (
    <div className="pdf-thai flex flex-col h-full overflow-hidden">
      <div className="sticky top-0 z-20 bg-white shrink-0">
        {toast && (
          <NotificationBar
            type={toast.type}
            message={toast.message}
            onClose={() => setToast(null)}
          />
        )}
        <ExportPDFHeader
          transactionId={exportData?.Title}
          exportSource={exportSource}
          canAct={canAct}
          isApproveMode={isApproveMode}
          isBypassMode={isBypassMode}
          showReject={showReject}
          isApproverWaiting={isApproverWaiting}
          canExportPdf={!!canExportPdf}
          pdfBusy={pdfBusy}
          permissions={permissions}
          onNavigate={onNavigate}
          onApprove={handleApprove}
          onReject={handleReject}
          onBypass={handleBypass}
          onExportPdf={handleExportPdf}
        />
      </div>

      <div className="flex-1 overflow-auto bg-[#f6f6f6] px-8 sm:px-6 lg:px-8 py-1">
        <div ref={pdfPagesRef} className="w-full">
          {error && <div className="text-red-600 text-sm mb-4">{error}</div>}
          {!error && !exportData && <div className="text-sm mb-4">Loading...</div>}

          <div
            ref={pdfRef}
            data-pdf-page="true"
            className={`bg-white border border-[#e8e8e8] px-10 sm:px-6 py-10 sm:py-6 ${isSinglePdfPage ? "" : "flex flex-col"}`}
            style={isSinglePdfPage ? undefined : { minHeight: `${PDF_PREVIEW_PAGE_MIN_HEIGHT}px` }}
          >
            <div className="flex justify-between items-center mb-2">
              <div className="flex items-center gap-4">
                <img src={LogoPng} alt="E-FORM" crossOrigin="anonymous" className="h-14 w-auto object-contain" />
                <div>
                  <p className="font-semibold text-gray-900 text-xl leading-tight">
                    บริษัท กิติรัญ จำกัด / Kititan Company Limited
                  </p>
                  <p className="font-medium text-sm text-gray-500">
                    แบบเบิกสินค้าคงเหลือเพื่อใช้ภายใน / Internal Usage Inventory Request Form
                  </p>
                </div>
              </div>
              <div className="flex flex-col items-start">
                {(() => {
                  const status = getStatusDisplay(exportData?.workflowStatus);
                  return (
                    <>
                      <div className="flex items-center gap-2">
                        <span className="text-lg font-semibold text-gray-900">Status :</span>
                        <span className={`text-lg font-semibold ${status.className}`}>{status.label}</span>
                      </div>
                      <div className="mt-1 text-sm font-medium text-gray-500">
                        Transaction ID : {exportData?.Title || ""}
                      </div>
                    </>
                  );
                })()}
              </div>
            </div>

            <hr className="my-6 h-px border-0 bg-gray-200" />

            <div className="grid grid-cols-2 gap-x-12 gap-y-3 text-sm leading-6 text-gray-600 font-medium mb-6">
              <div>
                <span className="font-medium text-gray-500">ชื่อผู้เบิก / Requestor :</span>{" "}
                <span className="font-medium text-gray-900">{exportData?.requester || " "}</span>
              </div>
              <div>
                <span className="font-medium text-gray-500">ตำแหน่งงาน / Job Position :</span>{" "}
                <span className="font-medium text-gray-900">{resolvedRequestorPosition}</span>
              </div>
              <div>
                <span className="font-medium text-gray-500">แผนก / Department :</span>{" "}
                <span className="font-medium text-gray-900">{resolvedRequestorDepartment || " "}</span>
              </div>
              <div>
                <span className="font-medium text-gray-500">ศูนย์ต้นทุน / Cost Center :</span>{" "}
                <span className="font-medium text-gray-900">{exportData?.costCenter || " "}</span>
              </div>
              <div>
                <span className="font-medium text-gray-500">วันที่ต้องการสินค้า / Expected Inventory Date :</span>{" "}
                <span className="font-medium text-gray-900">{exportData?.expectedDeliveryDate || " "}</span>
              </div>
              <div className="col-span-2">
                <div className="font-medium text-gray-500">สถานที่จัดส่งสินค้า / Place to Delivery :</div>
                <div className="mt-1 break-words font-medium text-gray-900">{placeToDelivery}</div>
              </div>
              <div className="col-span-2">
                <span className="font-medium text-gray-500">วัตถุประสงค์เบิก / Request Purpose :</span>{" "}
                <span className="font-medium text-gray-900">{exportData?.purpose || " "}</span>
              </div>
              <div className="col-span-2">
                <div className="font-medium text-gray-500">รายละเอียดในการเบิกสินค้าคงเหลือ / Purpose Detail :</div>
                <div className="mt-1 break-words font-medium text-gray-900">{exportData?.purposeDetails || " "}</div>
              </div>
            </div>

            {renderItemsTable(lineItemPages[0] ?? [], pageRowOffsets[0] ?? 0, totalPdfPages === 1)}
            {totalPdfPages === 1 && renderFooterSection()}
            <div className={`text-right text-sm font-semibold text-gray-700 ${isSinglePdfPage ? "mt-4" : "mt-auto pt-4"}`}>
              Page 1/{totalPdfPages}
            </div>
          </div>

          {totalPdfPages > 1 && <div className="h-10 bg-[#f6f6f6]" />}

          {lineItemPages.slice(1).map((rows, index) => {
            const pageNo = index + 2;
            const isLast = pageNo === totalPdfPages;
            const rowOffset = pageRowOffsets[index + 1] ?? 0;
            return (
              <React.Fragment key={`pdf-page-${pageNo}`}>
                <div
                  data-pdf-page="true"
                  className="bg-white border border-[#e8e8e8] px-10 sm:px-6 py-10 sm:py-6 flex flex-col"
                  style={{ minHeight: `${PDF_PREVIEW_PAGE_MIN_HEIGHT}px` }}
                >
                  <div className="mb-4 flex justify-end text-sm font-medium text-gray-500">
                    Transaction ID : {exportData?.Title || ""}
                  </div>
                  {renderItemsTable(rows, rowOffset, isLast)}
                  {isLast && renderFooterSection()}
                  <div className="mt-auto pt-4 text-right text-sm font-semibold text-gray-700">
                    Page {pageNo}/{totalPdfPages}
                  </div>
                </div>
                {!isLast && <div className="h-10 bg-[#f6f6f6]" />}
              </React.Fragment>
            );
          })}
        </div>
      </div>

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
              {pendingAction === "Approve" ? "Do you want to approve the request?" : "Do you want to reject the request?"}
            </h2>

            <p className={isRejectModal ? "mt-2 text-center text-[13px] font-semibold leading-tight text-black" : "mt-3 text-center text-[18px] font-semibold leading-tight text-black"}>
              ID &ldquo;{exportData?.Title || spId || "-"}&rdquo;
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
                    ? "h-[38px] w-[96px] rounded-[6px] border border-[#E8E8E8] bg-[#E8E8E8] text-[16px] leading-none font-medium text-[#474747] hover:bg-[#EBEBEB] disabled:opacity-50"
                    : "h-[46px] rounded-[10px] border border-[#E8E8E8] bg-[#E8E8E8] text-[16px] leading-none font-semibold text-[#494949] hover:bg-[#EBEBEB] disabled:opacity-50"
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
                        ? "h-[38px] w-[96px] rounded-[6px] bg-[#DA2128] text-[16px] leading-none font-semibold text-white hover:bg-[#AF0007]"
                        : "h-[46px] rounded-[10px] bg-[#DA2128] text-[16px] leading-none font-semibold text-white hover:bg-[#AF0007]"
                }
                onClick={confirmAction}
                disabled={confirmDisabled}
              >
                {actionSubmitting
                  ? pendingAction === "Approve" ? "Approving..."
                    : pendingAction === "Reject" ? "Rejecting..."
                      : "Bypassing..."
                  : pendingAction}
              </button>

            </div>
          </div>
        </div>
      )}

      {editQtyPopupOpen && (
        <div className="qty-popup">
          <h3>Correct Quantities</h3>
          {editRows.map((row, i) => (
            <div key={i}>
              <span>{row.sku}</span>
              <input
                type="number"
                value={row.qty}
                onChange={(e) => {
                  const next = [...editRows];
                  next[i].qty = Number(e.target.value);
                  setEditRows(next);
                }}
              />
            </div>
          ))}
          <button onClick={() => setEditQtyPopupOpen(false)}>Cancel</button>
          <button onClick={applyQtyAndApprove}>Save & Approve</button>
        </div>
      )}
    </div>
  );
}