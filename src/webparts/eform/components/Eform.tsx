import * as React from "react";
import styles from "./Eform.module.scss";

import ErrorBoundary from "./ErrorBoundary";

import AppShell from "./layout/AppShell";
import HomePage from "./pages/HomePage";

import CreateRequestPage from "./pages/CreateRequestPage/CreateRequestPage";
import ExportPDFPage from "./pages/ExportPDFPage";
import TeamRequestsPage from "./pages/TeamRequest/TeamRequestsPage";

import { IEformProps } from "./IEformProps";
import { SPHttpClient } from "@microsoft/sp-http";
import MyRequestsPage from "./pages/MyRequestsPage";
import Loading from "./ui/Loading";

import AllRequestPage from "./pages/AllRequestPage/AllRequestPage";
import { RouteState, SecurityPermissions } from "./spHelpers";

import { ExportActionInfo } from "./pages/ExportActionInfo";
import { submitApprovalDecision } from "./pages/ApprovalDecisionService";

import { getUserDataFromSharePoint } from "./pages/CreateRequestPage/CreateRequestService";

export type Route = "home" | "create" | "exportPdf" | "team" | "all" | "my";
export type ExportSource = "my" | "team" | "all";

const canAccessRoute = (
  perms: SecurityPermissions | undefined,
  route: Route
): boolean => {
  if (!perms) return false;
  if (route === "home" || route === "exportPdf") return true;
  if (route === "create") return perms.submit;
  if (route === "my") return perms.viewOwn;
  if (route === "team") return perms.viewTeam;
  if (route === "all") return perms.viewAll;
  return false;
};

// ── Parse deep-link params once at module level ──
// Accepts requestId / RequestId (case-insensitive) so both the Power Automate
// generated URL (?requestId=N) and any manually constructed links work.
const _dlParams = (() => {
  try {
    const p = new URLSearchParams(window.location.search);
    const rid = p.get("requestId") || p.get("RequestId");
    if (!rid) return undefined;
    const n = parseInt(rid, 10);
    if (isNaN(n) || n <= 0) return undefined;
    const tid = p.get("transactionId") || p.get("TransactionId") || undefined;
    return { requestId: n, transactionId: tid };
  } catch { /* ignore */ }
  return undefined;
})();

// ── Parse hash-based deep-link route (my / team / all / create) ──
// SPFx workbench runs the web part inside a same-origin iframe, so
// window.location is the iframe URL — the address-bar URL lives in window.top.
const _HASH_ROUTES: Record<string, Route> = { "/my": "my", "/team": "team", "/all": "all", "/create": "create" };
const _getTopHash = (): string => {
  try { return window.top?.location?.hash ?? window.location.hash; }
  catch { return window.location.hash; }
};
const _hashRoute = (() => {
  try {
    const h = _getTopHash();
    const path = h.startsWith("#") ? h.slice(1) : h;
    return _HASH_ROUTES[path] as Route | undefined;
  } catch { return undefined; }
})();
// Hash route is ignored when a PDF deep-link query param is present
const _initialHashRoute = _dlParams ? undefined : _hashRoute;

const Eform: React.FC<IEformProps> = ({ context, flowUrl, ApprovalFlowUrl }) => {
  // Always start at "home" when a deep-link is present.
  // The validation effect below checks access and navigates to exportPdf (or shows denied).
  const [route, setRoute] = React.useState<Route>(
    _dlParams ? "home" : (_initialHashRoute || "home")
  );
  const [selectedSpId, setSelectedSpId] = React.useState<number | undefined>(undefined);
  const [dlAccessDenied, setDlAccessDenied] = React.useState(false);
  const [exportSource, setExportSource] = React.useState<ExportSource>("my");
  const [quickPdfRequest, setQuickPdfRequest] = React.useState<{ spId: number; source: ExportSource } | undefined>(undefined);
  const [exportActionInfo, setExportActionInfo] = React.useState<ExportActionInfo | undefined>(undefined);
  const [draftId, setDraftId] = React.useState<number | undefined>(undefined);
  const [routeState, setRouteState] = React.useState<RouteState>(undefined);
  const [runtimeFlowUrl, setRuntimeFlowUrl] = React.useState<string | undefined>(undefined);
  const [permissionsReady, setPermissionsReady] = React.useState(false);
  const [securityError, setSecurityError] = React.useState(false);

  // ── All permissions default to true except viewAll ────────────────────────
  // viewAll controls AllRequestPage access — set dynamically from FinanceAdmin
  const [permissions, setPermissions] =
    React.useState<SecurityPermissions>({
      submit: true,
      viewOwn: true,
      viewTeam: true,
      viewAll: false,       // ← controlled by FinanceAdmin check
      approve: true,
      download: true,
      approveTeam: true,
      approveAll: false,    // ← controlled by FinanceAdmin check
      editMyRequest: true,
      exportOwnReport: true,
      exportAllReport: false,  // ← controlled by FinanceAdmin check
      exportTeamReport: true,
      viewExportOwnReport: true,
      viewExportTeamReport: true,
      viewExportAllReport: false,  // ← controlled by FinanceAdmin check
      emailNotifySubmit: true,
      emailNotifyAfterA1Approved: true,
      emailNotifyFinalApproval: true,
      emailOwnBulkUploadFinished: true,
    });

  // ── After permissions load, validate the hash-based initial route ──
  // If the user lacks access (e.g. #/all for a non-FinanceAdmin), fall back to home.
  const _hashValidated = React.useRef(false);
  React.useEffect(() => {
    if (!permissionsReady || _hashValidated.current || !_initialHashRoute) return;
    _hashValidated.current = true;
    if (!canAccessRoute(permissions, _initialHashRoute)) {
      setRoute("home");
      try { history.replaceState(null, "", "#"); } catch { /* ignore */ }
    }
  }, [permissionsReady, permissions]);

  const approvalApiKeyRef = React.useRef<string>("");
  const effectiveFlowUrl = runtimeFlowUrl || flowUrl;

  // ── Deep-link: validate record + enforce access control ──────────────────
  // Only the requestor, assigned approvers (canViewRequest=true in ApprovalChainJSON),
  // and Admin/Finance users (added to allApprovalLevels by buildApprovalChainFromLists)
  // are allowed to view the record. Everyone else sees an Access Denied screen.
  React.useEffect(() => {
    if (!_dlParams) return;
    const { requestId, transactionId } = _dlParams;

    const webUrl = context.pageContext.web.absoluteUrl;
    const myEmail = String(context.pageContext.user.email || "").toLowerCase().trim();

    // Remove the deep-link query params from the address bar.
    // Done here (in the effect) rather than in syncHash so the URL is clean
    // before any user interaction, and replaceState means back won't restore them.
    const _cleanDlUrl = (hash: string): void => {
      const myClean = window.location.pathname + hash;
      let topClean = myClean;
      try { topClean = (window.top?.location?.pathname ?? window.location.pathname) + hash; } catch { /* cross-origin */ }
      try { window.top?.history?.replaceState(null, "", topClean); } catch { /* cross-origin guard */ }
      try { history.replaceState(null, "", myClean); } catch { /* ignore */ }
    };

    // Fetch the record including ApprovalChainJSON and requester email for the access check
    const url =
      `${webUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${requestId})` +
      `?$select=Id,Title,ApprovalChainJSON,Requester/EMail` +
      `&$expand=Requester`;

    context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((res) => {
        if (!res.ok) throw new Error("Item not found");
        return res.json();
      })
      .then((item: {
        Id: number;
        Title: string;
        ApprovalChainJSON?: string;
        Requester?: { EMail?: string };
      }) => {
        // Validate TransactionId matches the record's Title (when provided)
        if (transactionId && item.Title !== transactionId) {
          setRoute("home");
          _cleanDlUrl("#");
          return;
        }

        // ── Access control ────────────────────────────────────────────────────
        // Rule 1: the requestor always has access
        const isRequester =
          String(item.Requester?.EMail || "").toLowerCase().trim() === myEmail;

        // Rule 2: any user listed in allApprovalLevels with canViewRequest=true
        // (covers: LOA approvers, Admin users, Finance users added in step 12b)
        let hasChainAccess = false;
        try {
          const chain = JSON.parse(String(item.ApprovalChainJSON || "{}")) as {
            allApprovalLevels?: Array<{ approverEmail?: string; canViewRequest?: boolean }>;
          };
          hasChainAccess = (chain.allApprovalLevels || []).some(
            (l) =>
              String(l.approverEmail || "").toLowerCase().trim() === myEmail &&
              l.canViewRequest === true
          );
        } catch { /* malformed JSON — deny */ }

        if (!isRequester && !hasChainAccess) {
          setDlAccessDenied(true);
          return;
        }

        // Authorized — open the transaction detail, then clean the URL so the
        // query param is gone before the user makes any further navigation.
        setSelectedSpId(item.Id);
        setExportSource("my");
        setRoute("exportPdf");
        _cleanDlUrl("#");
      })
      .catch(() => {
        setRoute("home");
        _cleanDlUrl("#");
      });
  }, [context]);

  React.useEffect(() => {
    // placeholder — handle runtimeFlowUrl changes if needed
  }, [runtimeFlowUrl]);

  // ── Check FinanceAdmin from EForm Users ───────────────────────────────────
  // Only controls viewAll — no other permission logic needed
  React.useEffect(() => {
    let cancelled = false;

    const loadSecurity = async (): Promise<void> => {
      try {
        const webUrl = context.pageContext.web.absoluteUrl;
        const userEmail = context.pageContext.user.email;
        const safeEmail = userEmail.replace(/'/g, "''");

        const userUrl =
          `${webUrl}/_api/web/lists/getbytitle('EForm%20Users')/items` +
          `?$top=1` +
          `&$select=FinanceAdmin` +
          `&$filter=Email eq '${safeEmail}'`;

        const res = await context.spHttpClient.get(
          userUrl,
          SPHttpClient.configurations.v1
        );

        if (!res.ok) {
          if (!cancelled) setPermissionsReady(true);
          return;
        }

        const json = await res.json() as {
          value: { FinanceAdmin?: string }[]
        };

        const user = json.value?.[0];
        const financeAdmin = String(user?.FinanceAdmin || "").trim().toUpperCase();
        const hasFullAccess = financeAdmin === "ADMIN" || financeAdmin === "FINANCE";

        if (!cancelled) {
          setPermissions((prev) => ({
            ...prev,
            viewAll: hasFullAccess,
            approveAll: hasFullAccess,
            exportAllReport: hasFullAccess,
            viewExportAllReport: hasFullAccess,
          }));
          setPermissionsReady(true);
        }
      } catch (err) {
        console.error("Security error:", err);
        if (!cancelled) setSecurityError(true);
        if (!cancelled) setPermissionsReady(true);
      }
    };

    loadSecurity().catch(() => undefined);
    getUserDataFromSharePoint(context).catch(() => undefined);

    return () => {
      cancelled = true;
    };
  }, [context]);

  const syncHash = (r: Route): void => {
    const h = (r === "my" || r === "team" || r === "all" || r === "create") ? `#/${r}` : "#";
    // Build pathname+hash so any existing query params (e.g. ?requestId=N from deep-links)
    // are stripped on every navigation, not just on initial deep-link load.
    const myClean = window.location.pathname + h;
    let topClean = myClean;
    try { topClean = (window.top?.location?.pathname ?? window.location.pathname) + h; } catch { /* cross-origin */ }
    try { window.top?.history?.replaceState(null, "", topClean); } catch { /* cross-origin guard */ }
    try { history.replaceState(null, "", myClean); } catch { /* ignore */ }
  };

  const safeNavigate = (r: Route, state?: RouteState): void => {
    if (!canAccessRoute(permissions, r)) {
      setRoute("home");
      setRouteState(undefined);
      syncHash("home");
      return;
    }
    setRouteState(state);
    setRoute(r);
    syncHash(r);
  };

  const getApprovalApiKey = React.useCallback(async (): Promise<string> => {
    if (approvalApiKeyRef.current) return approvalApiKeyRef.current;

    const webUrl = context.pageContext.web.absoluteUrl;
    const url =
      `${webUrl}/_api/web/lists/getbytitle('WorkflowConfiguration')/items` +
      `?$select=Title,ConfigKey,ConfigValue,IsActive` +
      `&$filter=IsActive eq 1`;

    const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) throw new Error(`Unable to load WorkflowConfiguration (${res.status})`);
    const json = (await res.json()) as { value?: Array<{ Title?: string; ConfigKey?: string; ConfigValue?: string }> };
    const rows = Array.isArray(json?.value) ? json.value : [];

    const candidates = [
      "APPROVAL_API_KEY",
      "APPROVE_REJECT_API_KEY",
      "SECUID",
      "SECURITY_ID",
      "X_API_KEY",
      "X-API-KEY",
      "API_KEY"
    ];

    const byKey = new Map<string, string>();
    rows.forEach((row) => {
      const k = String(row?.ConfigKey || row?.Title || "").trim().toUpperCase();
      const v = String(row?.ConfigValue || "").trim();
      if (k && v) byKey.set(k, v);
    });

    let apiKey = "";
    for (const key of candidates) {
      const value = byKey.get(key.toUpperCase());
      if (value) {
        apiKey = value;
        break;
      }
    }

    if (!apiKey) {
      apiKey = String(rows.find((x) => String(x?.ConfigValue || "").trim())?.ConfigValue || "").trim();
    }

    // eslint-disable-next-line require-atomic-updates
    approvalApiKeyRef.current = apiKey;
    return apiKey;
  }, [context]);

  const getCurrentUserPosition = React.useCallback(async (): Promise<string> => {
    const webUrl = context.pageContext.web.absoluteUrl;
    const userEmail = String(context.pageContext.user.email || "").replace(/'/g, "''");

    const url =
      `${webUrl}/_api/web/lists/getbytitle('EForm%20Users')/items` +
      `?$top=1` +
      `&$select=Jobposition` +
      `&$filter=Email eq '${userEmail}'`;

    const res = await context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!res.ok) return "Unknown";

    const json = await res.json() as { value?: { Jobposition?: string }[] };
    return String(json?.value?.[0]?.Jobposition || "Unknown");
  }, [context]);

  const submitExportDecision = React.useCallback(
    async (action: "Approve" | "Reject" | "Bypass", comments = ""): Promise<string> => {
      if (!selectedSpId) throw new Error("Request id is missing.");
      if (!ApprovalFlowUrl) throw new Error("Approval flow URL is missing.");
      if (exportSource === "team" && !permissions.approveTeam) {
        throw new Error("You do not have permission to approve/reject this team request.");
      }

      const apiKey = await getApprovalApiKey();
      if (!apiKey) throw new Error("Security key is missing in WorkflowConfiguration.");

      const actorEmail = String(context.pageContext.user.email || "").trim();
      const position = await getCurrentUserPosition();

      // ✅ Fetch request header to get transactionId, totals, purpose, status
      const webUrl = context.pageContext.web.absoluteUrl;
      const headerUrl =
        `${webUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${selectedSpId})` +
        `?$select=Title,TotalQuantity,TotalAmount,WorkflowStatus,Purpose/Title&$expand=Purpose`;

      const headerRes = await context.spHttpClient.get(headerUrl, SPHttpClient.configurations.v1);
      if (!headerRes.ok) {
        throw new Error(`Unable to load request header (${headerRes.status})`);
      }
      const header = await headerRes.json();

      // ✅ Use shared service — same logic as TeamRequestsPage
      const result = await submitApprovalDecision({
        context,
        approvalFlowUrl: ApprovalFlowUrl,
        apiKey,
        productFlowUrl: effectiveFlowUrl,
        requestId: selectedSpId,
        transactionId: String(header?.Title || ""),
        actorEmail,
        position,
        action,
        comments,
        totalQuantity: Number(header?.TotalQuantity || 0),
        totalTagPrice: Number(header?.TotalAmount || 0),
        purpose: String(header?.Purpose?.Title || ""),
        status: String(header?.WorkflowStatus || ""),
      });

      if (!result.success) throw new Error(result.message);

      return result.message; // ✅ return PA's message on success
    },
    [
      selectedSpId, ApprovalFlowUrl, exportSource, permissions.approveTeam,
      getApprovalApiKey, context, getCurrentUserPosition, effectiveFlowUrl
    ]
  );

  // ── REMOVED: buildExportActionHandlers ──
  // actionHandlers replaced by dedicated onBypass prop on ExportPDFPage

  // Detect Firefox — its ETP blocks SharePoint silent-auth iframes causing slow/failed loads
  const isFirefox = typeof navigator !== "undefined" && navigator.userAgent.toLowerCase().includes("firefox");

  const renderPage = (): JSX.Element => {
    // Deep-link access denied — user is not the requestor, an approver, or Admin/Finance
    if (dlAccessDenied) {
      return (
        <div className="flex flex-col items-center justify-center min-h-[400px] p-10 text-center gap-4">
          <svg xmlns="http://www.w3.org/2000/svg" className="h-14 w-14 text-gray-300" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
          </svg>
          <h2 className="text-lg font-semibold text-gray-700">Access Denied</h2>
          <p className="text-sm text-gray-500 max-w-xs leading-relaxed">
            You do not have permission to view this request. Only the requestor, assigned approvers, and Admin / Finance users can access this record.
          </p>
        </div>
      );
    }

    if (!permissionsReady) {
      return (
        <div className="p-6 text-[12px] text-gray-600">
          {isFirefox && (
            <div className="mb-4 p-3 rounded-lg bg-blue-50 border border-blue-200 text-[12px] text-blue-800">
              <strong>Firefox detected:</strong> If the page is slow or not loading, click the shield icon (&#128737;) in the address bar and turn off Enhanced Tracking Protection for this site, then refresh.
            </div>
          )}
          <Loading />
        </div>
      );
    }

    const securityErrorBanner = securityError ? (
      <div className="m-4 p-3 rounded-lg bg-yellow-50 border border-yellow-200 text-[12px] text-yellow-800">
        Unable to load your permissions. Some features may be unavailable. Please{" "}
        <button
          type="button"
          onClick={() => window.location.reload()}
          className="underline font-semibold"
        >
          refresh the page
        </button>
        .
      </div>
    ) : null;

    switch (route) {
      case "home":
        return (
          <>
            {securityErrorBanner}
            <HomePage
              permissions={permissions}
              onGoToForm={() => permissions.viewOwn && safeNavigate("my")}
              onCreateRequest={() => {
                if (permissions.submit) {
                  setDraftId(undefined);
                  safeNavigate("create");
                }
              }}
            />
          </>
        );

      case "create":
        return (
          <>
            {securityErrorBanner}
            <CreateRequestPage
              context={context}
              key={draftId ? `draft-${draftId}` : "new"}
              draftId={draftId}
              onSubmit={() => {
                setDraftId(undefined);
              }}
              onNavigate={safeNavigate}
              flowUrl={effectiveFlowUrl}
            />
          </>
        );

      case "my":
        return (
          <>
            {securityErrorBanner}
            <MyRequestsPage
              context={context}
              routeState={routeState}
              onRouteStateConsumed={() => setRouteState(undefined)}
              permissions={permissions}
              onOpenDraft={(id: number) => {
                setDraftId(id);
                safeNavigate("create");
              }}
              onOpenExportPdf={(spId: number) => {
                setSelectedSpId(spId);
                setExportSource("my");
                setExportActionInfo(undefined);
                safeNavigate("exportPdf");
              }}
              onDownloadPdf={(spId: number) => {
                setQuickPdfRequest({ spId, source: "my" });
              }}
              onCreate={() => {
                setDraftId(undefined);
                safeNavigate("create");
              }}
            />
          </>
        );

      case "team":
        return (
          <>
            {securityErrorBanner}
            <TeamRequestsPage
              context={context}
              permissions={permissions}
              routeState={routeState}
              onRouteStateConsumed={() => setRouteState(undefined)}
              onOpenExportPdf={(spId: number, actionInfo?: ExportActionInfo) => {
                setSelectedSpId(spId);
                setExportSource("team");
                setExportActionInfo(actionInfo);
                safeNavigate("exportPdf");
              }}
              onDownloadPdf={(spId: number) => {
                setQuickPdfRequest({ spId, source: "team" });
              }}
              ApprovalFlowUrl={ApprovalFlowUrl}
              productFlowUrl={effectiveFlowUrl}
            />
          </>
        );

      case "all":
        return (
          <>
            {securityErrorBanner}
            <AllRequestPage
              context={context}
              permissions={permissions}
              onOpenExportPdf={(spId: number) => {
                setSelectedSpId(spId);
                setExportSource("all");
                setExportActionInfo(undefined);
                safeNavigate("exportPdf");
              }}
              onDownloadPdf={(spId: number) => {
                setQuickPdfRequest({ spId, source: "all" });
              }}
            />
          </>
        );

      case "exportPdf":
        return (
          <>
            {securityErrorBanner}
            <ExportPDFPage
              context={context}
              spId={selectedSpId}
              permissions={permissions}
              exportSource={exportSource}
              exportActionInfo={exportSource === "team" ? exportActionInfo : undefined}
              onNavigate={safeNavigate}
              onApprove={async (comments?: string) =>
                submitExportDecision("Approve", comments || "")  // now returns Promise<string>
              }
              onReject={async (comments?: string) =>
                submitExportDecision("Reject", comments || "")
              }
              onBypass={async (comments?: string) =>
                submitExportDecision("Bypass", comments || "")
              }

            />
          </>
        );

      default:
        return (
          <>
            {securityErrorBanner}
            <HomePage
              permissions={permissions}
              onGoToForm={() => permissions.viewOwn && safeNavigate("my")}
              onCreateRequest={() => {
                if (permissions.submit) {
                  setDraftId(undefined);
                  safeNavigate("create");
                }
              }}
            />
          </>
        );
    }
  };

  return (
    <ErrorBoundary>
      <div className={`${styles.eform} pdf-thai`}>
        <AppShell
          activeRoute={route}
          exportSource={exportSource}
          onNavigate={(r: Route) => {
            if (r === "create") {
              setDraftId(undefined);
            }
            safeNavigate(r);
          }}
          onGoForm={() => safeNavigate("home")}
          context={context}
          roleReady={true}
          setRuntimeFlowUrl={setRuntimeFlowUrl}
          permissions={permissions}
        >
          {renderPage()}
          {quickPdfRequest && (
            <div style={{ position: "fixed", left: "-10000px", top: 0, width: "1200px" }}>
              <ExportPDFPage
                context={context}
                spId={quickPdfRequest.spId}
                permissions={permissions}
                exportSource={quickPdfRequest.source}
                onNavigate={safeNavigate}
                autoExport={true}
                onAutoExportDone={() => setQuickPdfRequest(undefined)}
              />
            </div>
          )}
        </AppShell>
      </div>
    </ErrorBoundary>
  );
};

export default Eform;