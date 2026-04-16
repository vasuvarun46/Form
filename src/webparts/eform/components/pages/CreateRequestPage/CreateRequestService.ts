// pages/CreateRequestPage.logic.ts
import * as XLSX from "xlsx";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, HttpClient } from "@microsoft/sp-http";
import { CostCenterOption, FlowItemResponse, IItemRow, ISuparaItem, ItemRowErrors, IUsageState, ODataResponse, PurposeOption, SPFileItem } from "./createRequestTypes";
import { logger } from "../../utils/logger";

/* =========================================================
   Flow URL
   ========================================================= */

/* =========================================================
   REQUIRED: Use the given function (NO any)
   ========================================================= */
export async function fetchProductFromAws(
  httpClient: HttpClient,
  productId: string,
  flowUrl: string
): Promise<unknown> {
  const response = await httpClient.post(flowUrl, HttpClient.configurations.v1, {
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sku: productId }),
  });

  if (!response.ok) {
    throw new Error(`Flow failed: ${response.statusText}`);
  }

  return await response.json();
}

/* =========================================================
   Types (NO any, NO null)
   ========================================================= */

/* =========================================================
BUILD APPROVAL CHAIN (Proper Staging from ApprovalMatrix)
========================================================= */

export async function buildApprovalChainFromLists(
  context: WebPartContext,
  totalAmount: number
): Promise<{
  approvalLabel: string;
  firstStageApproverId: number | undefined;
  firstStageApproverEmails: string[];
  stageCount: number;
  hasMultipleApprovals: boolean;
  chainJSON: string;
}> {

  const siteUrl: string = context.pageContext.web.absoluteUrl;
  const currentUserEmail: string = String(context.pageContext.user.email || "").trim();
  const escapedEmail: string = currentUserEmail.replace(/'/g, "''");

  const parseAmount = (value: unknown): number => {
    const raw = String(value ?? "").replace(/,/g, "").trim();
    const n = Number(raw);
    return Number.isFinite(n) ? n : 0;
  };

  const fallback = (label: string) => ({
    approvalLabel: label,
    firstStageApproverId: undefined,
    firstStageApproverEmails: [] as string[],
    stageCount: 0,
    hasMultipleApprovals: false,
    chainJSON: JSON.stringify({
      approvalType: label,
      totalAmount,
      stageCount: 0,
      hasMultipleApprovals: false,
      stages: [],
      allApprovalLevels: []
    })
  });

  const resolveUserId = async (email: string): Promise<number | undefined> => {
    try {
      const url = `${siteUrl}/_api/web/ensureUser('${email.replace(/'/g, "''")}')`;
      const res = await context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: { "Content-Type": "application/json;odata=verbose" },
          body: JSON.stringify({})
        }
      );
      if (!res.ok) return undefined;
      const json = await res.json() as { Id?: number };
      return json.Id;
    } catch {
      return undefined;
    }
  };

  // ── 1. FETCH LOA Details fields ──────────────────────────────────────────
  const loaFieldsUrl =
    `${siteUrl}/_api/web/lists/getbytitle('LOA%20Details')/fields` +
    `?$select=InternalName,Title,FieldTypeKind` +
    `&$filter=Hidden eq false and ReadOnlyField eq false`;

  const loaFieldsRes = await context.spHttpClient.get(loaFieldsUrl, SPHttpClient.configurations.v1);
  if (!loaFieldsRes.ok) return fallback("LOA_FIELDS_LOOKUP_FAILED");

  const loaFieldsJson = await loaFieldsRes.json() as {
    value: { InternalName: string; Title: string; FieldTypeKind: number }[]
  };

  const systemFields = new Set([
    "ContentType", "Attachments", "Edit", "LinkTitleNoMenu", "LinkTitle",
    "DocIcon", "ItemChildCount", "FolderChildCount", "_UIVersionString",
    "AppAuthor", "AppEditor"
  ]);

  const approvalFlagFields = loaFieldsJson.value.filter(
    (f) => f.FieldTypeKind === 8 && !systemFields.has(f.InternalName)
  );

  if (!approvalFlagFields.length) return fallback("NO_LOA_FLAG_FIELDS");

  logger.log("[ApprovalChain] LOA flag fields:", approvalFlagFields.map(f => `${f.Title} (${f.InternalName})`));

  // ── 2. FETCH EForm Users fields ──────────────────────────────────────────
  const userFieldsUrl =
    `${siteUrl}/_api/web/lists/getbytitle('EForm%20Users')/fields` +
    `?$select=InternalName,Title,FieldTypeKind` +
    `&$filter=Hidden eq false and ReadOnlyField eq false`;

  const userFieldsRes = await context.spHttpClient.get(userFieldsUrl, SPHttpClient.configurations.v1);
  if (!userFieldsRes.ok) return fallback("USER_FIELDS_LOOKUP_FAILED");

  const userFieldsJson = await userFieldsRes.json() as {
    value: { InternalName: string; Title: string; FieldTypeKind: number }[]
  };

  const emailFields = userFieldsJson.value.filter((f) => {
    const titleUpper = f.Title.trim().toUpperCase();
    const internalUpper = f.InternalName.toUpperCase();
    if (f.FieldTypeKind !== 2 && f.FieldTypeKind !== 9) return false;
    if (!titleUpper.includes("EMAIL") && !internalUpper.includes("EMAIL")) return false;
    if (internalUpper === "EMAIL") return false;
    return true;
  });

  logger.log("[ApprovalChain] Matched email fields:", emailFields.map(f => `Title="${f.Title}" Internal="${f.InternalName}"`));

  // ── 3. BUILD ROLE KEY HELPERS ────────────────────────────────────────────
  const normalize = (s: string): string =>
    s.trim().toUpperCase().replace(/[\s_]+/g, "");

  const roleKeys = (raw: string): string[] => {
    const base = normalize(raw);
    const keys = [base];
    const noLevel = base.replace(/LEVEL$/, "").replace(/LVL$/, "");
    if (noLevel && noLevel !== base) keys.push(noLevel);
    if (noLevel && !base.endsWith("LEVEL")) keys.push(noLevel + "LEVEL");
    const unique: string[] = [];
    for (const k of keys) {
      if (unique.indexOf(k) === -1) unique.push(k);
    }
    return unique;
  };

  const emailFieldMap = new Map<string, string>();
  for (const ef of emailFields) {
    const baseName = ef.Title.replace(/\s*email\s*/i, "").trim();
    if (baseName) {
      for (const key of roleKeys(baseName)) {
        if (!emailFieldMap.has(key)) emailFieldMap.set(key, ef.InternalName);
      }
    }
    const internalBase = ef.InternalName.replace(/email$/i, "").replace(/Email$/, "").trim();
    if (internalBase) {
      for (const key of roleKeys(internalBase)) {
        if (!emailFieldMap.has(key)) emailFieldMap.set(key, ef.InternalName);
      }
    }
  }

  logger.log("[ApprovalChain] Email field map:", Array.from(emailFieldMap.entries()));

  // ── 4. FETCH REQUESTER row from EForm Users ──────────────────────────────
  const emailFieldNames = emailFields.map((f) => f.InternalName).join(",");

  const userUrl =
    `${siteUrl}/_api/web/lists/getbytitle('EForm%20Users')/items` +
    `?$top=1` +
    `&$select=Id,Title,Email,Jobposition,Department/Title,${emailFieldNames}` +
    `&$expand=Department` +
    `&$filter=Email eq '${escapedEmail}'`;

  const userRes = await context.spHttpClient.get(userUrl, SPHttpClient.configurations.v1);
  if (!userRes.ok) return fallback("USER_LOOKUP_FAILED");

  const userJson = await userRes.json() as { value: Record<string, unknown>[] };
  const requester = userJson.value?.[0];
  if (!requester) return fallback("NO_REQUESTER");

  // ── 5. FETCH LOA Details rows ────────────────────────────────────────────
  const flagInternalNames = approvalFlagFields.map((f) => f.InternalName).join(",");

  const loaUrl =
    `${siteUrl}/_api/web/lists/getbytitle('LOA%20Details')/items` +
    `?$select=Title,Min,Max,${flagInternalNames}` +
    `&$orderby=Min asc`;

  const loaRes = await context.spHttpClient.get(loaUrl, SPHttpClient.configurations.v1);
  if (!loaRes.ok) return fallback("LOA_LOOKUP_FAILED");

  const loaJson = await loaRes.json() as { value: Record<string, unknown>[] };
  const loaRows = loaJson.value || [];

  // ── 6. MATCH LOA ROW by amount ───────────────────────────────────────────
  // ✅ FIX 1 & 2: dot notation for Min and Max
  const matchedLOA = loaRows.find((row) => {
    const min = parseAmount(row.Min);
    const max = parseAmount(row.Max);
    return totalAmount >= min && totalAmount <= max;
  });

  if (!matchedLOA) return fallback("NO_LOA_MATCH");

  // ── 7. COUNT ENABLED CHECKBOXES ─────────────────────────────────────────
  const enabledFlags = approvalFlagFields.filter(
    (f) => matchedLOA[f.InternalName] === true
  );

  if (!enabledFlags.length) return fallback("NO_APPROVER_FOUND");

  // ── 8. DETERMINE APPROVAL TYPE ──────────────────────────────────────────
  const isSingleApproval = enabledFlags.length === 1;
  const finalApprovalType = isSingleApproval ? "SingleApprover" : "MultipleApprover";

  logger.log(`[ApprovalChain] Approval type: ${finalApprovalType} (${enabledFlags.length} flag(s) enabled)`);

  // ── 9. BUILD ORDERED LEVELS (LOW → HIGH authority) ───────────────────────
  const approvalFlagFieldsLowToHigh = approvalFlagFields.slice().reverse();

  type RawLevel = {
    roleName: string;
    emailColInternalName: string;
    approverEmail: string;
  };

  const rawLevels: RawLevel[] = [];
  const seenEmails = new Set<string>();

  for (const flagField of approvalFlagFieldsLowToHigh) {
    if (matchedLOA[flagField.InternalName] !== true) continue;

    const candidates = roleKeys(flagField.Title);
    let emailColInternalName: string | undefined;
    for (const ck of candidates) {
      emailColInternalName = emailFieldMap.get(ck);
      if (emailColInternalName) break;
    }

    logger.log(
      `[ApprovalChain] Flag "${flagField.Title}": keys=${JSON.stringify(candidates)}, ` +
      `emailCol="${emailColInternalName || 'NOT FOUND'}"`
    );

    if (!emailColInternalName) {
      logger.log(`[ApprovalChain] Skipping "${flagField.Title}" — no email column found`);
      continue;
    }

    const approverEmail = String(requester[emailColInternalName] || "").trim();

    if (isSingleApproval) {
      const emailKey = approverEmail.toUpperCase();
      if (approverEmail && seenEmails.has(emailKey)) {
        logger.log(`[ApprovalChain] Skipping "${flagField.Title}" — duplicate email "${approverEmail}"`);
        continue;
      }
      if (approverEmail) seenEmails.add(emailKey);
    }

    rawLevels.push({
      roleName: flagField.Title,
      emailColInternalName,
      approverEmail
    });
  }

  logger.log(
    "[ApprovalChain] rawLevels (LOW→HIGH, enabled only):",
    rawLevels.map(r => `${r.roleName}: "${r.approverEmail || 'EMPTY'}"`)
  );

  // ── 10. CLIMB UP if lowest level has no email ────────────────────────────
  const levelsWithEmail = rawLevels.filter(r => r.approverEmail);

  if (!levelsWithEmail.length) return fallback("NO_APPROVER_EMAIL_FOUND");

  // ── 11. BUILD STAGES ─────────────────────────────────────────────────────
  type Stage = {
    stage: number;
    approvalRole: string;
    approverEmail: string;
    approverId: number | undefined;
    isApprove: boolean;
  };

  const stages: Stage[] = [];

  if (isSingleApproval) {
    const level = levelsWithEmail[0];
    const approverId = await resolveUserId(level.approverEmail);
    stages.push({
      stage: 1,
      approvalRole: level.roleName,
      approverEmail: level.approverEmail,
      approverId,
      isApprove: true
    });
  } else {
    for (const level of rawLevels) {
      if (!level.approverEmail) {
        logger.log(`[ApprovalChain] Skipping stage for "${level.roleName}" — empty email`);
        continue;
      }
      const approverId = await resolveUserId(level.approverEmail);
      stages.push({
        stage: stages.length + 1,
        approvalRole: level.roleName,
        approverEmail: level.approverEmail,
        approverId,
        isApprove: stages.length === 0
      });
    }
  }

  if (!stages.length) return fallback("NO_APPROVER_FOUND");

  // ── 12. BUILD allApprovalLevels ──────────────────────────────────────────
  //
  // ✅ UPDATED canApproveRequest logic for SingleApprover:
  //    - Find the ORDER (index) of the matched/enabled role in approvalFlagFieldsLowToHigh
  //    - Roles AT or ABOVE that order → canApproveRequest = true
  //    - Roles BELOW that order → canApproveRequest = false
  //
  // For MultipleApprover: canApproveRequest = true only if their email is in stages
  // (unchanged from before)

  type ApprovalLevel = {
    order: number;
    approvalRole: string;
    approverEmail: string;
    approverId: number | undefined;
    canApproveRequest: boolean;
    canViewRequest: boolean;
    isCurrentApprover: boolean;
    isEnabledByLOA: boolean;
    emailMissing: boolean;
  };

  const stageEmails = new Set(stages.map(s => s.approverEmail.toUpperCase()));

  // ✅ For SingleApprover: find the order index of the enabled flag
  // approvalFlagFieldsLowToHigh is LOW → HIGH, so index 0 = lowest authority
  // The enabled flag's index is the threshold: at or above = canApprove
  const singleEnabledFlagInternalName = isSingleApproval
    ? approvalFlagFields.find((f) => matchedLOA[f.InternalName] === true)?.InternalName
    : undefined;

  // In approvalFlagFieldsLowToHigh (reversed), find the threshold index
  const singleEnabledOrderIndex = isSingleApproval && singleEnabledFlagInternalName
    ? approvalFlagFieldsLowToHigh.findIndex((f) => f.InternalName === singleEnabledFlagInternalName)
    : -1;

  logger.log(
    `[ApprovalChain] Single enabled flag: "${singleEnabledFlagInternalName || 'N/A'}", ` +
    `threshold order index in LowToHigh: ${singleEnabledOrderIndex}`
  );

  const allApprovalLevels: ApprovalLevel[] = [];

  for (let i = 0; i < approvalFlagFieldsLowToHigh.length; i++) {
    const flagField = approvalFlagFieldsLowToHigh[i];
    const isEnabledByLOA = matchedLOA[flagField.InternalName] === true;

    const candidates = roleKeys(flagField.Title);
    let emailColInternalName: string | undefined;
    for (const ck of candidates) {
      emailColInternalName = emailFieldMap.get(ck);
      if (emailColInternalName) break;
    }

    const approverEmail = emailColInternalName
      ? String(requester[emailColInternalName] || "").trim()
      : "";

    const approverId = approverEmail
      ? await resolveUserId(approverEmail)
      : undefined;

    const emailMissing = !approverEmail;

    // ✅ canApproveRequest logic:
    let canApproveRequest: boolean;

    if (isSingleApproval) {
      // At or above the threshold index in LowToHigh array = higher/equal authority
      // i >= singleEnabledOrderIndex means this role is at or above the enabled level
      canApproveRequest =
        singleEnabledOrderIndex >= 0 &&
        i >= singleEnabledOrderIndex &&
        !emailMissing;
    } else {
      // MultipleApprover: only if their email is explicitly in stages
      canApproveRequest = isEnabledByLOA && stageEmails.has(approverEmail.toUpperCase());
    }

    const isCurrentApprover =
      stages[0]?.approverEmail?.toUpperCase() === approverEmail.toUpperCase();

    allApprovalLevels.push({
      order: i + 1,
      approvalRole: flagField.Title,
      approverEmail,
      approverId,
      canApproveRequest,
      canViewRequest: true,
      isCurrentApprover,
      isEnabledByLOA,
      emailMissing
    });
  }

  // ── 13. FINALIZE AND RETURN ──────────────────────────────────────────────
  const firstStage = stages[0];
  const pendingStages = stages.slice(1);
  const hasMultipleApprovals = stages.length > 1;

  logger.log(
    `[ApprovalChain] FINAL: type=${finalApprovalType}, stages=${stages.length}, ` +
    `firstApprover="${firstStage.approverEmail}"`
  );

  return {
    approvalLabel: finalApprovalType,
    firstStageApproverId: firstStage.approverId,
    firstStageApproverEmails: [firstStage.approverEmail],
    stageCount: stages.length,
    hasMultipleApprovals,
    // ✅ FIX 3, 4, 5: dot notation for Title, Jobposition, Department
    chainJSON: JSON.stringify({
      approvalType: finalApprovalType,
      totalAmount,
      stageCount: stages.length,
      hasMultipleApprovals,
      requesterEmail: currentUserEmail,
      requesterJobposition: String(requester.Jobposition || ""),
      requesterDepartment: (requester.Department as { Title?: string })?.Title || "",
      stages,
      pendingStages,
      allApprovalLevels
    })
  };
}

export const getItemsArray = <TItem,>(data: ODataResponse<TItem>): TItem[] => {
  if ("value" in data && Array.isArray(data.value)) return data.value;
  if ("d" in data && Array.isArray(data.d.results)) return data.d.results;
  return [];
};

export const spGet = async <T,>(context: WebPartContext, apiUrl: string): Promise<T> => {
  const res: SPHttpClientResponse = await context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(txt || res.statusText);
  }
  return (await res.json()) as T;
};

export const spCreateItem = async <TBody extends object>(
  context: WebPartContext,
  listTitle: string,
  body: TBody
): Promise<{ Id: number }> => {
  const siteUrl: string = context.pageContext.web.absoluteUrl;
  const url: string = `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;

  const res: SPHttpClientResponse = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "odata-version": "",
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(txt || res.statusText);
  }
  return (await res.json()) as { Id: number };
};

export const spUpdateItem = async <TBody extends object>(
  context: WebPartContext,
  listTitle: string,
  itemId: number,
  body: TBody
): Promise<void> => {
  const siteUrl: string = context.pageContext.web.absoluteUrl;
  const url: string = `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`;

  const res: SPHttpClientResponse = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: "application/json;odata=nometadata",
      "Content-type": "application/json;odata=nometadata",
      "odata-version": "",
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE",
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(txt || res.statusText);
  }
};

export const normalizeFlowToRow = (
  json: unknown
): FlowItemResponse | undefined => {
  if (!json) return undefined;

  let obj: unknown = json;

  // Case 1: Flow returned array
  if (Array.isArray(obj)) {
    obj = obj[0];
  }

  // Case 2: Flow wrapped inside { body: {...} }
  if (
    obj &&
    typeof obj === "object" &&
    "body" in obj
  ) {
    const wrapper = obj as { body?: unknown };

    if (wrapper.body && typeof wrapper.body === "object") {
      obj = wrapper.body;
    }
  }

  if (!obj || typeof obj !== "object") return undefined;

  const candidate = obj as Partial<FlowItemResponse>;

  if (!candidate.Barcode || String(candidate.Barcode).trim() === "") {
    return undefined;
  }

  return candidate as FlowItemResponse;
};




/** yyyy-MM-dd -> ISO */
export const yyyyMmDdToIso = (yyyyMmDd: string): string => {
  const [yyyy, mm, dd] = (yyyyMmDd || "").trim().split("-");
  if (!yyyy || !mm || !dd) throw new Error("Invalid date. Use yyyy-MM-dd");

  // Bangkok (UTC+7) midnight = subtract 7 hours from UTC
  // e.g. 2025-03-24 00:00 Bangkok = 2025-03-23T17:00:00Z
  const date = new Date(`${yyyy}-${mm}-${dd}T00:00:00+07:00`);
  return date.toISOString();
};

export const getCurrentUserId = async (context: WebPartContext): Promise<number> => {
  // 1) Try pageContext (no REST call)
  const legacy = context.pageContext.legacyPageContext as { userId?: number } | undefined;
  const ctxId = Number(legacy?.userId || 0);
  if (ctxId > 0) return ctxId;

  // 2) Fallback: resolve by email via siteusers
  const email = String(context.pageContext.user?.email || "").trim();
  if (!email) throw new Error("Cannot resolve user id: missing email in pageContext.");

  const siteUrl = context.pageContext.web.absoluteUrl;
  const url =
    `${siteUrl}/_api/web/siteusers?$select=Id,Email&$filter=` +
    encodeURIComponent(`Email eq '${email.replace(/'/g, "''")}'`);

  const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
    headers: { Accept: "application/json;odata=nometadata", "odata-version": "" },
  });

  if (!res.ok) throw new Error((await res.text()) || res.statusText);

  const json = (await res.json()) as { value: Array<{ Id: number; Email: string }> };
  const id = Number(json.value?.[0]?.Id || 0);
  if (!id) throw new Error("Cannot resolve user id from siteusers.");
  return id;
};

export const getUserDataFromSharePoint = async (
  context: WebPartContext
): Promise<ISuparaItem[]> => {
  try {
    const email = context.pageContext.user.email;
    const siteUrl = context.pageContext.web.absoluteUrl;
    const safeEmail = email.replace(/'/g, "''");

    // ── 1. FETCH EForm Users fields dynamically ───────────────────────────

    const fieldsUrl =
      `${siteUrl}/_api/web/lists/getbytitle('EForm%20Users')/fields` +
      `?$select=InternalName,Title,FieldTypeKind` +
      `&$filter=Hidden eq false and ReadOnlyField eq false`;

    const fieldsRes = await context.spHttpClient.get(
      fieldsUrl,
      SPHttpClient.configurations.v1
    );
    if (!fieldsRes.ok) return [];

    const fieldsJson = await fieldsRes.json() as {
      value: { InternalName: string; Title: string; FieldTypeKind: number }[]
    };

    const systemFields = new Set([
      "ContentType", "Attachments", "Edit", "LinkTitleNoMenu", "LinkTitle",
      "DocIcon", "ItemChildCount", "FolderChildCount", "_UIVersionString",
      "AppAuthor", "AppEditor"
    ]);

    // FieldTypeKind 7 = Lookup column
    const regularFields: string[] = ["Id", "Title", "Email", "Jobposition"];
    const lookupFields: string[] = [];
    const lookupSelectParts: string[] = [];

    for (let i = 0; i < fieldsJson.value.length; i++) {
      const f = fieldsJson.value[i];

      if (systemFields.has(f.InternalName)) continue;

      if (f.FieldTypeKind === 7) {
        // Lookup: needs /Id and /Title in $select, and field name in $expand
        lookupFields.push(f.InternalName);
        lookupSelectParts.push(f.InternalName + "/Id");
        lookupSelectParts.push(f.InternalName + "/Title");
      } else {
        // Only add if not already in core fields
        if (regularFields.indexOf(f.InternalName) === -1) {
          regularFields.push(f.InternalName);
        }
      }
    }

    // Build $select and $expand strings
    const finalSelect = regularFields.concat(lookupSelectParts).join(",");
    const finalExpand = lookupFields.join(",");

    // ── 2. FETCH REQUESTER row from EForm Users ───────────────────────────

    const apiUrl =
      `${siteUrl}/_api/web/lists/getbytitle('EForm%20Users')/items` +
      `?$top=1` +
      `&$select=${finalSelect}` +
      (finalExpand ? `&$expand=${finalExpand}` : "") +
      `&$filter=Email eq '${safeEmail}'`;

    const res = await context.spHttpClient.get(
      apiUrl,
      SPHttpClient.configurations.v1
    );
    if (!res.ok) return [];

    const data = await res.json();
    return (data.value ?? []) as ISuparaItem[];

  } catch (e) {
    console.error("Error fetching data from EForm Users:", e);
    return [];
  }
};



export const loadMasters = async (
  context: WebPartContext,
  setPurposes: (v: PurposeOption[]) => void,
  setCostCenters: (v: CostCenterOption[]) => void
): Promise<void> => {
  const siteUrl: string = context.pageContext.web.absoluteUrl;
  const costCenterInternalName: string = "CostCenterCode";

  // 1) Get login user's department id from RequesterDirectory
  const userRows = await getUserDataFromSharePoint(context);
  // const userRole = String(userRows?.[0]?.JobPosition || "").trim().toUpperCase();
  //const isGlobalRoleUser = userRole === "CEO" || userRole === "CFO";
  const deptId =
    userRows?.[0]?.Department?.Id ??

    null;

  const purposeUrl: string =
    `${siteUrl}/_api/web/lists/getbytitle('PurposeMaster')/items` +
    `?$top=5000` +
    `&$select=Id,Title,IsActive` +
    `&$filter=IsActive eq 1` +
    `&$orderby=Title`;

  // 2) Filter CostCenterMaster by DepartmentId (lookup filter) + IsActive
  //    Also expand Department to map it into CostCenterOption
  // const deptFilter = isGlobalRoleUser
  //   ? ` and Department/Title eq 'GLOBAL'`
  //   : (deptId ? ` and Department/Id eq ${deptId}` : "");
  logger.log(deptId);
  const costUrl: string =
    `${siteUrl}/_api/web/lists/getbytitle('CostCenterMaster')/items` +
    `?$top=5000` +
    `&$select=Id,Title,${costCenterInternalName},IsActive,Department/Id,Department/Title` +
    `&$expand=Department` +
    `&$filter=IsActive eq 1 and Department/Id eq ${deptId}` +
    `&$orderby=Title`;

  const [pJson, cJson] = await Promise.all([
    spGet<ODataResponse<{ Id: number; Title: string; IsActive: boolean }>>(context, purposeUrl),
    spGet<
      ODataResponse<
        { Id: number; Title: string; IsActive: boolean; Department?: { Id: number; Title: string } } &
        Record<string, unknown>
      >
    >(context, costUrl),
  ]);

  logger.log("deptId:", deptId);
  logger.log(costUrl);

  const pArr = getItemsArray<{ Id: number; Title: string; IsActive: boolean }>(pJson);
  const cArr = getItemsArray<
    { Id: number; Title: string; IsActive: boolean; Department?: { Id: number; Title: string } } &
    Record<string, unknown>
  >(cJson);

  const pList: PurposeOption[] = pArr.map((x) => ({ Id: x.Id, Title: x.Title }));

  const cList: CostCenterOption[] = cArr.map((x) => ({
    Id: x.Id,
    Title: x.Title,
    CostCenterCode: String(x[costCenterInternalName] ?? ""),
    // Department added (will compile only if CostCenterOption includes it)
    Department: x.Department ? { Id: x.Department.Id, Title: x.Department.Title } : undefined,
  }));
  logger.log(cList);
  setPurposes(pList);
  setCostCenters(cList);
};


export const sendMailGraph = async (
  context: WebPartContext,
  toEmails: string[],
  subject: string,
  htmlBody: string
): Promise<void> => {
  if (!toEmails.length) return;

  const client = await context.msGraphClientFactory.getClient("3");
  await client.api("/me/sendMail").post({
    message: {
      subject,
      body: { contentType: "HTML", content: htmlBody },
      toRecipients: toEmails.map((e) => ({ emailAddress: { address: e } })),
    },
    saveToSentItems: true,
  });
};

// export const makeTxnId = (id: number): string => {
//   const now = new Date();

//   const year = now.getFullYear();
//   const month = String(now.getMonth() + 1).padStart(2, "0");
//   const paddedId = String(id).padStart(4, "0");

//   return `II${year}${month}/${paddedId}`;
// };



/* =========================================================
   Internal helpers
========================================================= */

const getTxnDateParts = (
  asOf: Date
): { year: number; month: string; ym: string } => {
  const year = asOf.getFullYear();
  const month = String(asOf.getMonth() + 1).padStart(2, "0");
  return { year, month, ym: `${year}${month}` };
};

const to4Digits = (value: number): string => {
  const safe = Number.isFinite(value)
    ? Math.max(0, Math.trunc(value))
    : 0;
  return String(safe).padStart(4, "0").slice(-4);
};

/**
 * Parses the sequence number from a Title like:
 *   II202603/0001  → 1
 *   II202603/0042  → 42
 * Returns 0 if not matched.
 */
const parseMonthlySequenceFromTitle = (
  title: string,
  ym: string
): number => {
  const tx = String(title || "").trim();
  if (!tx) return 0;

  // Format: II202603/0001
  // ✅ Static regex — capture YYYYMM and sequence separately to avoid dynamic RegExp construction
  const match = /^II(\d{6})\/(\d{4})$/i.exec(tx);
  if (!match) return 0;

  // ✅ Validate the captured YM group matches the expected value
  if (match[1] !== ym) return 0;

  return Number(match[2]);
};
/* =========================================================
   OData helpers
========================================================= */



/* =========================================================
   getNextSequence
   Derives next sequence from the last known txnId string.
   Used as a quick in-memory fallback (no SP call needed).
========================================================= */

export const getNextSequence = (lastTxnId?: string): number => {
  const now = new Date();
  const { ym } = getTxnDateParts(now);

  if (!lastTxnId) return 1;

  const [prefix, seqRaw = ""] = String(lastTxnId).split("/");
  const lastYM = String(prefix || "").replace(/^II/i, "");

  // Different month → reset to 1
  if (lastYM !== ym) return 1;

  const seqDigits = String(seqRaw).replace(/\D/g, "");
  const seq4 = seqDigits.slice(-4);
  const parsed = Number(seq4);
  return Number.isFinite(parsed) && parsed > 0 ? parsed + 1 : 1;
};

/* =========================================================
   getNextMonthlyTxnSequence
   Queries SharePoint for the highest sequence number used
   this month and returns the next available one.
   Returns 1 if no matching titles found this month.
========================================================= */

export const getNextMonthlyTxnSequence = async (
  context: WebPartContext,
  asOf: Date = new Date()
): Promise<number> => {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const { ym } = getTxnDateParts(asOf);

  // Filter to only this month's items for efficiency
  const url =
    `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items` +
    `?$top=5000` +
    `&$select=Id,Title` +
    `&$filter=startswith(Title,'II${ym}/')` +
    `&$orderby=Id desc`;

  const json = await spGet<ODataResponse<{ Id: number; Title?: string }>>(
    context,
    url
  );
  const rows = getItemsArray<{ Id: number; Title?: string }>(json);

  // Find highest sequence this month
  let maxSeq = 0;
  for (let i = 0; i < rows.length; i++) {
    const seq = parseMonthlySequenceFromTitle(
      String(rows[i].Title || ""),
      ym
    );
    if (seq > maxSeq) maxSeq = seq;
  }

  return maxSeq + 1;
};

/* =========================================================
   makeTxnId
   submitted   → II202603/0001
   draft       → 0001  (no prefix, temp until submitted)
   ExcelUpload → 0001  (no prefix, assigned on submit)
========================================================= */

export const makeTxnId = (
  runningNo: number,
  mode: "submitted" | "draft" | "ExcelUpload" = "submitted",
  asOf: Date = new Date()
): string => {
  const { ym } = getTxnDateParts(asOf);
  const seq4 = to4Digits(runningNo);

  // ✅ Fixed: was (mode === "draft" || "ExcelUpload") — always true
  if (mode === "draft" || mode === "ExcelUpload") {
    return seq4;
  }

  return `II${ym}/${seq4}`;
};

/* =========================================================
   getOrCreateTxnId
   - If the item already has a valid Title like II202603/0001
     → return it as-is (no SP call needed)
   - Otherwise generate a new sequence, build the ID,
     and update the SharePoint item Title
========================================================= */

export const getOrCreateTxnId = async (
  context: WebPartContext,
  requestId: number,
  currentTitle: string,
  asOf: Date = new Date()
): Promise<string> => {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const { ym } = getTxnDateParts(asOf);

  // Already has a valid ID for this month — reuse it
  const existing = parseMonthlySequenceFromTitle(currentTitle, ym);
  if (existing > 0) {
    return currentTitle.trim();
  }

  // No valid ID — generate next sequence and build txnId
  const runningNo = await getNextMonthlyTxnSequence(context, asOf);
  const txnId = makeTxnId(runningNo, "submitted", asOf);

  // Update the SharePoint item Title
  const updateRes = await context.spHttpClient.post(
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
      body: JSON.stringify({ Title: txnId }),
    }
  );

  if (!updateRes.ok) {
    const txt = await updateRes.text();
    throw new Error(txt || "Failed to update Title with txnId");
  }

  return txnId;
};

/* =========================================================
   Excel read + validate (MAX 300 rows)
   ========================================================= */
export const normalizeHeader = (h: unknown): string =>
  String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/_/g, "");

export const readExcelToItemsMax300 = async (file: File): Promise<IItemRow[]> => {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error("Excel has no sheets.");

  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "" });
  // If no rows, return empty array (don't enforce content validation here).
  if (!rows || !rows.length) return [];

  // Map rows into IItemRow without performing content validations here.
  // We try to pick common header names but fall back to empty values.
  const keys = Object.keys(rows[0] || {});
  const map: Record<string, string> = {};
  keys.forEach((k) => (map[normalizeHeader(k)] = k));

  const get = (r: Record<string, unknown>, name: string): string => {
    const realKey = map[normalizeHeader(name)];
    return realKey ? String(r[realKey] ?? "").trim() : "";
  };

  return rows.map((r, idx) => {
    const sku = get(r, "sku");
    const qty = Number(get(r, "qty"));

    return {
      no: idx + 1,
      sku,
      productCode: "",
      color: "-",
      size: "-",
      unitPrice: 0,
      qty: isNaN(qty) ? 0 : qty,
    };
  });
};

export const callFlowBulk = async (
  context: WebPartContext,
  skus: string[],
  flowUrl: string
): Promise<FlowItemResponse[]> => {
  const out: FlowItemResponse[] = [];

  for (let i = 0; i < skus.length; i++) {
    const sku = String(skus[i] || "").trim();
    if (!sku) continue;

    const json = await fetchProductFromAws(context.httpClient, sku, flowUrl);
    const row = normalizeFlowToRow(json);
    if (row) out.push(row);
    else out.push({ Barcode: sku });
  }

  return out;
};

export async function downloadTemplateFromLibrary(
  context: WebPartContext,
  libraryTitle: string,
  options?: { extensionFilter?: string } // e.g. ".xlsx"
): Promise<void> {
  const siteUrl = context.pageContext.web.absoluteUrl;
  const safeTitle = libraryTitle.replace(/'/g, "''");

  // 1) Get library root folder server-relative URL
  const rootUrlApi =
    `${siteUrl}/_api/web/lists/getbytitle('${safeTitle}')/RootFolder` +
    `?$select=ServerRelativeUrl`;

  const rootRes: SPHttpClientResponse = await context.spHttpClient.get(
    rootUrlApi,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata", "odata-version": "" } }
  );

  if (!rootRes.ok) {
    const txt = await rootRes.text();

    // Helpful auth message in Incognito/InPrivate
    if (rootRes.status === 401 || rootRes.status === 403) {
      throw new Error("Not signed in or no access to the template library. Please sign in and try again.");
    }

    throw new Error(txt || rootRes.statusText);
  }

  const rootJson = (await rootRes.json()) as { ServerRelativeUrl?: string };
  const rootFolderRel = String(rootJson.ServerRelativeUrl || "").trim();
  if (!rootFolderRel) {
    throw new Error(`Unable to resolve RootFolder for library "${libraryTitle}".`);
  }

  // 2) List files in root folder (latest modified first)
  const filesApi =
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${rootFolderRel.replace(/'/g, "''")}')/Files` +
    `?$select=Name,ServerRelativeUrl,TimeLastModified` +
    `&$orderby=TimeLastModified desc` +
    `&$top=200`;

  const filesRes: SPHttpClientResponse = await context.spHttpClient.get(
    filesApi,
    SPHttpClient.configurations.v1,
    { headers: { Accept: "application/json;odata=nometadata", "odata-version": "" } }
  );

  if (!filesRes.ok) {
    const txt = await filesRes.text();

    if (filesRes.status === 401 || filesRes.status === 403) {
      throw new Error("Not signed in or no access to the template file(s). Please sign in and try again.");
    }

    throw new Error(txt || filesRes.statusText);
  }

  const filesJson = (await filesRes.json()) as { value: SPFileItem[] };
  let files = (filesJson.value || []).filter((f) => !!f?.ServerRelativeUrl);

  // Optional: only xlsx files
  if (options?.extensionFilter) {
    const ext = options.extensionFilter.toLowerCase();
    files = files.filter((f) => String(f.Name || "").toLowerCase().endsWith(ext));
  }

  if (!files.length) {
    throw new Error(`No template file found in library "${libraryTitle}".`);
  }

  // Latest file (already sorted)
  const chosen = files[0];

  // 3) Incognito-safe download: simulate user click on <a>
  const downloadUrl =
    `${siteUrl}/_layouts/15/download.aspx?SourceUrl=` +
    encodeURIComponent(chosen.ServerRelativeUrl);

  const a = document.createElement("a");
  a.href = downloadUrl;
  a.rel = "noopener noreferrer";
  a.target = "_blank"; // keep current page URL unchanged (prevents refresh-triggered re-download)
  a.download = String(chosen.Name || "ExcelTemplate.xlsx");
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  a.remove();
}

export const createExcelPendingRequest = async (
  context: WebPartContext,
  selectedExcelFile: File,
  userData: ISuparaItem[],
  usage: IUsageState,
  _totalAmount: number,
  _totalQty: number,
  uploadFileName?: string
): Promise<number> => {


  if (!selectedExcelFile) throw new Error("No file selected.");

  const siteUrl = context.pageContext.web.absoluteUrl;
  const nowIso = new Date().toISOString();
  const requesterUserId = await getCurrentUserId(context);
  const getApprovalAmountBucket = (amount: number): string => {
    if (amount <= 5000) return "LE_5000";
    if (amount <= 50000) return "5001_50000";
    if (amount <= 100000) return "50001_100000";
    if (amount <= 500000) return "100001_500000";
    return "GT_500000";
  };
  const getLookupIdByTitle = async (
    listTitle: string,
    title: string
  ): Promise<number | undefined> => {
    const trimmed = String(title || "").trim();
    if (!trimmed) return undefined;

    const safeList = listTitle.replace(/'/g, "''");
    const safeTitle = trimmed.replace(/'/g, "''");
    const url =
      `${siteUrl}/_api/web/lists/getbytitle('${safeList}')/items` +
      `?$top=1&$select=Id,Title&$filter=Title eq '${safeTitle}'`;

    const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: "application/json;odata=nometadata", "odata-version": "" }
    });
    if (!res.ok) return undefined;

    const json = (await res.json()) as { value?: Array<{ Id?: number }> };
    const id = Number(json.value?.[0]?.Id || 0);
    return id > 0 ? id : undefined;
  };
  const [purposeId, costCenterId] = await Promise.all([
    getLookupIdByTitle("PurposeMaster", usage.purpose),
    getLookupIdByTitle("CostCenterMaster", usage.costCenter),
  ]);
  if (!purposeId || !costCenterId) {
    throw new Error("Purpose/Cost Center lookup Id not found. Check master lists.");
  }

  /* =========================================================
     1ï¸âƒ£ BUILD APPROVAL CHAIN
  ========================================================= */

  // Initial Excel upload stage validates SKU/qty and parks request in Pending.
  // Do not persist computed amount at this stage.
  const initialQuantity = 0;
  const initialAmount = 0;
  // const chain = await buildApprovalChainFromLists(context, initialAmount);
  // const shouldStartAtStage1 = !!chain.firstStageApproverId;

  /* =========================================================
     2ï¸âƒ£ CREATE INVENTORY REQUEST (PENDING)
  ========================================================= */

  const headerBody: Record<string, unknown> = {
    Title: "TEMP",
    RequestCreatedOn: nowIso,
    LastUpdatedOn: nowIso,
    EntryMode: "Excel",

    RequesterId: requesterUserId,
    EmployeeID: userData?.[0]?.EmployeeID || "UNKNOWN",
    RequesterEmail: context.pageContext.user.email,
    JobPosition: userData?.[0]?.Jobposition || "",
    DepartmentId: userData?.[0]?.Department?.Id || undefined,
    PurposeId: purposeId,
    CostCenterId: costCenterId,

    ExpectedDeliveryDate: usage.expectedDelivery
      ? yyyyMmDdToIso(usage.expectedDelivery)
      : undefined,

    PlaceOfDelivery: usage.placeOfDelivery,
    PurposeDetails: usage.purposeDetail,
    Currency: "THB",

    TotalQuantity: initialQuantity,
    TotalAmount: initialAmount,
    ApprovalAmountBucket: getApprovalAmountBucket(initialAmount),

    WorkflowStatus: "Pending",

    //  PendingSince: nowIso,
    //RequestSubmittedOn: nowIso,

    HasAttachments: true,
    ExcelUploadFileName: uploadFileName || selectedExcelFile.name,
    ErrorFlag: false
  };

  const created = await spCreateItem(context, "InventoryRequests", headerBody);
  const requestId = created.Id;

  /* =========================================================
     3ï¸âƒ£ UPLOAD FILE TO DOCUMENT LIBRARY
  ========================================================= */

  const originalName = selectedExcelFile.name;
  const baseName = (uploadFileName || originalName || "").trim();

  const finalFileName = baseName.toLowerCase().endsWith(".xlsx")
    ? baseName
    : `${baseName}.xlsx`;

  const safeFileName = finalFileName.replace(/'/g, "");

  const libraryCandidates = ["InventoryRequestAttachments", "InventoryRequestAttachment"];
  let folderServerRelativeUrl = "";
  let resolvedLibraryTitle = "";
  let rootResolveError = "";

  for (let i = 0; i < libraryCandidates.length; i++) {
    const candidate = libraryCandidates[i];
    const safeCandidate = candidate.replace(/'/g, "''");
    const rootFolderRes = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${safeCandidate}')/RootFolder?$select=ServerRelativeUrl`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "odata-version": ""
        }
      }
    );

    if (!rootFolderRes.ok) {
      rootResolveError = await rootFolderRes.text();
      continue;
    }

    const rootFolderJson = (await rootFolderRes.json()) as { ServerRelativeUrl?: string };
    const rel = String(rootFolderJson.ServerRelativeUrl || "").trim();
    if (!rel) {
      rootResolveError = `Library "${candidate}" root folder is empty or missing.`;
      continue;
    }

    folderServerRelativeUrl = rel;
    resolvedLibraryTitle = candidate;
    break;
  }

  if (!folderServerRelativeUrl) {
    throw new Error(
      rootResolveError ||
      `Unable to resolve attachment library. Tried: ${libraryCandidates.join(", ")}`
    );
  }

  const uploadUrl =
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderServerRelativeUrl.replace(/'/g, "''")}')/Files/add(url='${safeFileName}',overwrite=true)`;

  // ðŸ”¥ IMPORTANT FIX: Use ArrayBuffer
  const fileBuffer = await selectedExcelFile.arrayBuffer();

  const uploadRes = await context.spHttpClient.post(
    uploadUrl,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/octet-stream",
        "odata-version": ""
      },
      body: fileBuffer
    }
  );

  if (!uploadRes.ok) {
    const text = await uploadRes.text();
    console.error("Upload error:", text);
    throw new Error(text || `File upload failed for library "${resolvedLibraryTitle}".`);
  }

  const uploadJson = await uploadRes.json();
  const fileRelUrl = uploadJson.ServerRelativeUrl;

  if (!fileRelUrl) {
    throw new Error("Upload succeeded but missing ServerRelativeUrl.");
  }

  /* =========================================================
     4ï¸âƒ£ UPDATE REQUIRED LIBRARY METADATA
  ========================================================= */

  const metaRes = await context.spHttpClient.post(
    `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${String(fileRelUrl).replace(/'/g, "''")}')/ListItemAllFields`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "odata-version": "",
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
      },
      body: JSON.stringify({
        FileLeafRef: safeFileName,
        Title: safeFileName,
        DocumentType: "Excel Upload",  // Must match your Choice value
        UploadedById: requesterUserId,
        UploadedOn: nowIso,
        RequestID: String(requestId),
        OriginalFileName: originalName   // ✅ store the original file name

      })
    }
  );
  if (!metaRes.ok) {
    const text = await metaRes.text();
    throw new Error(text || "File uploaded but metadata update failed.");
  }

  /* =========================================================
     5ï¸âƒ£ GENERATE TXN ID
  ========================================================= */

  const runningNo = await getNextMonthlyTxnSequence(context);
  const txnId = makeTxnId(runningNo, "ExcelUpload");

  await spUpdateItem(context, "InventoryRequests", requestId, {
    Title: txnId
  });

  return requestId;
};





export const validateExcelItemsWithFlowBulk = async (
  context: WebPartContext,
  excelItems: IItemRow[],
  flowUrl: string // keep param to avoid breaking calls
): Promise<{
  rowErrors: ItemRowErrors;
  messages: string[];
  finalItems: IItemRow[];
}> => {
  // Simpler flow enrichment: do not perform Excel content validations here.
  const rowErrors: ItemRowErrors = {};
  const messages: string[] = [];
  const finalItems: IItemRow[] = [];

  // Collect unique SKUs to fetch once and in parallel (with concurrency limit)
  const uniqueSkus: string[] = Array.from(
    new Set(
      excelItems
        .map((it) => String(it.sku || "").trim().toUpperCase())
        .filter((s) => s)
    )
  );

  const flowBySku = new Map<string, FlowItemResponse | undefined>();

  // Concurrency-limited fetch to avoid sequential waits and too many parallel requests.
  const concurrency = Math.min(12, Math.max(4, Math.ceil(uniqueSkus.length / 10)));
  let ptr = 0;
  const workers: Promise<void>[] = [];

  const fetchWorker = async (): Promise<void> => {
    while (ptr < uniqueSkus.length) {
      const i = ptr++;
      if (i >= uniqueSkus.length) break;
      const sku = uniqueSkus[i];
      try {
        const json = await fetchProductFromAws(context.httpClient, sku, flowUrl);
        flowBySku.set(sku, normalizeFlowToRow(json));
      } catch {
        flowBySku.set(sku, undefined);
      }
    }
  };

  for (let w = 0; w < concurrency; w++) {
    workers.push(fetchWorker());
  }

  await Promise.all(workers);

  // Map rows without awaiting per-row network calls.
  for (let i = 0; i < excelItems.length; i++) {
    const excelSku = String(excelItems[i].sku || "").trim().toUpperCase();
    const qty = Number(excelItems[i].qty || 0);

    const baseRow: IItemRow = {
      no: finalItems.length + 1,
      sku: excelSku,
      productCode: "",
      color: "-",
      size: "-",
      unitPrice: 0,
      qty: isNaN(qty) ? 0 : qty,
    };

    const flow = flowBySku.get(excelSku);

    if (!flow) {
      finalItems.push(baseRow);
      continue;
    }

    finalItems.push({
      ...baseRow,
      productCode: String(flow.ProductCode || ""),
      color: String(flow.ColourDesc || "-"),
      size: String(flow.SizeDesc || "-"),
      unitPrice: Number(flow.Original_sp || 0),
      maxStock: Number(flow.Warehouse_SOH_Qty ?? 0),
    });
  }

  // Renumber and return (no errors/messages added)
  const renumbered = finalItems.map((x, idx) => ({ ...x, no: idx + 1 }));
  return { rowErrors, messages, finalItems: renumbered };
};

export { IUsageState, ISuparaItem, IItemRow, FlowItemResponse };
