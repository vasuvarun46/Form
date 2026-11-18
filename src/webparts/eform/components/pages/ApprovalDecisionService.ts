/* =========================================================
   ApprovalDecisionService.ts
   Approve/reject for the ExportPDF / Transaction-detail screen
   (used by Eform.tsx submitExportDecision).

   This is now a THIN WRAPPER around the shared engine in
   ../ApprovalProcessor.ts. It only reads the approval chain from
   SharePoint (this screen doesn't have the row preloaded) and maps the
   DecisionInput to processApproval(). All approval logic — stock check,
   totals, column writes, emails, multi-stage hand-off — lives in one place
   shared with TeamService, so the two screens can no longer drift apart.
========================================================= */

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import {
  ApprovalLevel,
  FlowAction,
  processApproval,
} from "../ApprovalProcessor";

export type { FlowAction } from "../ApprovalProcessor";
export type { SkuMismatch } from "../ApprovalProcessor";

export interface DecisionInput {
  context: WebPartContext;
  approvalFlowUrl: string;
  apiKey: string;
  productFlowUrl?: string;

  requestId: number; // SP Id
  transactionId: string; // Display TX ID (Title)
  actorEmail: string;
  position: string;
  action: FlowAction;
  comments: string;

  // Optional — passed when caller already has them
  totalQuantity?: number;
  totalTagPrice?: number;
  purpose?: string;
  status?: string;
}

export interface DecisionResult {
  success: boolean;
  message: string;
  newStatus?: string;
}

/* =========================================================
   Read the approval chain for this request (header read).
   Only needed on Approve — mirrors the original behaviour.
========================================================= */
async function readApprovalChain(
  context: WebPartContext,
  siteUrl: string,
  requestId: number
): Promise<{
  approvalLevels: ApprovalLevel[];
  hasMultipleApprovals: boolean;
  workflowStage: number;
}> {
  let approvalLevels: ApprovalLevel[] = [];
  let hasMultipleApprovals = false;
  let workflowStage = 0;

  try {
    const chainRes = await context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items(${requestId})` +
        `?$select=ApprovalChainJSON,HasMultipleApprovals,WorkflowStage`,
      SPHttpClient.configurations.v1
    );
    if (chainRes.ok) {
      const chainData = (await chainRes.json()) as {
        ApprovalChainJSON?: string;
        HasMultipleApprovals?: boolean;
        WorkflowStage?: number;
      };
      hasMultipleApprovals = chainData?.HasMultipleApprovals === true;
      workflowStage = Number(chainData?.WorkflowStage ?? 0) || 0;
      try {
        const parsed = JSON.parse(String(chainData?.ApprovalChainJSON || "")) as {
          allApprovalLevels?: Array<{
            order?: number;
            approverEmail?: string;
            canApproveRequest?: boolean;
          }>;
        };
        if (Array.isArray(parsed?.allApprovalLevels)) {
          approvalLevels = parsed.allApprovalLevels
            .map((l) => ({
              order: Number(l?.order ?? 0),
              approverEmail: String(l?.approverEmail || "").trim().toLowerCase(),
              canApproveRequest: l?.canApproveRequest === true,
            }))
            .filter((l) => l.order > 0 && !!l.approverEmail);
        }
      } catch {
        /* ignore chain parse errors */
      }
    } else {
      console.warn(
        "[EForm Approval] Could not read ApprovalChainJSON/HasMultipleApprovals " +
          `(status ${chainRes.status}). isFinalApprover stays false — column write will be skipped.`
      );
    }
  } catch (err) {
    console.warn("[EForm Approval] approval-chain read failed (non-blocking):", err);
  }

  return { approvalLevels, hasMultipleApprovals, workflowStage };
}

/* =========================================================
   Submit decision — thin wrapper over processApproval.
========================================================= */
export async function submitApprovalDecision(
  input: DecisionInput
): Promise<DecisionResult> {
  if (!input.approvalFlowUrl) throw new Error("Approval flow URL is missing");
  if (!input.apiKey) throw new Error("Security key is missing");

  const siteUrl = input.context.pageContext.web.absoluteUrl;

  // The chain is only consulted on Approve (final-approver detection + hand-off).
  const chain =
    input.action === "Approve"
      ? await readApprovalChain(input.context, siteUrl, input.requestId)
      : { approvalLevels: [] as ApprovalLevel[], hasMultipleApprovals: false, workflowStage: 0 };

  const result = await processApproval({
    context: input.context,
    siteUrl,
    approvalFlowUrl: input.approvalFlowUrl,
    flowApiKey: input.apiKey,
    productFlowUrl: input.productFlowUrl,

    requestId: input.requestId,
    transactionId: input.transactionId,
    actorEmail: input.actorEmail,
    position: input.position,
    action: input.action,
    comments: input.comments,
    purpose: input.purpose,
    status: input.status,

    approvalLevels: chain.approvalLevels,
    hasMultipleApprovals: chain.hasMultipleApprovals,
    workflowStage: chain.workflowStage,

    requestedTotalQty: input.totalQuantity ?? 0,
    requestedTotalAmount: input.totalTagPrice ?? 0,
  });

  return {
    success: result.success,
    message: result.message,
    newStatus: result.newStatus,
  };
}
