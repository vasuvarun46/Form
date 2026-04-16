import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import {
  RequesterDirectoryItem,
  InventoryRequestSPItem,
  SPListResponse
} from "./AllRequestTypes";

/* =====================================================
   GET USER DEPARTMENT
===================================================== */



export async function getUserDepartment(
  context: WebPartContext
): Promise<string | null> {

  const siteUrl = context.pageContext.web.absoluteUrl;
  const userId = context.pageContext.legacyPageContext.userId;

  const response = await context.spHttpClient.get(
    `${siteUrl}/_api/web/lists/getbytitle('RequesterDirectory')/items` +
      `?$select=User/Id,Department/Title` +
      `&$expand=User,Department` +
      `&$filter=User/Id eq ${userId} and IsActive eq 1`,
    SPHttpClient.configurations.v1
  );

  const json: SPListResponse<RequesterDirectoryItem> =
    await response.json();

  return json.value[0]?.Department?.Title ?? null;
}

/* =====================================================
   GET PAGED DATA
===================================================== */

export async function getInventoryRequestsPaged(
  context: WebPartContext,
  departmentTitle: string,
  pageSize: number,
  skip: number
): Promise<InventoryRequestSPItem[]> {

  const siteUrl = context.pageContext.web.absoluteUrl;

  const url =
    `${siteUrl}/_api/web/lists/getbytitle('InventoryRequests')/items` +
    `?$top=${pageSize}` +
    `&$skip=${skip}` +
    `&$orderby=Id desc` +
    `&$filter=WorkflowStatus ne 'Draft' and WorkflowStatus ne 'Failed' and Department/Title eq '${departmentTitle}'` +
    `&$select=` +
    `Id,WorkflowStatus,TotalQuantity,TotalAmount,` +
    `PurposeDetails,RequestSubmittedOn,FinalDecisionDate,ExpectedDeliveryDate,` +
    `Requester/Title,CurrentApprover/Title,Department/Title,Purpose/Title,CostCenter/Title` +
    `&$expand=Requester,CurrentApprover,Department,Purpose,CostCenter`;

  const response = await context.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1
  );

  const json: SPListResponse<InventoryRequestSPItem> =
    await response.json();

  return json.value;
}
