import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SecurityPermissions } from "../../spHelpers";

/* ==========================================
   STATUS
========================================== */

export type Status =
  | "Submitted"
  | "In Approval"
  | "Approved"
  | "Rejected"
  | "Draft";

/* ==========================================
   TABLE ROW
========================================== */

export interface Row {
  id: string;
  status: Status;
  requestor: string;
  approver: string;
  department: string;
  purpose: string;
  PurposeDetails: string;
  costCenter: string;
  submitted: string;
  approved: string;
  expected: string;
  qty: number;
  price: number;
  spId: number;
  expectedIso?: string;
  submittedIso?: string;
  modifiedIso?: string;
  lastActivityIso?: string;
  ApproverNames?:string

}

/* ==========================================
   SHAREPOINT PERSON
========================================== */

export interface SPPerson {
  Id: number;
  Title: string;
  EMail: string;
}

/* ==========================================
   REQUESTER DIRECTORY ITEM
========================================== */

export interface RequesterDirectoryItem {
  Id: number;
  User?: SPPerson;
  Department?: { Id: number; Title: string };
  IsActive?: boolean;
}

/* ==========================================
   INVENTORY REQUEST ITEM
========================================== */

export interface InventoryRequestSPItem {
  Id: number;
  WorkflowStatus?: Status;
  Requester?: SPPerson;
  CurrentApprover?: SPPerson;
  Department?: { Id: number; Title: string };
  Purpose?: { Id: number; Title: string };
  CostCenter?: { Id: number; Title: string };
  PurposeDetails?: string;
  RequestSubmittedOn?: string;
  FinalDecisionDate?: string;
  ExpectedDeliveryDate?: string;
  TotalQuantity?: number;
  TotalAmount?: number;
  Modified?: string;
}

/* ==========================================
   GENERIC SP RESPONSE
========================================== */

export interface SPListResponse<T> {
  value: T[];
}

/* ==========================================
   PAGE PROPS
========================================== */

export interface AllRequestPageProps {
  context: WebPartContext;
  onOpenExportPdf: (spId: number) => void;
  onDownloadPdf: (spId: number) => void;
  permissions: SecurityPermissions;
   
}
