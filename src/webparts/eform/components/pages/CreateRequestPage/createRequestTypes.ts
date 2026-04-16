import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Route } from "../../Eform";
import { RouteState } from "../../spHelpers";

export interface IInventoryItemSP {
  SKUCode?: string;
  ProductDescription?: string;
  Color?: string;
  Size?: string;
  UnitPrice?: number | string;
  Quantity?: number | string;
    Warehouse_SOH_Qty?: number;
    StockAvailability?:number;
}

export interface CreateRequestPageProps {
  onSubmit: () => void;
  context: WebPartContext;
  draftId?: number;
    onNavigate: (route: Route, state?: RouteState) => void;
    flowUrl: string;
}

export interface IUsageState {
  purpose: string;
  costCenter: string;
  expectedDelivery: string; // yyyy-MM-dd
  placeOfDelivery: string;
  purposeDetail: string;
}

export interface ISPUserLookup {
  Id: number;
  Title: string;
  EMail: string;
}

export interface ISuparaItem {
  Title: string;
  EmployeeID: string;
Email:string;
  // Person/Group
  User: {
    Id: number;
    Title: string;
    EMail: string;
  };

  // Choice / Number / YesNo
  Jobposition?: string;
  RoleLevel?: number;
  IsActive: boolean;

  // Lookup fields
  Department?: { Id: number; Title: string };
  SecurityRole?: { Id: number; Title: string };
}

export type SPFileItem = {
  Name: string;
  ServerRelativeUrl: string;
  TimeLastModified?: string;
};



export interface IItemRow {
  no: number;
  sku: string;
  productCode: string;
  color: string;
  size: string;
  unitPrice: number;
  qty: number;
  maxStock?: number;
}

/** Approval Matrix list row */
export interface ApprovalMatrixRow {
  Id: number;
  Title?: string;
  Amount?: number;
  StageNumber?: number;
  ApproverType?: string;
  Approver?: ISPUserLookup;
  EscalationApprover?: ISPUserLookup;
  SLAHours?: number;
  IsActive?: boolean;
}

/** POST bodies for SharePoint create item (typed; Draft allows optional fields) */
export interface InventoryRequestsCreateBody {
  Title: string;

  RequestCreatedOn?: string;
  RequestSubmittedOn?: string;
  LastUpdatedOn?: string;
  PendingSince?: string;

  EntryMode?: string;

  RequesterId?: number;
  EmployeeID?: string;
  RequesterEmail?: string;
  JobPosition?: string;
  Department?: { Id: number; Title: string };

  ApprovalType?: string;

  PurposeId?: number;
  CostCenterId?: number;

  ExpectedDeliveryDate?: string;
  PlaceOfDelivery?: string;
  PurposeDetails?: string;

  TotalQuantity?: number;
  TotalAmount?: number;

  ApprovalAmountBucket?: string;
  WorkflowStatus?: string;
  WorkflowStage?: number;

  CurrentApproverId?: number;

  IsEditableByRequester?: boolean;
  HasAttachments?: boolean;

  ErrorFlag?: boolean;
  ErrorMessage?: string;
  ErrorDate?: string;
}

export interface InventoryRequestItemsCreateBody {
  Title: string;

  LineNumber?: number;
  SKUCode?: string;
  ProductDescription?: string;
  Color?: string;
  Size?: string;
  UnitPrice?: number;
  Quantity?: number;
  LineTotal?: number;

  LookupSource?: string;
  RowStatus?: string;

  RequestIDId: number;
}

export interface ValidationResult {
  ok: boolean;
  submitMessage: string;
  showAsPopup?: boolean;
  fieldErrors: UsageFieldErrors;
  itemErrors: ItemRowErrors;
}

/* =========================================================
   ✅ Flow Types
   ========================================================= */
export interface FlowItemResponse {
  Barcode?: string;
  ProductCode?: string;
  ColourDesc?: string;
  SizeDesc?: string;
  Original_sp?: number;
  Warehouse_SOH_Qty?: number;
}
export type InventoryMode = "manual" | "excel";


export type PurposeOption = { Id: number; Title: string };
export type CostCenterOption = {
  Id: number;
  Title: string;
  CostCenterCode: string;
  Department?: { Id: number; Title: string };
};




/** OData response shapes */
export type ODataV4<T> = { value: T[] };
export type ODataV3<T> = { d: { results: T[] } };
export type ODataResponse<T> = ODataV4<T> | ODataV3<T>;



/* =========================================================
   ✅ Validation Types
   ========================================================= */
export type UsageFieldKey = keyof IUsageState;
export type UsageFieldErrors = Partial<Record<UsageFieldKey, string>>;
export type ItemRowErrors = Record<number, string>;

export interface IRequesterDirectoryItem {
  Id: number;
  JobPosition?: string;

  IsOnLeave?: boolean;   // ✅ ADD THIS LINE

  User: {
    Id: number;
    Title: string;
    EMail: string;
  };

  Department?: {
    Id: number;
    Title: string;
  };

  DirectManager?: {
    Id: number;
    Title?: string;
    EMail: string;
  };
}

export interface IApprovalMatrixItem {
  Id: number;
  MinAmount: number;
  MaxAmount: number;
  Order: number;
  Role: string;
  ApprovalRole?: string;
  Department?: {
    Id: number;
    Title?: string;
  };
  SingleLevelApproval?: boolean | number;
  MultiLevelApproval?: boolean | number;


}
