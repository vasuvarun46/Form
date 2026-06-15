export type ExportActionInfo = {
  canAct: boolean;
  mode?: "approve" | "bypass";
  showReject?: boolean;
  isApproverWaiting?: boolean;  // user is in approval chain but not yet their turn
};

export interface ExportRow {
  id: string | number;
  status: string;
  requestor: string;
  approver: string;
  purpose: string;
  purposeDetail: string;
  costCenter: string;
  submitted: string;
  approved: string;
  expected: string;
  qty: number;
  price: number;
}
