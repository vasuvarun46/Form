

export type SecurityPermissions = {
  submit: boolean;
  viewOwn: boolean;
  viewTeam: boolean;
  viewAll: boolean;

  approve: boolean;
  download: boolean;

  approveTeam: boolean;
  approveAll: boolean;
  editMyRequest: boolean;

  exportOwnReport: boolean;
  exportAllReport: boolean;
  exportTeamReport: boolean;

  viewExportOwnReport: boolean;
  viewExportTeamReport: boolean;
  viewExportAllReport: boolean;

  emailNotifySubmit: boolean;
  emailNotifyAfterA1Approved: boolean;
  emailNotifyFinalApproval: boolean;
  emailOwnBulkUploadFinished: boolean;
  
};

export const makePerms = (overrides: Partial<SecurityPermissions> = {}): SecurityPermissions => ({
  submit: false,
  viewOwn: false,
  viewTeam: false,
  viewAll: false,

  approve: false,
  download: false,

  approveTeam: false,
  approveAll: false,
  editMyRequest: false,

  exportOwnReport: false,
  exportAllReport: false,
  exportTeamReport: false,

  viewExportOwnReport: false,
  viewExportTeamReport: false,
  viewExportAllReport: false,

  emailNotifySubmit: false,
  emailNotifyAfterA1Approved: false,
  emailNotifyFinalApproval: false,
  emailOwnBulkUploadFinished: false,

  ...overrides,
});
export type ToastType = "success" | "error" | "draft" | "pending";
 
export type RouteState =
  | {
      toastType?: ToastType;
      toastMessage?: string;
    }
  | undefined;
 
