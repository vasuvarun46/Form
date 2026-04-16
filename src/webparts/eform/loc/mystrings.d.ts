declare interface IEformWebPartStrings {
  Settings: string ;
  Settings: string ;
   
    EmptyTeamRequests: string  ;
    EmptyTitle: string ;
    EmptyDescription: string ;
    LoadingDescription: string ;
    LoadingTitle: string ;
  ExcelSelectFileFirst:string;
  ExcelUploadFailed: any;
  SkuRequired: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  DescriptionFieldLabel: string;
  PurposeRequired: string;
  CostCenterRequired: string;
  DraftSavedSuccess: string;
  SelectManualModeFirst: string;
  ErrorTitle: string;
  SkuNotFound: string;
  InvalidProductData: string;
  SkuOutOfStock: string;
  QuantityCannotBeZero: string;
  OnlyAvailable: string;
  ExcelFileRequired: string;
  ExcelValidateBeforeSubmit: string;
  CancelInvalidRows: string;
  DraftUpdatedSuccess: string;
  RequestSubmittedSuccess: string;
    ExcelSkuMissing: string;
  ExcelQtyInvalid: string;
  ExcelSkuNotFound: string;
  ExcelOutOfStock: string;
  ExcelOnlyAvailable: string; // uses {0}
  ExcelHasErrors: string;
  ConfirmApprove: string;
ConfirmReject: string;
ConfirmBypass: string;
Approve: string;
Reject: string;
Bypass: string;
Cancel: string;
RequestListTitle: string;
First: string;
Previous: string;
Next: string;
Last: string;
 
 
}
 
declare module 'EformWebPartStrings' {
  const strings: IEformWebPartStrings;
  export = strings;
 
}
 