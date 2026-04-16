import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEformProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  flowUrl: string;
  ApprovalFlowUrl: string;
}
