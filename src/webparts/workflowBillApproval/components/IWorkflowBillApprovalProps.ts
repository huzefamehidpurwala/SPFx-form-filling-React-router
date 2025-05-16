import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWorkflowBillApprovalProps {
  context: WebPartContext;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
