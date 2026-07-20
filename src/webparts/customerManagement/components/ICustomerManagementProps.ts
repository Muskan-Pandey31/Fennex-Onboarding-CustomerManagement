import { SPHttpClient } from "@microsoft/sp-http";

export interface ICustomerManagementProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  siteUrl: string;
  spHttpClient: SPHttpClient;
}