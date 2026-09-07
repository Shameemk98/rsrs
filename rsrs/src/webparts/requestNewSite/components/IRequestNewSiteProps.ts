import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IRequestNewSiteProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string; 
  context: WebPartContext;
  countryListName:string;
  siteTypeListName:string;
  LegacyCompanyListName:string;
  ContactsListName:string;
  OperatingGroupListName:string;
}
