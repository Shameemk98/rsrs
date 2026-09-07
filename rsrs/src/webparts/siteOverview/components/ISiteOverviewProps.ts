import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISiteOverviewProps {
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
  SiteListName:string;
  SiteAdditionalSupportListName:string;
}
