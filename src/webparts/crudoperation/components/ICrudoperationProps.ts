import { SPHttpClient } from "@microsoft/sp-http";

export interface ICrudoperationProps {
  description: string;
  listName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}
