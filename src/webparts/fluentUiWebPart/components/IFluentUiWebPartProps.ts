import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFluentUiWebPartProps {
  listTitle: string;
  listViewName: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
