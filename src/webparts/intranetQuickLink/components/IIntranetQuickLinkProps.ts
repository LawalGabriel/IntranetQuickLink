import { WebPartContext } from "@microsoft/sp-webpart-base";


export interface IIntranetQuickLinkProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  listTitle: string;
  headerColor: string;
  rowColor1: string;
  rowColor2: string;
  rowTextColor: string;
  rowHoverColor1: string;
  rowHoverColor2: string;
  maxRows: number;

}

export interface ILinkItem {
  Title: string;
  Link: string | { Url: string; Description?: string } ;
  Status?: boolean;
  Id: number;
  Created: string;
  Author: {
   title: string; 
   id: number;}
  
   ;

}
