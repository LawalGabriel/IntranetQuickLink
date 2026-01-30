/* eslint-disable @typescript-eslint/no-explicit-any */
export interface IIntranetQuickLinkProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  listTitle: string;
  context: any;
  headerBgColor: string; 
  headerTitle: string;    
  bodyBgColor: string; 
  bodyTextColor: string; 
  
  // Color properties
  headerColor?: string;
  itemBgColor?: string;
  itemTextColor?: string;
  itemHoverColor?: string;
  iconColor?: string;
  borderColor?: string;
  
  // Display properties
  maxItems?: number;
  itemsPerRow?: number;
  showBorder?: boolean;

  headerFontSize?: string;
  headerFontWeight?: string;
  headerHeight?: string;
}

export interface ILinkItem {
  Id: number;
  Title: string;
  Link: string | { Url: string };
  Status: number;
  IconName?: string;
}

export interface IIntranetQuickLinkWebPartProps {
  bodyTextColor: string;
  headerTitle: string;
  headerBgColor: string;
  bodyBgColor: string;
  description: string;
  listTitle: string;
  headerColor: string;
  itemBgColor: string;
  itemTextColor: string;
  itemHoverColor: string;
  iconColor: string;
  maxItems: number;
  itemsPerRow: number;
  showBorder: boolean;
  borderColor: string;
  
  // NEW: Header styling properties
  headerFontSize: string;
  headerFontWeight: string;
  headerHeight: string;
}