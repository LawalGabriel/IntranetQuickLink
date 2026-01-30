import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'IntranetQuickLinkWebPartStrings';
import IntranetQuickLink from './components/IntranetQuickLink';
import { IIntranetQuickLinkProps } from './components/IIntranetQuickLinkProps';

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
}

export default class IntranetQuickLinkWebPart extends BaseClientSideWebPart<IIntranetQuickLinkWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IIntranetQuickLinkProps> = React.createElement(
      IntranetQuickLink,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        listTitle: this.properties.listTitle || "QuickLink",
        headerColor: this.properties.headerColor || "#333333",
        itemBgColor: this.properties.itemBgColor || "#ffffff",
        itemTextColor: this.properties.itemTextColor || "#333333",
        itemHoverColor: this.properties.itemHoverColor || "#f3f2f1",
        iconColor: this.properties.iconColor || "#0078d4",
        maxItems: this.properties.maxItems || 12,
        itemsPerRow: this.properties.itemsPerRow || 4,
        showBorder: this.properties.showBorder !== false, 
        borderColor: this.properties.borderColor || "#e1e1e1",
        headerBgColor: this.properties.headerBgColor || "#f8f9fa", 
      headerTitle: this.properties.headerTitle || "QUICK LINKS", 
      bodyBgColor: this.properties.bodyBgColor || "#f8f9fa",
      bodyTextColor: this.properties.bodyTextColor || "#333333"
         
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Ensure Default Chrome Is Disabled
    this._ensureDefaultChromeIsDisabled();
    
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _ensureDefaultChromeIsDisabled(): void {
    const displayElements = ['#SuiteNavWrapper', '#spSiteHeader', '.sp-appBar', '.ms-CommandBar', '.commandBarWrapper', '#spCommandBar', '.ms-SPLegacyFabric', '.ms-footer', '.sp-pageLayout-footer', '.ms-workbenchFooter'];
    const widthElements = ["#workbenchPageContent", '.CanvasZone', '.SPCanvas-canvas'];
    
    displayElements.forEach(selector => {
      document.querySelectorAll(selector).forEach(element => {
        (element as HTMLElement).style.display = 'none';
      });
    });

    widthElements.forEach(selector => {
      document.querySelectorAll(selector).forEach(element => {
        (element as HTMLElement).style.maxWidth = 'none';
      });
    });

    if (this._isMobileDevice()) {
      const mobileElements = [
        '.spMobileHeader',
        '.ms-FocusZone',
        '.ms-CommandBar',
        '.spMobileNav',
        '#O365_MainLink_NavContainer',
        '.ms-Nav',
        '.ms-Nav-item'
      ];
      mobileElements.forEach(selector => {
        document.querySelectorAll(selector).forEach(element => {
          (element as HTMLElement).style.display = 'none';
        });
      });
    }
  }

  private _isMobileDevice(): boolean {
    const userAgent = navigator.userAgent.toLowerCase();
    const isMobile = /iphone|ipod|ipad|android|blackberry|windows phone/i.test(userAgent);
    return isMobile || window.innerWidth <= 768;
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

 protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
              PropertyPaneTextField('listTitle', {
                label: 'List Title',
                value: this.properties.listTitle || 'QuickLink',
                description: 'Enter the name of the SharePoint list containing quick links'
              }),
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel,
                value: this.properties.description
              }),
              // NEW: Header Title Field
              PropertyPaneTextField('headerTitle', {
                label: 'Header Title',
                description: 'Enter the title text for the header',
                value: this.properties.headerTitle || 'QUICK LINKS'
              })
            ]
          },
          {
            groupName: 'Display Settings',
            groupFields: [
              PropertyPaneSlider('maxItems', {
                label: 'Maximum Items to Display',
                min: 1,
                max: 50,
                step: 1,
                value: this.properties.maxItems || 12,
                showValue: true
              }),
              PropertyPaneSlider('itemsPerRow', {
                label: 'Items Per Row (Desktop)',
                min: 2,
                max: 6,
                step: 1,
                value: this.properties.itemsPerRow || 4,
                showValue: true
              }),
              PropertyPaneToggle('showBorder', {
                label: 'Show Item Borders',
                checked: this.properties.showBorder !== false
              })
            ]
          },
          {
            groupName: 'Color Settings',
            groupFields: [
              PropertyPaneTextField('headerBgColor', { // NEW: Header Background Color
                label: 'Header Background Color',
                description: 'Enter hex color code (e.g., #f8f9fa)',
                value: this.properties.headerBgColor || '#f8f9fa'
              }),
              PropertyPaneTextField('headerColor', {
                label: 'Header Text Color',
                description: 'Enter hex color code (e.g., #333333)',
                value: this.properties.headerColor || '#333333'
              }),
              PropertyPaneTextField('bodyBgColor', { // NEW: Body Background Color
                label: 'Body Background Color',
                description: 'Enter hex color code for the container background (e.g., #f8f9fa)',
                value: this.properties.bodyBgColor || '#f8f9fa'
              }),
              PropertyPaneTextField('itemBgColor', {
                label: 'Item Background Color',
                description: 'Enter hex color code (e.g., #FFFFFF)',
                value: this.properties.itemBgColor || '#FFFFFF'
              }),
              PropertyPaneTextField('itemTextColor', {
                label: 'Item Text Color',
                description: 'Enter hex color code (e.g., #333333)',
                value: this.properties.itemTextColor || '#333333'
              }),
              PropertyPaneTextField('itemHoverColor', {
                label: 'Item Hover Color',
                description: 'Enter hex color code (e.g., #F3F2F1)',
                value: this.properties.itemHoverColor || '#F3F2F1'
              }),
              PropertyPaneTextField('iconColor', {
                label: 'Icon Color',
                description: 'Enter hex color code (e.g., #0078D4)',
                value: this.properties.iconColor || '#0078D4'
              }),
              PropertyPaneTextField('borderColor', {
                label: 'Border Color',
                description: 'Enter hex color code (e.g., #E1E1E1)',
                value: this.properties.borderColor || '#E1E1E1'
              })
            ]
          }
        ]
      }
    ]
  };
}
}