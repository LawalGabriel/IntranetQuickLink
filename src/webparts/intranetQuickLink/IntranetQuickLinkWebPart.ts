import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,PropertyPaneSlider
} from '@microsoft/sp-property-pane';
//import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'IntranetQuickLinkWebPartStrings';
import IntranetQuickLink from './components/IntranetQuickLink';
import { IIntranetQuickLinkProps } from './components/IIntranetQuickLinkProps';

export interface IIntranetQuickLinkWebPartProps {
  description: string;
  listTitle: string;
  headerColor: string;
  rowColor1: string;
  rowColor2: string;
  rowTextColor: string;
  rowHoverColor1: string;
  rowHoverColor2: string;
  maxRows: number;
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
        rowColor1: this.properties.rowColor1 || "#ffffff",
        rowColor2: this.properties.rowColor2 || "#f8f9fa",
        rowTextColor: this.properties.rowTextColor || "#333333",
        rowHoverColor1: this.properties.rowHoverColor1 || "#f3f2f1",
        rowHoverColor2: this.properties.rowHoverColor2 || "#f3f2f1",
        maxRows: this.properties.maxRows || 4
        
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
    //'.SuiteNavWrapper','#spSiteHeader','.sp-appBar','#sp-appBar','#workbenchPageContent', '.SPCanvas-canvas', '.CanvasZone', '.ms-CommandBar', '#spSiteHeader', '.commandBarWrapper'
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

    //document.querySelector<HTMLElement>('#spCommandBar')?.style.setProperty('min-height', '0', 'important');




    // Check if the device is mobile and apply mobile-specific styles
    if (this._isMobileDevice()) {
      const mobileElements = [
        '.spMobileHeader',
        '.ms-FocusZone',
        '.ms-CommandBar',
        '.spMobileNav',
        '#O365_MainLink_NavContainer', // Waffle (App Launcher)
        '.ms-Nav', // Additional possible mobile navigation elements
        '.ms-Nav-item' // Possible item within the navigation
      ];
      mobileElements.forEach(selector => {
        document.querySelectorAll(selector).forEach(element => {
          (element as HTMLElement).style.display = 'none';
        });
      });
    }
  }

  private _isMobileDevice(): boolean {
    // Check if the user is on a mobile device based on user agent or screen width
    const userAgent = navigator.userAgent.toLowerCase();
    const isMobile = /iphone|ipod|ipad|android|blackberry|windows phone/i.test(userAgent);
    return isMobile || window.innerWidth <= 768; // Custom breakpoint for mobile
  }


  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
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
    const {
      semanticColors
    } = currentTheme;

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
                label: strings.DescriptionFieldLabel
              }),
              PropertyPaneTextField('description', {
                label: strings.DescriptionFieldLabel
              }),
              PropertyPaneTextField('headerColor', {
                label: 'Header Color',
                description: 'Enter a hex color code (e.g., #333333)',
                value: '#333333'
              }),
              PropertyPaneTextField('rowColor1', {
                label: 'Row Background Color 1',
                description: 'First alternating color (e.g., #FFFFFF)',
                value: '#FFFFFF'
              }),
              PropertyPaneTextField('rowColor2', {
                label: 'Row Background Color 2',
                description: 'Second alternating color (e.g., #F8F9FA)',
                value: '#F8F9FA'
              }),
              PropertyPaneTextField('rowTextColor', {
                label: 'Row Text Color',
                description: 'Enter a hex color code (e.g., #333333)',
                value: '#333333'
              }),
              PropertyPaneTextField('rowHoverColor1', {
                label: 'Row Hover Color 1',
                description: 'Hover color for first alternating row (e.g., #F3F2F1)',
                value: '#F3F2F1'
              }),
              PropertyPaneSlider('maxRows', {
  label: 'Number of Rows to Display',
  min: 1,
  max: 10,
  step: 1,
  value: this.properties.maxRows || 4,
  showValue: true
}),
              PropertyPaneTextField('rowHoverColor2', {
                label: 'Row Hover Color 2',
                description: 'Hover color for second alternating row (e.g., #E8E7E6)',
                value: '#E8E7E6'
              })
            ]
          }
        ]
      }
    ]
  };
}
}
