import * as React from 'react';
import * as ReactDom from 'react-dom';
import {  Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { SPHttpClient } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EformWebPartStrings';
import Eform from './components/Eform';
import { IEformProps } from './components/IEformProps';
import '../../../assets/dist/main.css';


export interface IEformWebPartProps {
  description: string;
  flowUrl: string;
  ApprovalFlowUrl: string;
}

export default class EformWebPart extends BaseClientSideWebPart<IEformWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _flowUrlFromConfig: string = '';
  private _approvalFlowUrlFromConfig: string = '';

  public render(): void {
           document.body.classList.add("gq-eform-fullpage");
    this.domElement.id = 'tw';
    const resolvedFlowUrl = this._flowUrlFromConfig || this.properties.flowUrl || "";
    const resolvedApprovalFlowUrl = this._approvalFlowUrlFromConfig || this.properties.ApprovalFlowUrl || "";


    // // ✅ Hide SharePoint chrome ONLY in View mode
    // if (this.displayMode === DisplayMode.Read) {
    //   document.body.classList.add("gq-eform-fullpage");
    // } else {
    //   document.body.classList.remove("gq-eform-fullpage");
    // }


    const element: React.ReactElement<IEformProps> = React.createElement(
      Eform,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        flowUrl: resolvedFlowUrl,
        ApprovalFlowUrl: resolvedApprovalFlowUrl,


      }
    );


    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage()
      .then(message => {
        this._environmentMessage = message;
      })
      .then(() => this._loadWorkflowConfiguration());
  }

    protected onDispose(): void {
      document.body.classList.remove("gq-eform-fullpage");
 
    ReactDom.unmountComponentAtNode(this.domElement);
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

  private _readConfigUrlValue(item: {
    Url?: string | { Url?: string };
    ConfigValue?: string;
  }): string {
    if (typeof item.Url === "string") return item.Url.trim();
    if (item.Url && typeof item.Url.Url === "string") return item.Url.Url.trim();
    if (typeof item.ConfigValue === "string") return item.ConfigValue.trim();
    return "";
  }

  private async _loadWorkflowConfiguration(): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const apiUrl =
        `${webUrl}/_api/web/lists/getbytitle('WorkflowConfiguration')/items` +
        `?$select=Title,Url,ConfigValue,IsActive&$top=200`;

      const response = await this.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return;
      }

      const json = await response.json();
      const items = Array.isArray(json?.value) ? json.value : [];

      for (const item of items) {
        const key = String(item?.Title || "").trim().toLowerCase();
        const isActive = item?.IsActive === undefined ? true : !!item.IsActive;
        const urlValue = this._readConfigUrlValue(item);

        if (!isActive || !urlValue) continue;

        if (key === "flowurl") this._flowUrlFromConfig = urlValue;
        if (key === "approvalflowurl") this._approvalFlowUrlFromConfig = urlValue;
      }
    } catch (error) {
      console.error("Failed to load WorkflowConfiguration list values.", error);
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const theme = currentTheme as IReadonlyTheme & {
      isInverted?: boolean;
      semanticColors?: {
        bodyText?: string;
        link?: string;
        linkHovered?: string;
      };
    };

    this._isDarkTheme = !!theme.isInverted;

    const { semanticColors } = theme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
      this.domElement.style.setProperty('--link', semanticColors.link || '');
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
    }
  }


  // protected onDispose(): void {
  //   ReactDom.unmountComponentAtNode(this.domElement);
  // }

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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('flowUrl', {
                  label: 'Power Automate Flow URL',
                  description: 'Full URL including sig parameter',
                  value: this.properties.flowUrl || "",
                }),
                PropertyPaneTextField('ApprovalFlowUrl', {
                  label: 'Power Automate Approval Flow URL',
                  description: 'Full URL including sig parameter',
                  value: this.properties.ApprovalFlowUrl || "",
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
