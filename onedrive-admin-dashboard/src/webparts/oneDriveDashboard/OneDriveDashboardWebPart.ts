import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'OneDriveDashboardWebPartStrings';
import OneDriveDashboard from './components/OneDriveDashboard';
import { IOneDriveDashboardProps } from './components/IOneDriveDashboardProps';

export interface IOneDriveDashboardWebPartProps {
  description: string;
  useMockData: boolean;
  apiBaseUrl: string;
  apiResourceUri: string;
}

export default class OneDriveDashboardWebPart extends BaseClientSideWebPart<IOneDriveDashboardWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _theme: IReadonlyTheme | undefined;

  public render(): void {
    const element: React.ReactElement<IOneDriveDashboardProps> = React.createElement(
      OneDriveDashboard,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        theme: this._theme,
        useMockData: this.properties.useMockData === true,
        apiBaseUrl: this.properties.apiBaseUrl || '',
        apiResourceUri: this.properties.apiResourceUri || ''
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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

    this._theme = currentTheme;
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('apiBaseUrl', {
                  label: 'Azure Function API Base URL',
                  description: 'Base URL of the secure Azure Function backend, e.g. https://<app-name>.azurewebsites.net (no trailing slash).'
                }),
                PropertyPaneTextField('apiResourceUri', {
                  label: 'Azure Function API Resource (App ID URI)',
                  description: 'App ID URI exposed by the Azure Function\'s Entra ID app registration (Expose an API), e.g. api://<client-id> or api://<app-name>.azurewebsites.net. Required for authenticated calls via AadHttpClient.'
                }),
                PropertyPaneToggle('useMockData', {
                  label: 'Use sample data (local development only)',
                  onText: 'Sample data',
                  offText: 'Live Graph data (via Azure Function)'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
