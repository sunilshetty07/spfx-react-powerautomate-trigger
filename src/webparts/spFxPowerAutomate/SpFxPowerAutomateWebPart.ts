import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpFxPowerAutomateWebPartStrings';
import SpFxPowerAutomate from './components/SpFxPowerAutomate';
import { ISpFxPowerAutomateProps } from './components/ISpFxPowerAutomateProps';
//import { HttpClient } from '@microsoft/sp-http';

export interface ISpFxPowerAutomateWebPartProps {
  description: string;
}

export default class SpFxPowerAutomateWebPart extends BaseClientSideWebPart<ISpFxPowerAutomateWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ISpFxPowerAutomateProps> = React.createElement(
      SpFxPowerAutomate,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // const powerautomateURL = "https://81b07adac380e965b84fe5494a9635.dd.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/fd094997c0ff42109252ca4adcdda243/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_TyMnzAMImPsoe2bg-qdrRtMw_HV3jT-DyjkHyGg-UM";

    // try {
    //   console.log("Triggering Power Automate...");
    //   const response = await this.context.httpClient.post(
    //     powerautomateURL,
    //     HttpClient.configurations.v1,
    //     {
    //       headers: {
    //         "Content-Type": "application/json",
    //       },
    //       body: JSON.stringify({ message: "Triggered from SPFx!" }),
    //     }
    //   );

    //   if (response.ok) {
    //     console.log(
    //       "Power Automate triggered successfully:",
    //       await response.json()
    //     );
    //   } else {
    //     console.error(
    //       "Failed to trigger Power Automate:",
    //       response.statusText,
    //       await response.text()
    //     );
    //   }
    // } catch (error) {
    //   console.error("Error triggering Power Automate:", error);
    // }
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
