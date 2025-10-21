import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'VisionMissionWebPartStrings';
import VisionMission from './components/VisionMission';
import { IVisionMissionProps } from './components/IVisionMissionProps';

export interface IVisionMissionWebPartProps {
  description: string;
  missiontitle: string;
  missiondescription: string;
  visiontitle: string;
  visiondescription: string;
}

export default class VisionMissionWebPart extends BaseClientSideWebPart<IVisionMissionWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IVisionMissionProps> = React.createElement(
      VisionMission,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        missiontitle: this.properties.missiontitle ? this.properties.missiontitle : "Mission",
        missiondescription: this.properties.missiondescription ? this.properties.missiondescription : "Empowering the world to preserve and grow their wealth through informed investing and seamless access to physical precious metals and other tangible assets.",
        visiontitle: this.properties.visiontitle ? this.properties.visiontitle : "Vision",
        visiondescription: this.properties.visiondescription ? this.properties.visiondescription : "To be the most trusted brand in physical precious metals investment, with an unwavering focus on superior products, comprehensive education, and exceptional service.",
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
          groups: [
            {
              groupName: "Vision & Mission",
              groupFields: [
                PropertyPaneTextField('visiontitle', {
                  label: "Vision Title"
                }),
                PropertyPaneTextField('visiondescription', {
                  label: "Vision Description",
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneTextField('missiontitle', {
                  label: "Mission Title"
                }),
                PropertyPaneTextField('missiondescription', {
                  label: "Mission Description",
                  multiline: true,
                  rows: 3
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
