import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { InjectHeaders } from '@pnp/queryable';
import { setupSP } from 'sp-preset';
import { SPnotify } from 'sp-react-notifications';

import * as strings from 'CipAggregatorWebPartStrings';
import CipAggregator from './components/CipAggregator';
import { ICipAggregatorProps } from './components/ICipAggregatorProps';

export interface ICipAggregatorWebPartProps {
  description: string;
}

export default class CipAggregatorWebPart extends BaseClientSideWebPart<ICipAggregatorWebPartProps> {

  private _isDarkTheme: boolean = false;

  protected async onInit(): Promise<void> {
    setupSP({
      context: this.context,
      useRPM: true,
      rpmTreshold: 800,
      rpmTracing: false,
      rpmAlerting: true,
      additionalTimelinePipes: [
        InjectHeaders({
            "Accept": "application/json;odata=nometadata"
        }),
      ],
    });

    SPnotify({
      message: 'Hello world'
    });

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ICipAggregatorProps> = React.createElement(
      CipAggregator,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: "",
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
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
