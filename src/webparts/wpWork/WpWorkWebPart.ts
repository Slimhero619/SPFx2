import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import WpWork from './components/WpWork';
import { IWpWorkProps } from './components/IWpWorkProps';

export interface IWpWorkWebPartProps {
  // No web part properties needed for this Single Part App Page
  // All configuration can be managed through the Settings page
}

export default class WpWorkWebPart extends BaseClientSideWebPart<IWpWorkWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IWpWorkProps> = React.createElement(
      WpWork,
      {
        context: this.context,
        isDarkTheme: this._isDarkTheme,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
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
            description: 'Single Part App Page Configuration'
          },
          groups: [
            {
              groupName: 'General Settings',
              groupFields: [
                // Configuration fields can be added here as needed
                // For now, this Single Part App Page is self-contained
              ]
            }
          ]
        }
      ]
    };
  }
}