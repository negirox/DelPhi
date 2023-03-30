import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import DelphiBanner from './components/DelphiBanner';
import { IDelphiBannerProps } from './components/IDelphiBannerProps';
import { HelperUtils } from '../../utils/HelperUtils';

export interface IDelphiBannerWebPartProps {
  headerText:string;
  description: string;
  bannerDescription:string;
  bannerImageUrl :string;
}

export default class DelphiBannerWebPart extends BaseClientSideWebPart<IDelphiBannerWebPartProps> {
  private _isDarkTheme: boolean = false;
  public render(): void {
    const element: React.ReactElement<IDelphiBannerProps> = React.createElement(
      DelphiBanner,
      {
        description: this.properties.bannerDescription,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        headerText: this.properties.headerText,
        bannerImageUrl: this.properties.bannerImageUrl,
        isDarkTheme: this._isDarkTheme
      }
    );
    if(!HelperUtils.IsNullOrEmpty(this.properties.headerText) && !HelperUtils.IsNullOrEmpty(this.properties.bannerImageUrl)) 
    {
     ReactDom.render(element, this.domElement);
    }
    else
    {
     const myElement = React.createElement('h1', {}, 'Please Configure the property.');
     ReactDom.render(myElement,this.domElement);
    }
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
          groups: [
            {
              groupName: 'Banner Configurations',
              groupFields: [
                PropertyPaneTextField('headerText', {
                  label: 'Banner heading'
                }),
                PropertyPaneTextField('bannerDescription', {
                  label: 'Banner Description'
                }),
                PropertyPaneTextField('bannerImageUrl', {
                  label: 'Banner Image'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
