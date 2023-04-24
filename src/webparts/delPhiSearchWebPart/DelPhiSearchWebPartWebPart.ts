import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DelPhiSearchWebPartWebPartStrings';
import DelPhiSearchWebPart from './components/DelPhiSearchWebPart';
import { IDelPhiSearchWebPartProps } from './components/IDelPhiSearchWebPartProps';
import { getGraph, getSP } from '../../pnpjsConfig';

export interface IDelPhiSearchWebPartWebPartProps {
  description: string;
}

export default class DelPhiSearchWebPartWebPart extends BaseClientSideWebPart<IDelPhiSearchWebPartWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDelPhiSearchWebPartProps> = React.createElement(
      DelPhiSearchWebPart,
      {
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
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
    /**
  * Initialize the web part.
  */
    public async onInit(): Promise<void> {
      await super.onInit();
  
      //Initialize our _sp object that we can then use in other packages without having to pass around the context.
      // Check out pnpjsConfig.ts for an example of a project setup file.
      getSP(this.context);
      getGraph(this.context);
    }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'DelPhi Search WebPart'
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
