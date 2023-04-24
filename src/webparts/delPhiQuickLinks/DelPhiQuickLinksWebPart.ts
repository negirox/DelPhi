import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import DelPhiQuickLinks from './components/DelPhiQuickLinks';
import { IDelPhiQuickLinksProps } from './components/IDelPhiQuickLinksProps';
import { getSP } from '../../pnpjsConfig';
export interface IDelPhiQuickLinksWebPartProps {
  tilesImageUrl: string;
  listName:string;
}

export default class DelPhiQuickLinksWebPart extends BaseClientSideWebPart<IDelPhiQuickLinksWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDelPhiQuickLinksProps> = React.createElement(
      DelPhiQuickLinks,
      {
        tilesImageUrl: this.properties.tilesImageUrl,
        context: this.context,
        listName: this.properties.listName ?? 'Quick Links'
      }
    );

    ReactDom.render(element, this.domElement);
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
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'DelPhi Quick Links'
          },
          groups: [
            {
              groupName: 'Quick Links',
              groupFields: [
                PropertyPaneTextField('tilesImageUrl', {
                  label: 'Enter tiles backgroun image url'
                }),
                PropertyPaneTextField('listName', {
                  label: 'Enter List Name'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
