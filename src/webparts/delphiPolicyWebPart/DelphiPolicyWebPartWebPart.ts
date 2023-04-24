import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import DelphiPolicyWebPart from './components/DelphiPolicyWebPart';
import { IDelphiPolicyWebPartProps } from './components/IDelphiPolicyWebPartProps';
import { getSP } from '../../pnpjsConfig';

export interface IDelphiPolicyWebPartWebPartProps {
  listName: string;
}

export default class DelphiPolicyWebPartWebPart extends BaseClientSideWebPart<IDelphiPolicyWebPartWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDelphiPolicyWebPartProps> = React.createElement(
      DelphiPolicyWebPart,
      {
        listName: this.properties.listName,
        userDisplayName: this.context.pageContext.user.displayName
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
    }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Delphi Policy WebPart'
          },
          groups: [
            {
              groupName: 'Policies',
              groupFields: [
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
