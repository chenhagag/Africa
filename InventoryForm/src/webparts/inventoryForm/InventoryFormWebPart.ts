import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'InventoryFormWebPartStrings';
import InventoryForm from './components/InventoryForm';
import { IInventoryFormProps } from './components/IInventoryFormProps';

import { sp } from "@pnp/sp/presets/all";

import { MSGraphClientV3 } from '@microsoft/sp-http';


export default class InventoryFormWebPart extends BaseClientSideWebPart<IInventoryFormProps> {

  public render(): void {
    const element: React.ReactElement<IInventoryFormProps> = React.createElement(
      InventoryForm,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }


  public async onInit(): Promise<void> {
    await super.onInit();

    await this.context.serviceScope.whenFinished(async () => {
   
      sp.setup({
        spfxContext: this.context as any
      });
  
      console.log('msGraphClientFactory:', this.context.msGraphClientFactory); // צריך להיות מלא ולא undefined

      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const response = await client
        .api('/me')
        .get();

      console.log(response); // תראי פרטים על עצמך (המשתמש הנוכחי)
    });
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
