import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AfricaDeliveryNewFormWebPartStrings';
import AfricaDeliveryNewForm from './components/AfricaDeliveryNewForm';
import { IAfricaDeliveryNewFormProps } from './components/IAfricaDeliveryNewFormProps';

export interface IAfricaDeliveryNewFormWebPartProps {
  description: string;
}

export default class AfricaDeliveryNewFormWebPart extends BaseClientSideWebPart<IAfricaDeliveryNewFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAfricaDeliveryNewFormProps> = React.createElement(
      AfricaDeliveryNewForm,
      {
        description: this.properties.description,
        context: this.context  
      }
    );
    
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.resolve(); // תיקון ל-onInit
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
