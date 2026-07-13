import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import HandoverManagment from './components/HandoverManagment';
import { IHandoverManagmentProps } from './components/IHandoverManagmentProps';

export interface IHandoverManagmentWebPartProps {
  siteUrl: string;
  listName: string;
  apartmentsPerFloor: number;
  colorDefault: string;
  colorDone: string;
  colorInProgress: string;
  colorNotStarted: string;
}

export default class HandoverManagmentWebPart extends BaseClientSideWebPart<IHandoverManagmentWebPartProps> {

  public render(): void {
    const statusColors: { [status: string]: string } = {
      'הושלם': this.properties.colorDone || '#4CAF50',
      'בתהליך': this.properties.colorInProgress || '#FF9800',
      'טרם התחיל': this.properties.colorNotStarted || '#F44336',
      '_default': this.properties.colorDefault || '#9E9E9E'
    };

    const element: React.ReactElement<IHandoverManagmentProps> = React.createElement(
      HandoverManagment,
      {
        siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
        listName: this.properties.listName || 'Andrius3',
        apartmentsPerFloor: this.properties.apartmentsPerFloor || 6,
        spHttpClient: this.context.spHttpClient,
        statusColors: statusColors
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
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
            description: 'הגדרות תצוגת בניין'
          },
          groups: [
            {
              groupName: 'הגדרות כלליות',
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: 'כתובת האתר',
                  description: 'כתובת אתר המסירות (ברירת מחדל: האתר הנוכחי)',
                  value: ''
                }),
                PropertyPaneTextField('listName', {
                  label: 'שם הרשימה (בניין)',
                  description: 'שם הרשימה בSharePoint',
                  value: 'Andrius3'
                }),
                PropertyPaneSlider('apartmentsPerFloor', {
                  label: 'דירות לקומה',
                  min: 1,
                  max: 20,
                  value: 6
                })
              ]
            },
            {
              groupName: 'צבעי סטטוס',
              groupFields: [
                PropertyPaneTextField('colorDone', {
                  label: 'צבע - הושלם',
                  value: '#4CAF50'
                }),
                PropertyPaneTextField('colorInProgress', {
                  label: 'צבע - בתהליך',
                  value: '#FF9800'
                }),
                PropertyPaneTextField('colorNotStarted', {
                  label: 'צבע - טרם התחיל',
                  value: '#F44336'
                }),
                PropertyPaneTextField('colorDefault', {
                  label: 'צבע ברירת מחדל',
                  value: '#9E9E9E'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
