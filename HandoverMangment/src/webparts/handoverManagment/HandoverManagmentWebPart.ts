import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import HandoverManagment from './components/HandoverManagment';
import { IHandoverManagmentProps } from './components/IHandoverManagmentProps';

export interface IHandoverManagmentWebPartProps {
  siteUrl: string;
  listName: string;
  // Sale status colors
  colorSaleAvailable: string;
  colorSaleSold: string;
  colorSaleOwner: string;
  colorSaleHold: string;
  colorSaleDefault: string;
  // Finish status colors
  colorFinishMinimal: string;
  colorFinishStandard: string;
  colorFinishUpgrade: string;
  colorFinishMarketing: string;
  // Public unit type colors
  colorPublicBuilding: string;
  colorCommercial: string;
  colorWelfare: string;
  colorTechnical: string;
  colorPool: string;
}

export default class HandoverManagmentWebPart extends BaseClientSideWebPart<IHandoverManagmentWebPartProps> {

  public render(): void {
    const saleColors: { [status: string]: string } = {
      'פנויה': this.properties.colorSaleAvailable || '#9E9E9E',
      'מכורה': this.properties.colorSaleSold || '#2196F3',
      'בעלים': this.properties.colorSaleOwner || '#4CAF50',
      'הולד': this.properties.colorSaleHold || '#FF9800',
      '_default': this.properties.colorSaleDefault || '#BDBDBD'
    };

    const finishColors: { [status: string]: string } = {
      'גמר מינימלי': this.properties.colorFinishMinimal || '#FFEB3B',
      'גמר מלא סטנדרט': this.properties.colorFinishStandard || '#FE8B34',
      'גמר מלא טיוב תכנון': this.properties.colorFinishUpgrade || '#66FF66',
      'דירת שיווק': this.properties.colorFinishMarketing || '#00BCD4'
    };

    const publicUnitColors: { [type: string]: string } = {
      'מבנה ציבורי': this.properties.colorPublicBuilding || '#66BB6A',
      'מסחר': this.properties.colorCommercial || '#B0BEC5',
      'רווחה': this.properties.colorWelfare || '#FFC1EA',
      'אזור טכני': this.properties.colorTechnical || '#7989BD',
      'בריכה': this.properties.colorPool || '#81D4FA'
    };

    const element: React.ReactElement<IHandoverManagmentProps> = React.createElement(
      HandoverManagment,
      {
        siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
        listName: this.properties.listName || 'andreius3',
        spHttpClient: this.context.spHttpClient,
        saleColors: saleColors,
        finishColors: finishColors,
        publicUnitColors: publicUnitColors
      }
    );

    ReactDom.render(element, this.domElement);

    // Expand canvas zone to allow wider layout
    const container = this.domElement.closest('.CanvasZoneSectionContainer') as HTMLElement;
    if (container) {
      container.style.maxWidth = '1700px';
    }
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
                  value: 'andreius3'
                })
              ]
            },
            {
              groupName: 'צבעי סטטוס מכירה',
              groupFields: [
                PropertyFieldColorPicker('colorSaleAvailable', {
                  label: 'פנויה',
                  selectedColor: this.properties.colorSaleAvailable || '#9E9E9E',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorSaleAvailable'
                }),
                PropertyFieldColorPicker('colorSaleSold', {
                  label: 'מכורה',
                  selectedColor: this.properties.colorSaleSold || '#2196F3',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorSaleSold'
                }),
                PropertyFieldColorPicker('colorSaleOwner', {
                  label: 'בעלים',
                  selectedColor: this.properties.colorSaleOwner || '#4CAF50',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorSaleOwner'
                }),
                PropertyFieldColorPicker('colorSaleHold', {
                  label: 'הולד',
                  selectedColor: this.properties.colorSaleHold || '#FF9800',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorSaleHold'
                }),
                PropertyFieldColorPicker('colorSaleDefault', {
                  label: 'ברירת מחדל',
                  selectedColor: this.properties.colorSaleDefault || '#BDBDBD',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorSaleDefault'
                })
              ]
            },
            {
              groupName: 'צבעי סטטוס גמר',
              groupFields: [
                PropertyFieldColorPicker('colorFinishMinimal', {
                  label: 'גמר מינימלי',
                  selectedColor: this.properties.colorFinishMinimal || '#FFEB3B',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorFinishMinimal'
                }),
                PropertyFieldColorPicker('colorFinishStandard', {
                  label: 'גמר מלא סטנדרט',
                  selectedColor: this.properties.colorFinishStandard || '#FE8B34',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorFinishStandard'
                }),
                PropertyFieldColorPicker('colorFinishUpgrade', {
                  label: 'גמר מלא טיוב תכנון',
                  selectedColor: this.properties.colorFinishUpgrade || '#66FF66',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorFinishUpgrade'
                }),
                PropertyFieldColorPicker('colorFinishMarketing', {
                  label: 'דירת שיווק',
                  selectedColor: this.properties.colorFinishMarketing || '#00BCD4',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorFinishMarketing'
                })
              ]
            },
            {
              groupName: 'צבעי מבנים ציבוריים',
              groupFields: [
                PropertyFieldColorPicker('colorPublicBuilding', {
                  label: 'מבנה ציבורי',
                  selectedColor: this.properties.colorPublicBuilding || '#66BB6A',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorPublicBuilding'
                }),
                PropertyFieldColorPicker('colorCommercial', {
                  label: 'מסחר',
                  selectedColor: this.properties.colorCommercial || '#B0BEC5',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorCommercial'
                }),
                PropertyFieldColorPicker('colorWelfare', {
                  label: 'רווחה',
                  selectedColor: this.properties.colorWelfare || '#FFC1EA',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorWelfare'
                }),
                PropertyFieldColorPicker('colorTechnical', {
                  label: 'אזור טכני',
                  selectedColor: this.properties.colorTechnical || '#7989BD',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorTechnical'
                }),
                PropertyFieldColorPicker('colorPool', {
                  label: 'בריכה',
                  selectedColor: this.properties.colorPool || '#81D4FA',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'colorPool'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
