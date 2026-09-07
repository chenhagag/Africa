import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import ExpenseReportForm from './components/ExpenseReportForm';
import { IExpenseReportFormProps } from './components/IExpenseReportFormProps';

import { sp } from "@pnp/sp/presets/all";

export interface IExpenseReportFormWebPartProps {
  listName: string;
  siteUrl: string;
}

export default class ExpenseReportFormWebPart extends BaseClientSideWebPart<IExpenseReportFormWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IExpenseReportFormProps> = React.createElement(
      ExpenseReportForm,
      {
        context: this.context,
        listName: this.properties.listName || 'ExpenseReport',
        siteUrl: this.properties.siteUrl || 'https://africaisrael.sharepoint.com/ExpenseReports'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public async onInit(): Promise<void> {
    await super.onInit();

    sp.setup({
      spfxContext: this.context as any
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
            description: 'הגדרות טופס החזר הוצאות'
          },
          groups: [
            {
              groupName: 'הגדרות',
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'שם הרשימה'
                }),
                PropertyPaneTextField('siteUrl', {
                  label: 'כתובת האתר (אם שונה מהאתר הנוכחי)'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
