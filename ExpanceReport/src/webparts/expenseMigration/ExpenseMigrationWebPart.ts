import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import ExpenseMigration from './components/ExpenseMigration';
import { IExpenseMigrationProps } from './components/IExpenseMigrationProps';

import { sp } from "@pnp/sp/presets/all";

export interface IExpenseMigrationWebPartProps {
  sourceLibrary: string;
  sourceSiteUrl: string;
  targetList: string;
  targetSiteUrl: string;
}

export default class ExpenseMigrationWebPart extends BaseClientSideWebPart<IExpenseMigrationWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IExpenseMigrationProps> = React.createElement(
      ExpenseMigration,
      {
        context: this.context,
        sourceLibrary: this.properties.sourceLibrary || 'ExpenseReport',
        sourceSiteUrl: this.properties.sourceSiteUrl || 'https://africaisrael.sharepoint.com/sites/AFIResidences',
        targetList: this.properties.targetList || 'ExpenseReport',
        targetSiteUrl: this.properties.targetSiteUrl || 'https://africaisrael.sharepoint.com/ExpenseReports'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    sp.setup({ spfxContext: this.context as any });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'הגדרות הסבה' },
        groups: [{
          groupName: 'מקור',
          groupFields: [
            PropertyPaneTextField('sourceSiteUrl', { label: 'כתובת אתר מקור' }),
            PropertyPaneTextField('sourceLibrary', { label: 'שם ספריית InfoPath' })
          ]
        }, {
          groupName: 'יעד',
          groupFields: [
            PropertyPaneTextField('targetSiteUrl', { label: 'כתובת אתר יעד' }),
            PropertyPaneTextField('targetList', { label: 'שם רשימת יעד' })
          ]
        }]
      }]
    };
  }
}
