import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import TemplatePicker, { ITemplatePickerProps } from './components/ContractsTemplates';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface ITemplatePickerWebPartProps {
  templateTypesListTitle: string;
  templateLinksListTitle: string;
}

export default class TemplatePickerWebPart extends BaseClientSideWebPart<ITemplatePickerWebPartProps> {

  public render(): void {
    const sp = spfi().using(SPFx(this.context));

    const element: React.ReactElement<ITemplatePickerProps> = React.createElement(
      TemplatePicker,
      {
        sp,
        templateTypesListTitle: this.properties.templateTypesListTitle || "TemplateTypes",
        templateLinksListTitle: this.properties.templateLinksListTitle || "TemplateLinks",
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
}
