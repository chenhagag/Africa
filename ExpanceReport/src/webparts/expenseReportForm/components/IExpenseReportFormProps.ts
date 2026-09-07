import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IExpenseReportFormProps {
  context: WebPartContext;
  listName: string;
  siteUrl: string;
}
