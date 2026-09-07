import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IExpenseMigrationProps {
  context: WebPartContext;
  sourceLibrary: string;
  sourceSiteUrl: string;
  targetList: string;
  targetSiteUrl: string;
}
