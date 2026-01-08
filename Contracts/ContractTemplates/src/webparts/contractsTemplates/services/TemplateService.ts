import { SPFI } from "@pnp/sp";
import { ITemplateType, ITemplateLink } from "../models/models";

export class TemplateService {
  constructor(
    private sp: SPFI,
    private typesListTitle: string,
    private linksListTitle: string
  ) {}

  public async getTypes(): Promise<ITemplateType[]> {
    const items = await this.sp.web.lists
      .getByTitle(this.typesListTitle)
      .items.select("Id", "Title", "SortOrder")
      .orderBy("SortOrder", true)
      .orderBy("Title", true)();

    return items as ITemplateType[];
  }

  public async getAllTemplates(): Promise<ITemplateLink[]> {
    const items: any[] = await this.sp.web.lists
      .getByTitle(this.linksListTitle)
      .items
      .select(
        "Id",
        "Title",
        "SortOrder",
        "Description",
        "TemplateType/Id",
        "TemplateType/Title",
        "CreateUrl"
      )
      .expand("TemplateType")
      .orderBy("TemplateType/Title", true)
      .orderBy("SortOrder", true)
      .orderBy("Title", true)();
  
    return items.map(i => ({
      Id: i.Id,
      Title: i.Title,
      Description: i.Description,
      SortOrder: i.SortOrder,
      TemplateTypeId: i.TemplateType?.Id,
      TemplateTypeTitle: i.TemplateType?.Title,
      CreateUrl: (i.CreateUrl?.Url || i.CreateUrl || "").toString()
    })) as ITemplateLink[];
  }
  

  public async getTemplatesByTypeId(typeId: number): Promise<ITemplateLink[]> {
   
    const items: any[] = await this.sp.web.lists
      .getByTitle(this.linksListTitle)
      .items
      .select(
        "Id",
        "Title",
        "SortOrder",
        "Description",
        "TemplateType/Id",
        "TemplateType/Title",
        "CreateUrl"
      )
      .expand("TemplateType")
      .filter(`TemplateType/Id eq ${typeId}`)
      .orderBy("SortOrder", true)
      .orderBy("Title", true)();

    return items.map(i => ({
      Id: i.Id,
      Title: i.Title,
      Description: i.Description,
      SortOrder: i.SortOrder,

      TemplateTypeId: i.TemplateType?.Id,
      TemplateTypeTitle: i.TemplateType?.Title,

      CreateUrl: (i.CreateUrl?.Url || i.CreateUrl || "").toString()
    })) as ITemplateLink[];
  }
}
