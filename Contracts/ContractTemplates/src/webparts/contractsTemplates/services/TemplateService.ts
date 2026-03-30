import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { ITemplateType, ITemplateLink } from "../models/models";

const TEMPLATE_URL_FIELD = "TemplateServerRelativeUrl";

export class TemplateService {
  constructor(
    private sp: SPFI,
    private typesListTitle: string,
    private linksListTitle: string
  ) {}

  public async getTypes(): Promise<ITemplateType[]> {

    debugger;

    const items = await this.sp.web.lists
      .getByTitle(this.typesListTitle)
      .items.select("Id", "Title", "SortOrder")
      .orderBy("SortOrder", true)
      .orderBy("Title", true)();

    return items as ITemplateType[];
  }

  public async getAllTemplates(): Promise<ITemplateLink[]> {

    debugger;

    const selectFields = [
      "Id",
      "Title",
      "Description",
      "SortOrder",
      TEMPLATE_URL_FIELD,
      "TemplateType/Id",
      "TemplateType/Title"
    ].join(",");

    const items: any[] = await this.sp.web.lists
      .getByTitle(this.linksListTitle)
      .items
      .select(selectFields)
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
      TemplateServerRelativeUrl: (i[TEMPLATE_URL_FIELD] || "").toString()
    })) as ITemplateLink[];
  }

  public async createDocFromTemplateBlob(params: {
    templateServerRelativeUrl: string;
    targetFolderServerRelativeUrl: string; 
    newFileName: string;                   
  }): Promise<{ newFileServerRelativeUrl: string }> {
    const { templateServerRelativeUrl, targetFolderServerRelativeUrl, newFileName } = params;

    // 1) download template as blob
    const blob = await this.sp.web
      .getFileByServerRelativePath(templateServerRelativeUrl)
      .getBlob();

    const targetFileServerRelativeUrl = `${targetFolderServerRelativeUrl}/${newFileName}`;

    // 2) upload: try addUsingPath if exists in this PnP version
    const folderAny: any = this.sp.web.getFolderByServerRelativePath(targetFolderServerRelativeUrl) as any;

    // PnP versions differ; some expose folder.files.addUsingPath(...)
    if (folderAny?.files?.addUsingPath) {
      const res: any = await folderAny.files.addUsingPath(newFileName, blob, { Overwrite: false });
      const url = res?.data?.ServerRelativeUrl || targetFileServerRelativeUrl;
      return { newFileServerRelativeUrl: url };
    }

    const webUrl = (this.sp.web as any).toUrl ? (this.sp.web as any).toUrl() : "";
    if (!webUrl) {
    }

    const siteBase =
      webUrl ||
      `${window.location.origin}${window.location.pathname.split("/").slice(0, 2).join("/")}`;

    const endpoint =
      `${siteBase}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(targetFolderServerRelativeUrl).replace(/%2F/g, "/")}')/Files/add(url='${encodeURIComponent(newFileName)}',overwrite=false)`;

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Accept": "application/json;odata=verbose",
        // Digest חובה ל-POST ב-SharePoint
        "X-RequestDigest": (document.getElementById("__REQUESTDIGEST") as HTMLInputElement)?.value || ""
      },
      body: blob
    });

    if (!response.ok) {
      const txt = await response.text();
      throw new Error(`Upload failed (${response.status}). ${txt}`);
    }

    return { newFileServerRelativeUrl: targetFileServerRelativeUrl };
  }
}
