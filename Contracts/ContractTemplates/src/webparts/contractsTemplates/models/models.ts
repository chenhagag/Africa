export interface ITemplateType {
    Id: number;
    Title: string;
    SortOrder?: number;
  }
  
  export interface ITemplateLink {
    Id: number;
    Title: string;
    CreateUrl: string;
    Description?: string;
    SortOrder?: number;
  
    TemplateTypeId: number;
    TemplateTypeTitle?: string;
  }
  