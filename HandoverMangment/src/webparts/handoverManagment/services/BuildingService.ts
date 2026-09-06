import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IApartment } from '../components/IHandoverManagmentProps';

export interface IListInfo {
  Title: string;
  Id: string;
  RootFolder: { ServerRelativeUrl: string };
}

export class BuildingService {
  private _spHttpClient: SPHttpClient;
  private _siteUrl: string;
  private _listUrlMap: { [title: string]: string } = {};

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  public getLists(): Promise<IListInfo[]> {
    const url = `${this._siteUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Title,Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    return this._spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value?: IListInfo[] }) => {
        const lists = data.value || [];
        lists.forEach((l: IListInfo) => {
          if (l.RootFolder && l.RootFolder.ServerRelativeUrl) {
            this._listUrlMap[l.Title] = l.RootFolder.ServerRelativeUrl;
          }
        });
        return lists;
      });
  }

  public getApartments(listName: string): Promise<IApartment[]> {
    const fields = 'Id,Title,msrAppNum,Floor,Tower,msrAppStatus,HandoverStatus,FinishStatus,ApartmentType,Rooms,Standard,ShadowNumber,ContentTypeId,SpaceUnits,PublicUnitType';
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=${fields}&$top=5000&$orderby=Id`;
    return this._spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value?: Array<Record<string, unknown>> }) => {
        const items = data.value || [];

        return items.map((item: Record<string, unknown>, index: number) => {
          const spaceUnits = parseInt(String(item.SpaceUnits || '0'), 10) || 0;
          const isPublicUnit = spaceUnits > 0;

          let apartmentNumber = 0;
          if (!isPublicUnit) {
            const rawAppNum = item.msrAppNum;
            if (rawAppNum !== undefined && rawAppNum !== null) {
              apartmentNumber = parseInt(String(rawAppNum), 10) || 0;
            }
            if (apartmentNumber <= 0) {
              const extracted = this._extractApartmentNumber(String(item.Title || ''));
              apartmentNumber = extracted > 0 ? extracted : (index + 1);
            }
          }

          const floor = (item.Floor !== undefined && item.Floor !== null) ? Number(item.Floor) : 0;
          const tower = String(item.Tower || '');
          const apartmentType = String(item.ApartmentType || '');

          // Derive standard from apartment type if Standard field is empty
          let standard = String(item.Standard || '');
          if (!standard && apartmentType) {
            const upper = apartmentType.toUpperCase();
            if (upper.indexOf('PH') === 0) {
              standard = 'פנטהאוס';
            } else if (upper.charAt(upper.length - 1) === 'P') {
              standard = 'פרימיום';
            } else {
              standard = 'סטנדרט';
            }
          }

          return {
            id: Number(item.Id),
            title: String(item.Title || ''),
            apartmentNumber,
            floor,
            tower,
            saleStatus: String(item.msrAppStatus || ''),
            handoverStatus: String(item.HandoverStatus || ''),
            finishStatus: String(item.FinishStatus || ''),
            apartmentType,
            rooms: parseInt(String(item.Rooms || '0'), 10) || 0,
            standard,
            shadowNumber: parseInt(String(item.ShadowNumber || '0'), 10) || 0,
            isPublicUnit,
            spaceUnits: isPublicUnit ? spaceUnits : 1,
            publicUnitType: String(item.PublicUnitType || '')
          };
        });
      });
  }

  public getItemFormUrl(listName: string, itemId: number): string {
    const listRelUrl = this._listUrlMap[listName];
    if (listRelUrl) {
      return `${window.location.origin}${listRelUrl}/DispForm.aspx?ID=${itemId}`;
    }
    return `${this._siteUrl}/Lists/${listName}/DispForm.aspx?ID=${itemId}`;
  }

  private _extractApartmentNumber(title: string): number {
    const match = title.match(/\d+/);
    return match ? parseInt(match[0], 10) : 0;
  }
}
