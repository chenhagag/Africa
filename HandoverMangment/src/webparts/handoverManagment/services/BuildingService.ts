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

  public getApartments(listName: string, apartmentsPerFloor: number): Promise<IApartment[]> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title,msrAppNum&$top=5000&$orderby=Id`;
    return this._spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value?: Array<{ Id: number; Title: string; msrAppNum?: string | number }> }) => {
        const items = data.value || [];
        // POC: varying floor sizes pattern
        const floorPattern = [6, 4, 3, 2, 6, 4, 5, 3];
        const floorAssignments = this._assignFloors(items.length, floorPattern);

        return items.map((item: { Id: number; Title: string; msrAppNum?: string | number }, index: number) => {
          let apartmentNumber = 0;
          if (item.msrAppNum !== undefined && item.msrAppNum !== null) {
            apartmentNumber = parseInt(String(item.msrAppNum), 10) || 0;
          }
          if (apartmentNumber <= 0) {
            const extracted = this._extractApartmentNumber(String(item.Title || ''));
            apartmentNumber = extracted > 0 ? extracted : (index + 1);
          }
          return {
            id: Number(item.Id),
            title: String(item.Title || ''),
            apartmentNumber: apartmentNumber,
            floor: floorAssignments[index],
            status: ''
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

  // POC: assign floors with varying sizes based on a repeating pattern
  private _assignFloors(totalItems: number, pattern: number[]): number[] {
    const assignments: number[] = [];
    let floor = 1;
    let patternIndex = 0;
    let count = 0;
    for (let i = 0; i < totalItems; i++) {
      assignments.push(floor);
      count++;
      if (count >= pattern[patternIndex % pattern.length]) {
        floor++;
        patternIndex++;
        count = 0;
      }
    }
    return assignments;
  }

  private _extractApartmentNumber(title: string): number {
    const match = title.match(/\d+/);
    return match ? parseInt(match[0], 10) : 0;
  }
}
