import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

// Fields to update (only if empty in existing item)
const UPDATE_FIELDS = ['msrAppStatus'];

// Data array index mapping: [msrAppNum, Tower, Floor, ShadowNumber, ApartmentType, Rooms, Standard, HandoverStatus, FinishStatus, msrAppStatus]
const FIELD_INDEX: Record<string, number> = {
  Tower: 1,
  Floor: 2,
  ShadowNumber: 3,
  ApartmentType: 4,
  Rooms: 5,
  Standard: 6,
  HandoverStatus: 7,
  FinishStatus: 8,
  msrAppStatus: 9
};

export interface IUpdateResult {
  updated: number;
  skipped: number;
  notFound: number;
  duplicates: number;
  errors: number;
  log: string[];
}

export class ApartmentUpdateService {
  private _spHttpClient: SPHttpClient;
  private _siteUrl: string;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  /**
   * Fetch all items from the list (handles pagination for large lists)
   */
  private async _getAllItems(listName: string): Promise<Array<Record<string, unknown>>> {
    const allItems: Array<Record<string, unknown>> = [];
    const fields = 'Id,msrAppNum,msrBldNum,' + UPDATE_FIELDS.join(',');
    let url: string | null = `${this._siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=${fields}&$top=5000&$orderby=Id`;

    while (url) {
      const response: SPHttpClientResponse = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      const items = data.value || [];
      allItems.push(...items);

      // Follow @odata.nextLink for pagination
      url = data['@odata.nextLink'] || data['odata.nextLink'] || null;
    }

    return allItems;
  }

  /**
   * Update a single item - only empty fields
   */
  private async _updateItem(
    listName: string,
    itemId: number,
    existingItem: Record<string, unknown>,
    dataRow: string[]
  ): Promise<{ fieldsUpdated: string[]; skippedFields: string[] }> {
    const fieldsUpdated: string[] = [];
    const skippedFields: string[] = [];
    const body: Record<string, unknown> = {};

    for (const field of UPDATE_FIELDS) {
      const existingValue = existingItem[field];
      const newValue = dataRow[FIELD_INDEX[field]];

      // Skip if no new value from CSV
      if (!newValue) {
        continue;
      }

      // Skip if existing field already has data
      if (existingValue !== null && existingValue !== undefined && existingValue !== '' && existingValue !== 0) {
        skippedFields.push(`${field}=[${typeof existingValue}]${String(existingValue)}`);
        continue;
      }

      body[field] = newValue;
      fieldsUpdated.push(field);
    }

    if (fieldsUpdated.length === 0) {
      return { fieldsUpdated, skippedFields };
    }

    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;
    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`HTTP ${response.status}: ${errText}`);
    }

    return { fieldsUpdated, skippedFields };
  }

  /**
   * Main update method: matches CSV data to list items by msrAppNum
   */
  public async updateApartments(
    listName: string,
    data: string[][],
    onProgress: (msg: string) => void
  ): Promise<IUpdateResult> {
    const result: IUpdateResult = {
      updated: 0,
      skipped: 0,
      notFound: 0,
      duplicates: 0,
      errors: 0,
      log: []
    };

    onProgress('שולף פריטים מהרשימה...');
    const allItems = await this._getAllItems(listName);
    onProgress(`נמצאו ${allItems.length} פריטים ברשימה`);

    // Build map: "msrAppNum_msrBldNum" -> item
    // msrBldNum: 1=צפוני, 2=דרומי
    const TOWER_TO_BLDNUM: Record<string, string> = { 'צפוני': '1', 'דרומי': '2' };
    const itemMap: Record<string, Record<string, unknown>> = {};
    let duplicateKeys = 0;
    allItems.forEach(item => {
      const appNum = String(item.msrAppNum || '');
      const bldNum = String(item.msrBldNum || '');
      if (appNum && bldNum) {
        const key = `${appNum}_${bldNum}`;
        if (itemMap[key]) {
          duplicateKeys++;
        } else {
          itemMap[key] = item;
        }
      }
    });
    if (duplicateKeys > 0) {
      onProgress(`אזהרה: ${duplicateKeys} מפתחות כפולים ברשימה`);
    }

    onProgress(`מתחיל עדכון ${data.length} רשומות...`);

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const msrAppNum = row[0];
      const tower = row[1];

      try {
        const bldNum = TOWER_TO_BLDNUM[tower];
        if (!bldNum) {
          result.errors++;
          result.log.push(`שגיאה: msrAppNum=${msrAppNum} — מגדל לא מוכר: ${tower}`);
          continue;
        }

        const key = `${msrAppNum}_${bldNum}`;
        const existingItem = itemMap[key];

        if (!existingItem) {
          result.notFound++;
          result.log.push(`לא נמצא: msrAppNum=${msrAppNum} (${tower}, bldNum=${bldNum})`);
          continue;
        }
        const itemId = Number(existingItem.Id);

        const { fieldsUpdated, skippedFields } = await this._updateItem(
          listName, itemId, existingItem, row
        );

        if (fieldsUpdated.length > 0) {
          result.updated++;
          result.log.push(`עודכן: msrAppNum=${msrAppNum} (${tower}) — שדות: ${fieldsUpdated.join(', ')}`);
        } else {
          result.skipped++;
          if (skippedFields.length > 0) {
            result.log.push(`דילוג: msrAppNum=${msrAppNum} (${tower}) — כל השדות מאוכלסים`);
          }
        }
      } catch (err) {
        result.errors++;
        result.log.push(`שגיאה: msrAppNum=${msrAppNum} (${tower}) — ${(err as Error).message}`);
      }

      if ((i + 1) % 50 === 0) {
        onProgress(`עודכנו ${i + 1} / ${data.length}...`);
      }
    }

    onProgress(`סיום! עודכנו: ${result.updated}, דילוג: ${result.skipped}, לא נמצאו: ${result.notFound}, כפילויות: ${result.duplicates}, שגיאות: ${result.errors}`);
    return result;
  }
}
