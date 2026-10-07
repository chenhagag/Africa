import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

const SUPPLIERS_LIST = "רשימת ספקים";
const SUPPLIER_TYPES_LIST = "סוגי ספקים";

export class SupplierImportService {
  constructor(private sp: SPFI) {}

  public async deleteAllItems(onProgress: (msg: string) => void): Promise<number> {
    const list = this.sp.web.lists.getByTitle(SUPPLIERS_LIST);
    let deleted = 0;

    while (true) {
      const items: { Id: number }[] = await list.items.select("Id").top(100)();
      if (items.length === 0) break;

      for (const item of items) {
        await list.items.getById(item.Id).delete();
        deleted++;
      }
      onProgress(`נמחקו ${deleted} פריטים...`);
    }

    return deleted;
  }

  public async importData(
    data: string[][],
    onProgress: (msg: string) => void
  ): Promise<{ added: number; errors: number }> {
    // Build lookup map: SupplierType title -> Id
    onProgress("טוען סוגי ספקים...");
    const typeItems: { Id: number; Title: string }[] = await this.sp.web.lists
      .getByTitle(SUPPLIER_TYPES_LIST)
      .items.select("Id", "Title").top(5000)();

    const typeMap: Record<string, number> = {};
    for (const t of typeItems) {
      typeMap[t.Title] = t.Id;
    }
    onProgress(`נטענו ${typeItems.length} סוגי ספקים. בודק שדות...`);

    const list = this.sp.web.lists.getByTitle(SUPPLIERS_LIST);

    onProgress(`מתחיל ייבוא (${data.length} שורות)...`);
    let added = 0;
    let errors = 0;

    for (const row of data) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const values: Record<string, any> = { Title: row[0] };
      try {
        if (row[1]) values.Address = row[1];
        if (row[2] && typeMap[row[2]]) values.SupplierTypeId = typeMap[row[2]];
        if (row[3]) values.cmtSplrCode = row[3];
        if (row[4]) values.AfricaVendorCode = row[4];

        await list.items.add(values);
        added++;

        if (added % 100 === 0) {
          onProgress(`נוספו ${added} / ${data.length}...`);
        }
      } catch (e) {
        errors++;
        console.error(`Error adding '${row[0]}':`, e, 'Values:', JSON.stringify(values));
        if (errors <= 3) {
          onProgress(`שגיאה #${errors} ב-'${row[0]}': ${e}`);
        }
      }
    }

    return { added, errors };
  }
}
