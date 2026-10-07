import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const MEDAD_LIST = "ערכי מדד";

export class MadadImportService {
  constructor(private sp: SPFI) {}

  public async deleteAllItems(onProgress: (msg: string) => void): Promise<number> {
    const list = this.sp.web.lists.getByTitle(MEDAD_LIST);
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
    const list = this.sp.web.lists.getByTitle(MEDAD_LIST);
    let added = 0;
    let errors = 0;

    for (const row of data) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const values: Record<string, any> = { Title: row[0] };
      try {
        if (row[1]) values.MadadDate = row[1];
        if (row[2]) values.MadadValue = row[2];

        await list.items.add(values);
        added++;

        if (added % 100 === 0) {
          onProgress(`נוספו ${added} / ${data.length}...`);
        }
      } catch (e) {
        errors++;
        console.error(`Error adding medad '${row[0]}':`, e, 'Values:', JSON.stringify(values));
        if (errors <= 3) {
          onProgress(`שגיאה #${errors}: ${e}`);
        }
      }
    }

    return { added, errors };
  }
}
