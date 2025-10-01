import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const LOG_SOURCE: string = 'InventoryMsApplicationCustomizer';

export interface IInventoryMsApplicationCustomizerProperties {
  adminGroupName?: string; // למשל: "Inventory Admins"
}

export default class InventoryMsApplicationCustomizer
  extends BaseApplicationCustomizer<IInventoryMsApplicationCustomizerProperties> {

  private _observer: MutationObserver | null = null;
  private _isAdmin: boolean = false;

  // ======= הגדרות להתאמה =======
  private adminGroupID = 615;

  private blockedUrlFragments: string[] = [
    '/_layouts/15/viewlsts.aspx',        
    '/_layouts/15/RecycleBin.aspx', 
    '/Lists/Inventory/',                  
    '/Lists/InventoryManagement/',             
    '/Lists/Signatures/',
    '/Lists/SiteAffiliation/',
    '/Lists/inventoryType/',                     
    '/Lists/CRM/'
  ];

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${LOG_SOURCE}`);

    this._isAdmin = await this._isUserInGroup(this.adminGroupID);

    debugger;
    
    if (!this._isAdmin) {
      // 2) חסימת ניווט ישיר לעמודים אסורים
      this._blockDirectAccess();

      // 3) הזרקת CSS להסתרת “תוכן האתר” + קישורים לרשימות
      this._injectCss();

      // 4) מאזין לשינויים ב‑DOM כדי להסיר קישורים שנוספים דינמית
      this._attachMutationObserver();

      // 5) ניסיון מיידי להסיר קישורים קיימים
      this._hideRestrictedLinks();
    }

    return Promise.resolve();
  }

// groupId לדוגמה: 615
private async _isUserInGroup(groupId: number): Promise<boolean> {
  try {
    debugger;
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser/groups?$select=Id`;
    const res: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const json: any = await res.json();
    debugger;

    const groups: Array<{ Id: number }> = json?.value || [];
    return groups.some(g => g.Id === groupId);
  } catch (e) {
    Log.warn(LOG_SOURCE, `Failed to resolve user groups: ${e}`);
    return false;
  }
}



  // ===== חסימת כניסה ישירה לעמודים אסורים =====
  private _blockDirectAccess(): void {
    const href = (window.location && window.location.href) ? window.location.href.toLowerCase() : '';
    if (!href) return;

    for (const f of this.blockedUrlFragments) {
      const frag = f ? f.toLowerCase() : '';
      if (frag && href.indexOf(frag) !== -1) {
        window.location.replace(this.context.pageContext.web.absoluteUrl);
        return;
      }
    }
  }

  // ===== CSS הסתרה גלובלי =====
  private _injectCss(): void {
    const lines: string[] = [];

    // הסתרת "תוכן האתר" לפי href
    lines.push(`a[href*="/_layouts/15/viewlsts.aspx"] { display: none !important; }`);

    // הסתרת “תוכן האתר” לפי aria-label נפוץ
    lines.push(`a[aria-label="תוכן האתר"],
a[aria-label="Site contents"] { display: none !important; }`);

    // הסתרת אריח/קיצור ל"תוכן האתר" בעמוד הבית אם קיים
    lines.push(`div[data-automation-id="siteContents"] { display: none !important; }`);

    // הסתרת קישורי רשימות ספציפיות בניווט הצדדי/העליון
    for (const f of this.blockedUrlFragments) {
      const fragLower = f ? f.toLowerCase() : '';
      // דלגי על ה‑viewlsts.aspx שכבר כוסה
      if (fragLower && fragLower.indexOf('viewlsts.aspx') === -1) {
        // שימוש ב‑href*= כדי להסתיר קישורים המכילים את התבנית
        lines.push(`a[href*="${f}"] { display: none !important; }`);
      }
    }

    const css = lines.join('\n');

    const style = document.createElement('style');
    style.setAttribute('data-inventory-hide-style', 'true');
    style.innerHTML = css;
    document.head.appendChild(style);
  }

  // ===== הסרה אקטיבית של קישורים בעייתיים =====
  private _hideRestrictedLinks(): void {
    const nodes = document.querySelectorAll('a[href]');
    const links = Array.prototype.slice.call(nodes) as HTMLAnchorElement[];

    for (const a of links) {
      const rawHref = a.getAttribute('href') || '';
      const href = rawHref.toLowerCase();

      // בדיקה מול כל התבניות
      let shouldHide = false;
      for (const f of this.blockedUrlFragments) {
        const frag = f ? f.toLowerCase() : '';
        if (frag && href.indexOf(frag) !== -1) {
          shouldHide = true;
          break;
        }
      }

      if (shouldHide) {
        (a.style as CSSStyleDeclaration).display = 'none';
        const li = a.closest('li');
        if (li) (li as HTMLElement).style.display = 'none';
      }
    }
  }

  // ===== מעקב אחרי רינדורים דינמיים =====
  private _attachMutationObserver(): void {
    this._observer = new MutationObserver(() => {
      // בכל שינוי ב‑DOM—ננסה שוב להחביא קישורים
      this._hideRestrictedLinks();

      // ואם המשתמש איכשהו הגיע לעמוד אסור דרך ניווט פנימי—נחזיר אותו
      this._blockDirectAccess();
    });

    this._observer.observe(document.body, {
      childList: true,
      subtree: true
    });
  }

  // נקה משאבים אם צריך
  public onDispose(): void {
    if (this._observer) {
      this._observer.disconnect();
      this._observer = null;
    }
    const style = document.querySelector('style[data-inventory-hide-style="true"]');
    if (style && style.parentElement) style.parentElement.removeChild(style);
  }
}
