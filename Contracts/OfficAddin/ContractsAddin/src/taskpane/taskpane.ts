import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";

const MSAL_CONFIG = {
  auth: {
    clientId: "d8f0fc93-7736-43c1-8e12-8e193f543cd4",
    authority: "https://login.microsoftonline.com/b4d149d3-3aef-42b5-a6f1-b5018284caf9",
    redirectUri: "https://localhost:3000/taskpane.html"
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true,
  }
};
const GRAPH_SCOPES = ["User.Read", "Sites.Read.All"]; 

const SP_HOSTNAME = "africaisrael.sharepoint.com";
const SP_SITE_PATH = "ContractsNEW";
const LIST_DISPLAY_NAME = "ContractStatus";
const LIST_SERVER_RELATIVE_URL = "/ContractsNEW/Lists/ContractStatus";  
const DROPDOWN_TEXT_FIELD = "Title";


async function getListId(siteId: string, token: string): Promise<string> {
  const lists = await graph<{ value: Array<{ id: string; displayName: string; webUrl?: string }> }>(
    `/sites/${siteId}/lists?$select=id,displayName,webUrl`, token
  );

  // 1) נסה לפי displayName
  let found = lists.value.find(l => l.displayName === LIST_DISPLAY_NAME);
  if (found) return found.id;

  // 2) Fallback: נסה לפי webUrl מסתיים בנתיב הרשימה
  const wanted = LIST_SERVER_RELATIVE_URL.toLowerCase();
  found = lists.value.find(l => (l.webUrl || "").toLowerCase().endsWith(wanted));
  if (found) return found.id;

  // לוג דיבוג נוח
  console.warn("lists returned:", lists.value.map(l => ({ displayName: l.displayName, webUrl: l.webUrl })));

  throw new Error(`לא נמצאה רשימה. נסי לעדכן LIST_DISPLAY_NAME או LIST_SERVER_RELATIVE_URL`);
}


// ====== MSAL BOOTSTRAP ======
const msal = new PublicClientApplication(MSAL_CONFIG);
const msalInitPromise = msal.initialize(); // <-- הוספה חשובה
let activeAccount: AccountInfo | null = null;

async function ensureLogin(): Promise<void> {
  await msalInitPromise;
  const accounts = msal.getAllAccounts();
  if (accounts.length) {
    activeAccount = accounts[0];
    msal.setActiveAccount(activeAccount);
    return;
  }

  const loginResp = await msal.loginPopup({
    prompt: "select_account",
    scopes: GRAPH_SCOPES,
  });
  activeAccount = loginResp.account!;
  msal.setActiveAccount(activeAccount);
}



async function getGraphToken(): Promise<string> {
  await msalInitPromise;
  if (!activeAccount) await ensureLogin();

  try {
    const res = await msal.acquireTokenSilent({
      account: activeAccount!,
      scopes: GRAPH_SCOPES,
    });
    return res.accessToken;
  } catch {
    const res = await msal.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    return res.accessToken;
  }
}

// ====== GRAPH HELPERS ======
async function graph<T>(url: string, token: string): Promise<T> {
  const resp = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Graph ${resp.status}: ${text}`);
  }
  return resp.json() as Promise<T>;
}

// /sites/{hostname}:/{sitePath}
async function getSiteId(token: string): Promise<string> {
  const url = `/sites/${SP_HOSTNAME}:/${SP_SITE_PATH}`;
  console.log("Graph getSiteId URL:", url);
  const data = await graph<{ id: string }>(url, token);
  console.log("Resolved siteId:", data.id);
  return data.id;
}


// read items (expand fields to get Title/etc.)
async function getListItems(siteId: string, listId: string, token: string): Promise<string[]> {
  const res = await graph<{ value: Array<{ id: string; fields: Record<string, any> }> }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields($select=${DROPDOWN_TEXT_FIELD})`, token
  );
  const values = res.value
    .map(it => (it.fields?.[DROPDOWN_TEXT_FIELD] ?? "").toString().trim())
    .filter(Boolean);

  // ייחוד ומיון נעים
  return Array.from(new Set(values)).sort((a, b) => a.localeCompare(b, "he"));
}

// ====== UI ======
function fillSelect(options: string[]) {
  const sel = document.getElementById("recipientSelect") as HTMLSelectElement | null;
  if (!sel) return;
  sel.innerHTML = ""; // נקה
  sel.append(new Option("— בחר/י —", ""));
  options.forEach(v => sel.append(new Option(v, v)));
  // כאשר בוחרים בדרופדאון, נשים את זה גם בתיבת הטקסט כדי שתראי מה ייכתב למסמך
  sel.onchange = () => {
    const input = document.getElementById("recipientInput") as HTMLInputElement | null;
    if (input) input.value = sel.value || "";
  };
}

async function loadDropdownFromSharePoint() {
  try {
    const token = await getGraphToken();
    const siteId = await getSiteId(token);
    const listId = await getListId(siteId, token);
    const values = await getListItems(siteId, listId, token);
    fillSelect(values);
  } catch (e: any) {
    console.error("שגיאה בטעינת הדרופדאון:", e);
    const sel = document.getElementById("recipientSelect") as HTMLSelectElement | null;
    if (sel) {
      sel.innerHTML = "";
      sel.append(new Option("שגיאה בטעינה — רענן/י או בדקי הרשאות", ""));
    }
    alert("לא הצלחתי לטעון נתונים מהרשימה. בדקי הרשאות/Client ID בקוד.");
  }
}

// ====== WORD RUN (לפי הקוד שלך, עם עדכון קטן) ======
export async function run() {
  const input = document.getElementById("recipientInput") as HTMLInputElement | null;
  const recipient = (input?.value || "").trim();

  if (!recipient) {
    input?.focus();
    alert("אנא בחר/י מהדרופדאון או הקלד/י נמען");
    return;
  }

  try {
    await Word.run(async (context) => {
      const doc = context.document;
      const byTagEn = doc.contentControls.getByTag("recipient");
      byTagEn.load("items");
      await context.sync();

      const target = byTagEn.items[0];
      if (!target) {
        // אם אין, ניצור אחד במקום הסמן הנוכחי (או תחליפי ללוגיקה שלך)
        const range = doc.getSelection();
        const cc = range.insertContentControl();
        cc.tag = "recipient";
        cc.title = "recipient";
        cc.insertText(recipient, Word.InsertLocation.replace);
      } else {
        target.insertText(recipient, Word.InsertLocation.replace);
      }
      await context.sync();
    });

    console.log("עודכן בהצלחה");
  } catch (e: any) {
    console.error("Error in run:", e);
    alert("שגיאה: " + (e?.message || "לא ידועה"));
  }
}

// ====== OFFICE READY ======
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    (document.getElementById("sideload-msg") as HTMLElement).style.display = "none";
    (document.getElementById("app-body") as HTMLElement).style.display = "flex";

    const runButton = document.getElementById("run");
    if (runButton) runButton.onclick = run;

    // טען את הדרופדאון ברקע
    loadDropdownFromSharePoint();
  } else {
    (document.getElementById("sideload-msg") as HTMLElement).style.display = "block";
    (document.getElementById("app-body") as HTMLElement).style.display = "none";
  }
});



