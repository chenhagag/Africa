// taskpane.ts
import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";

/* =========
   Config
   ========= */
const SP_HOSTNAME = "knowedge.sharepoint.com";
const SP_SITE_PATH = "Knowedge";
const SITE_IS_UNDER_SITES = true;

// Status dropdown source
const STATUS_LIST_DISPLAY_NAME = "ContractStatus";
const STATUS_FIELD_NAME = "Title";
const STATUS_LIST_SERVER_RELATIVE_URL = SITE_IS_UNDER_SITES
  ? `/sites/${SP_SITE_PATH}/Lists/${STATUS_LIST_DISPLAY_NAME}`
  : `/${SP_SITE_PATH}/Lists/${STATUS_LIST_DISPLAY_NAME}`;

// Project dropdown source
const PROJECTS_LIST_DISPLAY_NAME = "projects";
const PROJECTS_FIELD_NAME = "Title";
const PROJECTS_LIST_SERVER_RELATIVE_URL = SITE_IS_UNDER_SITES
  ? `/sites/${SP_SITE_PATH}/Lists/${PROJECTS_LIST_DISPLAY_NAME}`
  : `/${SP_SITE_PATH}/Lists/${PROJECTS_LIST_DISPLAY_NAME}`;

// Helper list target
const HELPER_LIST_DISPLAY_NAME = "FieldsUpdateHelper";
const HELPER_LIST_SERVER_RELATIVE_URL = SITE_IS_UNDER_SITES
  ? `/sites/${SP_SITE_PATH}/Lists/${HELPER_LIST_DISPLAY_NAME}`
  : `/${SP_SITE_PATH}/Lists/${HELPER_LIST_DISPLAY_NAME}`;

const DEFAULT_REDIRECT_URI = "https://knowedge.co.il/matrix/downloads/taskpane.html";

const runtimeRedirectUri =
  typeof window !== "undefined"
    ? window.location.origin + window.location.pathname 
    : DEFAULT_REDIRECT_URI;

const MSAL_CONFIG = {
  auth: {
    clientId: "5fef90a1-8b89-4f73-a0c9-490e8ec84f4e",
    authority: "https://login.microsoftonline.com/b9465ebd-f455-470a-895c-7e26882877fa",
    redirectUri: runtimeRedirectUri
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
};

const GRAPH_SCOPES = ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"];
/* =========
   Auth (MSAL)
   ========= */
const msal = new PublicClientApplication(MSAL_CONFIG);
const INTERACTION_KEY = "msal.interaction.status";
const isInteractionBusy = () =>
  sessionStorage.getItem(INTERACTION_KEY) === "interaction_in_progress" ||
  localStorage.getItem(INTERACTION_KEY) === "interaction_in_progress";

function clearStuckInteraction() {
  try {
    if (isInteractionBusy()) {
      sessionStorage.removeItem(INTERACTION_KEY);
      localStorage.removeItem(INTERACTION_KEY);
    }
  } catch {}
}

const msalInitPromise = (async () => {
  clearStuckInteraction();
  await msal.initialize();
  try { await msal.handleRedirectPromise(); } catch {}
})();

let activeAccount: AccountInfo | null = null;
let loginPromise: Promise<AccountInfo> | null = null;

function delay(ms: number) { return new Promise(res => setTimeout(res, ms)); }

async function waitWhileBusy(maxMs = 2500) {
  const start = Date.now();
  while (isInteractionBusy() && Date.now() - start < maxMs) await delay(250);
  if (isInteractionBusy()) clearStuckInteraction();
}

async function ensureLogin(): Promise<void> {
  await msalInitPromise;

  const accounts = msal.getAllAccounts();
  if (accounts.length) {
    activeAccount = accounts[0];
    msal.setActiveAccount(activeAccount);
    return;
  }

  if (loginPromise) {
    activeAccount = await loginPromise;
    msal.setActiveAccount(activeAccount);
    return;
  }

  await waitWhileBusy();
  loginPromise = msal.loginPopup({ prompt: "select_account", scopes: GRAPH_SCOPES })
    .then(r => r.account!).finally(() => { loginPromise = null; });

  activeAccount = await loginPromise;
  msal.setActiveAccount(activeAccount);
}

async function getGraphToken(): Promise<string> {
  await msalInitPromise;
  if (!activeAccount) await ensureLogin();

  try {
    const res = await msal.acquireTokenSilent({ account: activeAccount!, scopes: GRAPH_SCOPES });
    return res.accessToken;
  } catch {
    if (loginPromise || isInteractionBusy()) {
      await waitWhileBusy();
      if (loginPromise) await loginPromise;
      const res2 = await msal.acquireTokenSilent({ account: msal.getActiveAccount()!, scopes: GRAPH_SCOPES });
      return res2.accessToken;
    }
    await waitWhileBusy();
    const res = await msal.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    activeAccount = res.account!;
    msal.setActiveAccount(activeAccount);
    return res.accessToken;
  }
}

/* =========
   Graph utils
   ========= */
async function graph<T>(url: string, token: string, init?: RequestInit): Promise<T> {
  const resp = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init?.headers || {})
    }
  });
  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Graph ${resp.status}: ${text}`);
  }
  return resp.status === 204 ? (undefined as unknown as T) : (resp.json() as Promise<T>);
}

async function getSiteId(token: string): Promise<string> {
  const candidateUrls = SITE_IS_UNDER_SITES
    ? [`/sites/${SP_HOSTNAME}:/sites/${SP_SITE_PATH}`]
    : [`/sites/${SP_HOSTNAME}:/${SP_SITE_PATH}`];

  for (const u of candidateUrls) {
    try { return (await graph<{ id: string }>(u, token)).id; }
    catch (e) { console.warn("Site resolve failed with", u, e); }
  }
  throw new Error("Failed to resolve siteId. Check host/path/under-sites flag.");
}

/* =========
   SP list helpers
   ========= */
async function getListId(siteId: string, token: string, displayName: string, wantedServerRelativeUrl: string): Promise<string> {
  const lists = await graph<{ value: Array<{ id: string; displayName: string; webUrl?: string }> }>(
    `/sites/${siteId}/lists?$select=id,displayName,webUrl`, token
  );

  let found = lists.value.find(l => l.displayName === displayName);
  if (found) return found.id;

  const wanted = wantedServerRelativeUrl.toLowerCase();
  found = lists.value.find(l => (l.webUrl || "").toLowerCase().endsWith(wanted));
  if (found) return found.id;

  console.warn("Lists returned:", lists.value.map(l => ({ displayName: l.displayName, webUrl: l.webUrl })));
  throw new Error(`List not found: ${displayName}`);
}

async function getListItemsByField(siteId: string, listId: string, token: string, fieldInternalName: string): Promise<string[]> {
  const res = await graph<{ value: Array<{ id: string; fields: Record<string, any> }> }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields($select=${fieldInternalName})`, token
  );
  const values = res.value.map(it => (it.fields?.[fieldInternalName] ?? "").toString().trim()).filter(Boolean);
  return Array.from(new Set(values)).sort((a, b) => a.localeCompare(b, "he"));
}

async function createListItem(siteId: string, listId: string, token: string, fields: Record<string, any>): Promise<{ id: string }> {
  return graph<{ id: string }>(`/sites/${siteId}/lists/${listId}/items`, token, {
    method: "POST", body: JSON.stringify({ fields })
  });
}

/* =========
   UI helpers
   ========= */
function setSelectDisabled(id: string, disabled: boolean, placeholderWhenDisabled?: string) {
  const el = document.getElementById(id) as HTMLSelectElement | null;
  if (!el) return;
  el.disabled = disabled;
  if (disabled && placeholderWhenDisabled) {
    el.innerHTML = "";
    el.append(new Option(placeholderWhenDisabled, ""));
  }
}

function fillSelectById(id: string, options: string[], firstLabel = "— בחר/י —") {
  const sel = document.getElementById(id) as HTMLSelectElement | null;
  if (!sel) return;
  sel.innerHTML = "";
  sel.append(new Option(firstLabel, ""));
  options.forEach(v => sel.append(new Option(v, v)));
}

function getInputValue(id: string): string {
  const el = document.getElementById(id) as HTMLInputElement | null;
  return (el?.value || "").trim();
}

function getSelectValue(id: string): string {
  const el = document.getElementById(id) as HTMLSelectElement | null;
  return (el?.value || "").trim();
}

function showCloseDocMessage() {
  const root = document.getElementById("app-body");
  if (!root) return;
  root.innerHTML = `
    <div style="direction:rtl; text-align:right; padding:12px">
      <h3 style="margin-top:0">השינויים נשמרו במערכת</h3>
      <p class="ms-font-l">יש לסגור את המסמך על מנת לאפשר לשינויים להתעדכן.</p>
    </div>
  `;
}

async function loadDropdowns() {
  try {
    const token = await getGraphToken();
    const siteId = await getSiteId(token);

    // Status
    setSelectDisabled("statusSelect", true, "— טוען סטטוס… —");
    const statusListId = await getListId(siteId, token, STATUS_LIST_DISPLAY_NAME, STATUS_LIST_SERVER_RELATIVE_URL);
    const statusValues = await getListItemsByField(siteId, statusListId, token, STATUS_FIELD_NAME);
    fillSelectById("statusSelect", statusValues, "— בחר/י סטטוס —");
    setSelectDisabled("statusSelect", false);

    // Projects
    setSelectDisabled("projectSelect", true, "— טוען פרויקטים… —");
    const projectsListId = await getListId(siteId, token, PROJECTS_LIST_DISPLAY_NAME, PROJECTS_LIST_SERVER_RELATIVE_URL);
    const projectValues = await getListItemsByField(siteId, projectsListId, token, PROJECTS_FIELD_NAME);
    fillSelectById("projectSelect", projectValues, "— בחר/י פרויקט —");
    setSelectDisabled("projectSelect", false);
  } catch (e: any) {
    console.error("Dropdowns load error:", e);
    alert("לא ניתן לטעון ערכים לסטטוס/פרויקט. בדקי הרשאות וכתובות.");
  }
}

/* =========
   Word actions
   ========= */
async function updateDocumentFields(fields: { recipient?: string; status?: string; address?: string; project?: string }) {
  await Word.run(async (context) => {
    const tags = ["recipient", "status", "address", "project"] as const;

    // Load existing content controls by tags once
    const maps = await Promise.all(tags.map(async (t) => {
      const col = context.document.contentControls.getByTag(t);
      col.load("items");
      return col;
    }));
    await context.sync();

    const values: Record<string, string | undefined> = { ...fields };
    tags.forEach((tag, i) => {
      const val = (values[tag] || "").toString();
      if (!val) return;
      const col = maps[i];
      const existing = col.items[0];
      if (existing) {
        existing.insertText(val, Word.InsertLocation.replace);
      } else {
        const range = context.document.getSelection();
        const cc = range.insertContentControl();
        cc.tag = tag;
        cc.title = tag;
        cc.insertText(val, Word.InsertLocation.replace);
      }
    });

    await context.sync();
  });
}

async function getCurrentDocumentUrl(): Promise<string | null> {
  return new Promise((resolve) => {
    Office.context.document.getFilePropertiesAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        const url = (res.value?.url as string) || "";
        resolve(url || null);
      } else {
        resolve(null);
      }
    });
  });
}

/* =========
   File identity helpers
   ========= */
// Build Graph shareId from a webUrl (UTF-8 safe)
function toShareIdFromWebUrl(webUrl: string): string {
  const bytes = new TextEncoder().encode(webUrl);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  const b64 = btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return `u!${b64}`;
}

// Get List Item ID from file webUrl
async function getListItemIdByWebUrl(token: string, webUrl: string): Promise<string | null> {
  const shareId = toShareIdFromWebUrl(webUrl);
  const data = await graph<{ id: string; listItem?: { id: string } }>(
    `/shares/${encodeURIComponent(shareId)}/driveItem?$expand=listItem`, token
  );
  return data.listItem?.id ?? null;
}

// Extract library display name from file webUrl
function getLibraryNameFromWebUrl(webUrl: string): string | null {
  try {
    const u = new URL(webUrl);
    const parts = u.pathname.replace(/^\/+/, "").split("/");
    if (SITE_IS_UNDER_SITES) {
      if (parts.length >= 3 && parts[0].toLowerCase() === "sites") {
        return decodeURIComponent(parts[2] || "");
      }
    } else {
      if (parts.length >= 2 && parts[0].toLowerCase() === SP_SITE_PATH.toLowerCase()) {
        return decodeURIComponent(parts[1] || "");
      }
    }
    return null;
  } catch { return null; }
}

/* =========
   Save to Helper list
   ========= */
async function saveToHelper(fields: {
  recipient: string;
  status: string;
  address: string;
  project: string;
  titleItemId: string;
  libName?: string;
}): Promise<string> {
  const token = await getGraphToken();
  const siteId = await getSiteId(token);
  const helperListId = await getListId(siteId, token, HELPER_LIST_DISPLAY_NAME, HELPER_LIST_SERVER_RELATIVE_URL);

  // Field names in helper list should be identical to keys here
  const helperFields: Record<string, any> = {
    Title: fields.titleItemId,
    recipient: fields.recipient,
    status: fields.status,
    address: fields.address,
    project: fields.project
  };
  if (fields.libName) helperFields.libName = fields.libName;

  const created = await createListItem(siteId, helperListId, token, helperFields);
  return created.id;
}

/* =========
   Buttons
   ========= */
export async function runUpdateDoc() {
  const recipient = getInputValue("recipientInput");
  const status = getSelectValue("statusSelect");
  const address = getInputValue("addressInput");
  const project = getSelectValue("projectSelect");

  if (!recipient && !status && !address && !project) {
    alert("יש למלא לפחות שדה אחד לעדכון במסמך.");
    return;
  }

  try {
    await updateDocumentFields({ recipient, status, address, project });
    const lbl = document.getElementById("item-subject");
    if (lbl) lbl.textContent = "המסמך עודכן בהצלחה.";
  } catch (e: any) {
    console.error("runUpdateDoc error:", e);
    alert("שגיאה בעדכון המסמך: " + (e?.message || "לא ידועה"));
  }
}

// Treat Word content-control placeholder text as "missing"
function isCcPlaceholder(raw: string): boolean {
  if (!raw) return true;
  // normalize: trim + collapse spaces + lower
  const s = raw.replace(/\u200f|\u200e|\u202a|\u202c/g, "") // strip RTL/LTR marks
               .replace(/\s+/g, " ")
               .trim()
               .toLowerCase();

  // common placeholders (he/en + legacy variants)
  const candidates = [
    "לחץ או הקש כאן להזנת טקסט.",
    "לחץ או הקש כאן להזנת טקסט",
    "click or tap here to enter text.",
    "click or tap here to enter text",
    "click here to enter text",
    "type your text"
  ].map(x => x.toLowerCase());

  return !s || candidates.includes(s);
}

function normalizeDocValue(raw?: string): string | undefined {
  if (!raw) return undefined;
  return isCcPlaceholder(raw) ? undefined : raw.trim();
}


export async function runSaveSystem() {
  // 1) Read UI
  const uiRecipient = getInputValue("recipientInput");
  const uiStatus    = getSelectValue("statusSelect");
  const uiAddress   = getInputValue("addressInput");
  const uiProject   = getSelectValue("projectSelect");

  try {
    // 2) Update only fields the user actually typed/selected
    await updateDocumentFields({
      recipient: uiRecipient || undefined,
      status:    uiStatus || undefined,
      address:   uiAddress || undefined,
      project:   uiProject || undefined
    });

    // 3) Ensure we have full values for Helper: UI -> doc -> "נתון חסר"
    const docVals = await readDocumentFields();
  const recipient = (uiRecipient || docVals.recipient) ?? "נתון חסר";
  const status    = (uiStatus    || docVals.status)    ?? "נתון חסר";
  const address   = (uiAddress   || docVals.address)   ?? "נתון חסר";
  const project   = (uiProject   || docVals.project)   ?? "נתון חסר";


    // 4) Resolve file identity
    const absUrl = await getCurrentDocumentUrl();
    if (!absUrl) { alert("לא מזוהה כתובת למסמך. שמרי את המסמך ב-SharePoint ונסי שוב."); return; }

    const token = await getGraphToken();
    const itemId = await getListItemIdByWebUrl(token, absUrl);
    if (!itemId) { alert("לא ניתן להביא את מזהה הפריט של המסמך (List Item ID)."); return; }

    const libName = getLibraryNameFromWebUrl(absUrl) || undefined;

    // 5) Save to helper (all fields guaranteed non-empty now)
    await saveToHelper({
      recipient, status, address, project,
      titleItemId: itemId,
      libName
    });

    // 6) Save the Word doc (optional but recommended)
    await Word.run(async (ctx) => { await ctx.document.save(); });

    // 7) Final message
    showCloseDocMessage();

  } catch (e: any) {
    console.error("runSaveSystem error:", e);
    alert("שגיאה בשמירה במערכת: " + (e?.message || "לא ידועה"));
  }
}


async function readDocumentFields(): Promise<{
  recipient?: string; status?: string; address?: string; project?: string;
}> {
  const TAGS = ["recipient", "status", "address", "project"] as const;
  const result: Record<string, string | undefined> = {};

  await Word.run(async (context) => {
    const collections = TAGS.map(tag => {
      const col = context.document.contentControls.getByTag(tag);
      col.load("items");
      return { tag, col };
    });
    await context.sync();

    const ranges: Array<{ tag: string; range: Word.Range }> = [];
    for (const { tag, col } of collections) {
      const cc = col.items[0];
      if (!cc) continue;
      const r = cc.getRange();
      r.load("text");
      ranges.push({ tag, range: r });
    }
    await context.sync();

    for (const r of ranges) {
      const val = (r.range.text ?? "").toString();
      result[r.tag] = normalizeDocValue(val);
    }
  });

  return {
    recipient: result.recipient,
    status:    result.status,
    address:   result.address,
    project:   result.project
  };
}

/* =========
   Bootstrap
   ========= */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    (document.getElementById("sideload-msg") as HTMLElement).style.display = "none";
    (document.getElementById("app-body") as HTMLElement).style.display = "flex";

    const btnUpdate = document.getElementById("runUpdateDoc");
    if (btnUpdate) (btnUpdate as HTMLDivElement).onclick = runUpdateDoc;

    const btnSave = document.getElementById("runSaveSystem");
    if (btnSave) (btnSave as HTMLDivElement).onclick = runSaveSystem;

    loadDropdowns();
  } else {
    (document.getElementById("sideload-msg") as HTMLElement).style.display = "block";
    (document.getElementById("app-body") as HTMLElement).style.display = "none";
  }
});
