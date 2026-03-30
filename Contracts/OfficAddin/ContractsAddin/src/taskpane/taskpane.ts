import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";

/* =========
   Config
   ========= */
const SP_HOSTNAME = "africaisrael.sharepoint.com";
const SP_SITE_PATH = "ContractsNew";
const SITE_IS_UNDER_SITES = false;

/* =========
   LOOKUP LISTS
   ========= */
const STATUS_LIST_DISPLAY_NAME = "ContractStatus";
const STATUS_FIELD_NAME = "Title";

const PROJECTS_LIST_DISPLAY_NAME = "projects";
const PROJECTS_FIELD_NAME = "Title";

const TEMPLATES_LIST_DISPLAY_NAME = "ניהול תבניות";
const TEMPLATES_FIELD_NAME = "Title";

const SITES_LIST_DISPLAY_NAME = "אתרים";
const SITES_FIELD_NAME = "Title";

const MUNICIPALITIES_LIST_DISPLAY_NAME = "רשויות מקומיות";
const MUNICIPALITIES_FIELD_NAME = "Title";

const INDEX_TYPES_LIST_DISPLAY_NAME = "סוגי מדדים";
const INDEX_TYPES_FIELD_NAME = "Title";

const COMPANIES_LIST_DISPLAY_NAME = "חברות";
const COMPANIES_FIELDS = {
  title: "Title",
  address: "Address",
  hp: "hp"
};

const SUPPLIER_TYPES_LIST_DISPLAY_NAME = "סוגי ספקים";
const SUPPLIER_TYPES_FIELD_NAME = "Title";

const SUPPLIERS_LIST_DISPLAY_NAME = "Suppliers";
const SUPPLIERS_FIELDS = {
  title: "Title",
  address: "Address",
  type: "SupplierType"
};

const TAGS = {
  contractNumber: "cntContractNumber",
  contractVersion: "cntContractVersion", // ✅ חובה
  template: "cntTemplateName",

  project: "cntProjectName",
  site: "cntSite",
  supplier: "cntSupllierName", // ⚠️ להשאיר אם זה מה שיש בתבנית

  municipality: "cntMunicipality",
  workDescription: "cntWorkDescription",
  signDate: "cntSignDate",
  startDate: "cntStartDate",
  months: "cntDurationMonths",
  expectedEndDate: "cntExpectedEndDate", // ✅ חובה

  status: "cntStatus",
  recipient: "cntTzadA",
  otherSides: "cntTzadB",

  costCompMethod: "cntCostCompMethod",
  costContractScope: "cntCostContractScope",
  costCurrency: "cntCostCurrency",
  costIndexType: "cntCostIndexType",
  costBaseIndexDate: "cntCostBaseIndexDate",
  costIndexMode: "cntCostIndexMode",
  costIndexPoints: "cntCostIndexPoints",
  costPaymentTerms: "cntCostPaymentTerms",

  // ✅ TODO
  localAuth: "cntLocalAuth",
  tzadAPercent: "cmtTzadAPercent",
  madadTypeTitle: "cntMadadTypeTitle",
  isKnownTitle: "cntIsKnownTitle",
  madadBase: "cntMadadBase",
  madadPoints: "cntMadadPoints",
  jobDesc: "cntJobDesc",

  // ✅ NEW: Party A Name (separate from summary)
  partyAName: "cntPartyAName",

  // ✅ NEW: Custom fields 1..8
  customField1: "cntCustomField1",
  customField2: "cntCustomField2",
  customField3: "cntCustomField3",
  customField4: "cntCustomField4",
  customField5: "cntCustomField5",
  customField6: "cntCustomField6",
  customField7: "cntCustomField7",
  customField8: "cntCustomField8",
} as const;

// ================================
// ✅ Fields tab: catalog
// ================================
type FieldDef = { label: string; tag: string };

const FIELD_CATALOG: FieldDef[] = [
  { label: "מספר חוזה", tag: TAGS.contractNumber },
  { label: "גרסת חוזה", tag: TAGS.contractVersion },
  { label: "תבנית", tag: TAGS.template },

  { label: "שם הפרויקט", tag: TAGS.project },
  { label: "אתר", tag: TAGS.site },
  { label: "ספק", tag: TAGS.supplier },

  { label: "רשות מקומית", tag: TAGS.municipality },
  { label: "תיאור העבודה", tag: TAGS.workDescription },

  { label: "תאריך חתימה החוזה", tag: TAGS.signDate },
  { label: "מועד התחלה", tag: TAGS.startDate },
  { label: "מספר חודשים", tag: TAGS.months },
  { label: "מועד סיום צפוי", tag: TAGS.expectedEndDate },

  { label: "סטטוס", tag: TAGS.status },

  { label: "צד א' בחוזה", tag: TAGS.recipient },
  { label: "צדדים נוספים", tag: TAGS.otherSides },

  // ✅ NEW: separate name for צד א'
  { label: "שם צד א'", tag: TAGS.partyAName },

  { label: "אופן התמורה", tag: TAGS.costCompMethod },
  { label: "היקף החוזה", tag: TAGS.costContractScope },
  { label: "שם מטבע", tag: TAGS.costCurrency },
  { label: "סוג מדד", tag: TAGS.costIndexType },
  { label: "מדד בסיס (תאריך)", tag: TAGS.costBaseIndexDate },
  { label: "שיטת מדד", tag: TAGS.costIndexMode },
  { label: "נקודות מדד", tag: TAGS.costIndexPoints },
  { label: "תנאי תשלום", tag: TAGS.costPaymentTerms },

  // TODO
  { label: "רשות מקומית (LocalAuth)", tag: TAGS.localAuth },
  { label: "אחוז השתתפות צד א'", tag: TAGS.tzadAPercent },
  { label: "סוג מדד (כותרת)", tag: TAGS.madadTypeTitle },
  { label: "מדד בגין/ידוע (כותרת)", tag: TAGS.isKnownTitle },
  { label: "מדד בסיס", tag: TAGS.madadBase },
  { label: "נקודות מדד (Madad)", tag: TAGS.madadPoints },
  { label: "תיאור עבודה (JobDesc)", tag: TAGS.jobDesc },

  // ✅ NEW: Custom 1..8
  { label: "שדה נוסף 1", tag: TAGS.customField1 },
  { label: "שדה נוסף 2", tag: TAGS.customField2 },
  { label: "שדה נוסף 3", tag: TAGS.customField3 },
  { label: "שדה נוסף 4", tag: TAGS.customField4 },
  { label: "שדה נוסף 5", tag: TAGS.customField5 },
  { label: "שדה נוסף 6", tag: TAGS.customField6 },
  { label: "שדה נוסף 7", tag: TAGS.customField7 },
  { label: "שדה נוסף 8", tag: TAGS.customField8 },
];

/* =========
   Helper list target
   ========= */
const HELPER_LIST_DISPLAY_NAME = "FieldsUpdateHelper";
const HELPER_LIST_SERVER_RELATIVE_URL = SITE_IS_UNDER_SITES
  ? `/sites/${SP_SITE_PATH}/Lists/${HELPER_LIST_DISPLAY_NAME}`
  : `/${SP_SITE_PATH}/Lists/${HELPER_LIST_DISPLAY_NAME}`;

function listUrl(displayName: string) {
  return SITE_IS_UNDER_SITES
    ? `/sites/${SP_SITE_PATH}/Lists/${displayName}`
    : `/${SP_SITE_PATH}/Lists/${displayName}`;
}

const STATUS_LIST_SERVER_RELATIVE_URL = listUrl(STATUS_LIST_DISPLAY_NAME);
const PROJECTS_LIST_SERVER_RELATIVE_URL = listUrl(PROJECTS_LIST_DISPLAY_NAME);
const TEMPLATES_LIST_SERVER_RELATIVE_URL = listUrl(TEMPLATES_LIST_DISPLAY_NAME);
const SITES_LIST_SERVER_RELATIVE_URL = listUrl(SITES_LIST_DISPLAY_NAME);
const MUNICIPALITIES_LIST_SERVER_RELATIVE_URL = listUrl(MUNICIPALITIES_LIST_DISPLAY_NAME);
const INDEX_TYPES_LIST_SERVER_RELATIVE_URL = listUrl(INDEX_TYPES_LIST_DISPLAY_NAME);
const COMPANIES_LIST_SERVER_RELATIVE_URL = listUrl(COMPANIES_LIST_DISPLAY_NAME);
const SUPPLIER_TYPES_LIST_SERVER_RELATIVE_URL = listUrl(SUPPLIER_TYPES_LIST_DISPLAY_NAME);
const SUPPLIERS_LIST_SERVER_RELATIVE_URL = listUrl(SUPPLIERS_LIST_DISPLAY_NAME);

/* =========
   MSAL
   ========= */
const MSAL_CONFIG = {
  auth: {
    clientId: "d8f0fc93-7736-43c1-8e12-8e193f543cd4",
    authority: "https://login.microsoftonline.com/b4d149d3-3aef-42b5-a6f1-b5018284caf9",
    redirectUri: "https://knowedge.co.il/matrix/downloads/taskpane.html"
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
};

const GRAPH_SCOPES = ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"];

/* =========
   Auth (MSAL)
   ========= */

// Key used to persist the last signed-in account's homeAccountId across pane reloads.
// This supplements MSAL's own localStorage cache and makes restore more deterministic.
const LAST_ACCOUNT_KEY = "contracts_addin_last_account_id";

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

// Persist the homeAccountId after a successful login so the next pane open can restore it.
function saveAccountHint(account: AccountInfo) {
  try { localStorage.setItem(LAST_ACCOUNT_KEY, account.homeAccountId); } catch {}
}

// Try to restore a previously used account from MSAL's cache via the stored hint.
function tryRestoreAccountFromHint(): AccountInfo | null {
  try {
    const storedId = localStorage.getItem(LAST_ACCOUNT_KEY);
    if (storedId) {
      const account = msal.getAccountByHomeId(storedId);
      if (account) {
        console.log("[auth] restored account from hint:", account.username);
        return account;
      }
    }
  } catch {}
  return null;
}

async function ensureLogin(): Promise<void> {
  await msalInitPromise;

  // 1. Try to restore the previously used account via our own stored hint (most deterministic)
  let account = tryRestoreAccountFromHint();

  // 2. Fall back to any account MSAL has in its localStorage cache
  if (!account) {
    const accounts = msal.getAllAccounts();
    console.log("[auth] getAllAccounts:", accounts.length, accounts.map(a => a.username));
    if (accounts.length) account = accounts[0];
  }

  if (account) {
    activeAccount = account;
    msal.setActiveAccount(activeAccount);
    console.log("[auth] ensureLogin: using cached account:", activeAccount.username);
    return;
  }

  if (loginPromise) {
    activeAccount = await loginPromise;
    msal.setActiveAccount(activeAccount);
    return;
  }

  // 3. Try ssoSilent before showing any interactive UI.
  //    This succeeds when the user has an active AAD session cookie (common on corporate devices).
  //    It uses a hidden iframe; may fail if third-party cookies are blocked — that's expected.
  try {
    console.log("[auth] no cached account — trying ssoSilent...");
    const silentResult = await msal.ssoSilent({ scopes: GRAPH_SCOPES });
    activeAccount = silentResult.account!;
    msal.setActiveAccount(activeAccount);
    saveAccountHint(activeAccount);
    console.log("[auth] ssoSilent succeeded:", activeAccount.username);
    return;
  } catch (ssoErr) {
    console.log("[auth] ssoSilent failed (expected when no active session):", (ssoErr as any)?.errorCode);
  }

  // 4. Last resort: show the login popup.
  //    prompt: "select_account" removed — without it, AAD reuses an existing session automatically.
  console.log("[auth] falling back to interactive loginPopup");
  await waitWhileBusy();
  loginPromise = msal.loginPopup({ scopes: GRAPH_SCOPES })
    .then(r => r.account!).finally(() => { loginPromise = null; });

  activeAccount = await loginPromise;
  msal.setActiveAccount(activeAccount);
  saveAccountHint(activeAccount);
}

async function getGraphToken(): Promise<string> {
  await msalInitPromise;
  if (!activeAccount) await ensureLogin();

  try {
    const res = await msal.acquireTokenSilent({ account: activeAccount!, scopes: GRAPH_SCOPES });
    console.log("[auth] acquireTokenSilent succeeded");
    return res.accessToken;
  } catch (silentErr) {
    console.warn("[auth] acquireTokenSilent failed:", (silentErr as any)?.errorCode, (silentErr as any)?.message);
    if (loginPromise || isInteractionBusy()) {
      await waitWhileBusy();
      if (loginPromise) await loginPromise;
      const res2 = await msal.acquireTokenSilent({ account: msal.getActiveAccount()!, scopes: GRAPH_SCOPES });
      return res2.accessToken;
    }
    console.log("[auth] falling back to acquireTokenPopup");
    await waitWhileBusy();
    const res = await msal.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    activeAccount = res.account!;
    msal.setActiveAccount(activeAccount);
    saveAccountHint(activeAccount);
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

  throw new Error(`List not found: ${displayName}`);
}

async function getListItemsByField(siteId: string, listId: string, token: string, fieldInternalName: string): Promise<string[]> {
  const res = await graph<{ value: Array<{ id: string; fields: Record<string, any> }> }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields($select=${fieldInternalName})`, token
  );
  const values = res.value
    .map(it => (it.fields?.[fieldInternalName] ?? "").toString().trim())
    .filter(Boolean);
  return Array.from(new Set(values)).sort((a, b) => a.localeCompare(b, "he"));
}

type ListItemFields = { id: string; fields: Record<string, any> };

async function getListItemsFields(siteId: string, listId: string, token: string, selectFields: string[]): Promise<ListItemFields[]> {
  const unique = Array.from(new Set(selectFields.filter(Boolean)));
  const select = unique.join(",");
  const res = await graph<{ value: Array<{ id: string; fields: Record<string, any> }> }>(
    `/sites/${siteId}/lists/${listId}/items?expand=fields($select=${select})`, token
  );
  return res.value || [];
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

function fillSelectByIdPairs(id: string, options: Array<{ value: string; label: string }>, firstLabel = "— בחר/י —") {
  const sel = document.getElementById(id) as HTMLSelectElement | null;
  if (!sel) return;
  sel.innerHTML = "";
  sel.append(new Option(firstLabel, ""));
  options.forEach(o => sel.append(new Option(o.label, o.value)));
}

function getInputValue(id: string): string {
  const el = document.getElementById(id) as HTMLInputElement | HTMLTextAreaElement | null;
  return (el?.value || "").trim();
}

function setInputValue(id: string, value: string) {
  const el = document.getElementById(id) as HTMLInputElement | HTMLTextAreaElement | null;
  if (!el) return;
  el.value = value ?? "";
}

function getSelectValue(id: string): string {
  const el = document.getElementById(id) as HTMLSelectElement | null;
  return (el?.value || "").trim();
}

function setSelectValue(id: string, value: string) {
  const el = document.getElementById(id) as HTMLSelectElement | null;
  if (!el) return;
  el.value = (value ?? "").toString();
}

function setText(id: string, text: string) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = text;
}

function toggleHidden(id: string, hidden: boolean) {
  const el = document.getElementById(id);
  if (!el) return;
  el.classList.toggle("hidden", hidden);
}

function pickField(fields: Record<string, any>, ...keys: string[]) {
  for (const k of keys) {
    const v = fields?.[k];
    if (v !== undefined && v !== null && `${v}`.trim() !== "") return v;
  }
  return undefined;
}

/* =========
   Local UI state
   ========= */
const uiState = {
  partyA: {
    companyName: "",
    address: "",
    hp: "",
    contactName: "", 
    namePercent: "",
    summary: ""
  },
  partyB: {
    useSuppliers: false,
    manualCompany: "",
    manualAddress: "",
    manualHp: "",
    supplierType: "",
    supplierName: "",
    supplierAddress: "",
    summary: ""
  },
  cost: {
    compMethod: "",
    contractScope: "",
    currency: "",
    indexType: "",
    baseIndexDate: "",
    indexMode: "",
    indexPoints: "",
    paymentTerms: "",
    summary: ""
  },
  //  Custom fields 1..8 (saved to library as customField1..8)
  custom: {
    customField1: "",
    customField2: "",
    customField3: "",
    customField4: "",
    customField5: "",
    customField6: "",
    customField7: "",
    customField8: ""
  }
};

/* =========
   Lookup caches
   ========= */
type Company = { title: string; address: string; hp: string };
type Supplier = { title: string; address: string; type: string };

let companiesByTitle = new Map<string, Company>();
let suppliersAll: Supplier[] = [];

// Task 5: accumulates committed supplier names (append mode)
let committedSupplierNames = "";

/* =========
   Dates calc (ExpectedEndDate)
   ========= */
function addMonthsToDateISO(startIso: string, months: number): string {
  if (!startIso || !months || months <= 0) return "";
  const [y, m, d] = startIso.split("-").map(Number);
  if (!y || !m || !d) return "";
  const dt = new Date(Date.UTC(y, m - 1, d));
  dt.setUTCMonth(dt.getUTCMonth() + months);

  const yy = dt.getUTCFullYear();
  const mm = String(dt.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(dt.getUTCDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

function wireDates() {
  const startEl = document.getElementById("startDateInput") as HTMLInputElement | null;
  const monthsEl = document.getElementById("monthsInput") as HTMLInputElement | null;

  const recalc = () => {
    const start = getInputValue("startDateInput");
    const months = Number(getInputValue("monthsInput") || 0);
    const end = addMonthsToDateISO(start, months);
    setInputValue("expectedEndDateInput", end);
    refreshReadonly();
  };

  if (startEl) startEl.addEventListener("change", recalc);
  if (monthsEl) monthsEl.addEventListener("input", recalc);
}

/* =========
   Parties summaries
   ========= */
function buildPartyASummary() {
  const lines: string[] = [];
  if (uiState.partyA.companyName) lines.push(`חברה: ${uiState.partyA.companyName}`);
  if (uiState.partyA.address) lines.push(`כתובת: ${uiState.partyA.address}`);
  if (uiState.partyA.hp) lines.push(`ח.פ: ${uiState.partyA.hp}`);
  if (uiState.partyA.contactName) lines.push(`שם: ${uiState.partyA.contactName}`); // ✅ NEW
  if (uiState.partyA.namePercent) lines.push(`שם ואחוז: ${uiState.partyA.namePercent}`);
  return lines.join("\n");
}

function buildPartyBSummary() {
  const lines: string[] = [];
  if (uiState.partyB.useSuppliers) {
    if (uiState.partyB.supplierType) lines.push(`סוג ספק: ${uiState.partyB.supplierType}`);
    if (uiState.partyB.supplierName) lines.push(`ספק: ${uiState.partyB.supplierName}`);
    if (uiState.partyB.supplierAddress) lines.push(`כתובת: ${uiState.partyB.supplierAddress}`);
  } else {
    if (uiState.partyB.manualCompany) lines.push(`חברה: ${uiState.partyB.manualCompany}`);
    if (uiState.partyB.manualAddress) lines.push(`כתובת: ${uiState.partyB.manualAddress}`);
    if (uiState.partyB.manualHp) lines.push(`ח.פ: ${uiState.partyB.manualHp}`);
  }
  return lines.join("\n");
}

/* =========
   Cost summary + preview
   ========= */
function getCostIndexMode(): string {
  const r1 = document.getElementById("costIndexModeFor") as HTMLInputElement | null;
  const r2 = document.getElementById("costIndexModeKnown") as HTMLInputElement | null;
  if (r1 && r1.checked) return r1.value;
  if (r2 && r2.checked) return r2.value;
  return "";
}

function setCostIndexMode(value: string) {
  const r1 = document.getElementById("costIndexModeFor") as HTMLInputElement | null;
  const r2 = document.getElementById("costIndexModeKnown") as HTMLInputElement | null;
  if (r1) r1.checked = value === r1.value;
  if (r2) r2.checked = value === r2.value;
}

function buildCostSummary(): string {
  const lines: string[] = [];
  if (uiState.cost.compMethod) lines.push(`אופן התמורה: ${uiState.cost.compMethod}`);
  if (uiState.cost.contractScope) lines.push(`היקף החוזה: ${uiState.cost.contractScope}`);
  if (uiState.cost.currency) lines.push(`מטבע: ${uiState.cost.currency}`);
  if (uiState.cost.indexType) lines.push(`סוג מדד: ${uiState.cost.indexType}`);
  if (uiState.cost.baseIndexDate) lines.push(`מדד בסיס (תאריך): ${uiState.cost.baseIndexDate}`);
  if (uiState.cost.indexMode) lines.push(`שיטת מדד: ${uiState.cost.indexMode}`);
  if (uiState.cost.indexPoints) lines.push(`נקודות מדד: ${uiState.cost.indexPoints}`);
  if (uiState.cost.paymentTerms) lines.push(`תנאי תשלום: ${uiState.cost.paymentTerms}`);
  return lines.join("\n");
}

function refreshCostStateFromUI() {
  uiState.cost.compMethod = getSelectValue("costCompMethodSelect");
  uiState.cost.contractScope = getInputValue("costContractScopeInput");
  uiState.cost.currency = getSelectValue("costCurrencySelect");
  uiState.cost.indexType = getSelectValue("costIndexTypeSelect");
  uiState.cost.baseIndexDate = getInputValue("costBaseIndexDateInput");
  uiState.cost.indexMode = getCostIndexMode();
  uiState.cost.indexPoints = getInputValue("costIndexPointsInput");
  uiState.cost.paymentTerms = getSelectValue("costPaymentTermsSelect");
}

function refreshCostPreview() {
  const txt = uiState.cost.summary || "לא הוזנו נתוני עלות עדיין.";
  setText("costPreview", txt);
}

function wireCostUI() {
  const addBtn = document.getElementById("costAddBtn") as HTMLButtonElement | null;
  if (addBtn) {
    addBtn.addEventListener("click", () => {
      refreshCostStateFromUI();
      uiState.cost.summary = buildCostSummary();
      refreshCostPreview();
      refreshReadonly();
    });
  }

  const clearBtn = document.getElementById("costClearBtn") as HTMLButtonElement | null;
  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      const compSel = document.getElementById("costCompMethodSelect") as HTMLSelectElement | null;
      if (compSel) compSel.value = "";
      setInputValue("costContractScopeInput", "");

      const curSel = document.getElementById("costCurrencySelect") as HTMLSelectElement | null;
      if (curSel) curSel.value = "";

      const idxSel = document.getElementById("costIndexTypeSelect") as HTMLSelectElement | null;
      if (idxSel) idxSel.value = "";

      setInputValue("costBaseIndexDateInput", "");
      setCostIndexMode("");
      setInputValue("costIndexPointsInput", "");

      const paySel = document.getElementById("costPaymentTermsSelect") as HTMLSelectElement | null;
      if (paySel) paySel.value = "";

      uiState.cost = {
        compMethod: "",
        contractScope: "",
        currency: "",
        indexType: "",
        baseIndexDate: "",
        indexMode: "",
        indexPoints: "",
        paymentTerms: "",
        summary: ""
      };

      refreshCostPreview();
      refreshReadonly();
    });
  }

  const ids = [
    "costCompMethodSelect",
    "costContractScopeInput",
    "costCurrencySelect",
    "costIndexTypeSelect",
    "costBaseIndexDateInput",
    "costIndexPointsInput",
    "costPaymentTermsSelect"
  ];
  ids.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("change", refreshReadonly);
    el.addEventListener("input", refreshReadonly);
  });

  const r1 = document.getElementById("costIndexModeFor") as HTMLInputElement | null;
  const r2 = document.getElementById("costIndexModeKnown") as HTMLInputElement | null;
  if (r1) r1.addEventListener("change", refreshReadonly);
  if (r2) r2.addEventListener("change", refreshReadonly);
}

/* =========
   Previews + readonly
   ========= */
function refreshPartyPreviews() {
  const a = uiState.partyA.summary || "לא הוזן צד א' עדיין.";
  const b = uiState.partyB.summary || "לא הוזן צד ב' עדיין.";
  setText("partyAPreview", a);
  setText("partyBPreview", b);
}

function refreshReadonly() {
  const generalLines: string[] = [];
  const contractNumber = getInputValue("contractNumberInput");
  const contractVersion = getInputValue("contractVersionInput");
  const template = getSelectValue("templateSelect");
  const project = getSelectValue("projectSelect");
  const site = getSelectValue("siteSelect");
  const municipality = getSelectValue("municipalitySelect");
  const workDesc = getInputValue("workDescriptionInput");
  const signDate = getInputValue("signDateInput");
  const startDate = getInputValue("startDateInput");
  const months = getInputValue("monthsInput");
  const expectedEnd = getInputValue("expectedEndDateInput");
  const status = getSelectValue("statusSelect");

  if (contractNumber) generalLines.push(`מספר חוזה: ${contractNumber}`);
  if (contractVersion) generalLines.push(`גרסת חוזה: ${contractVersion}`);
  if (template) generalLines.push(`תבנית: ${template}`);
  if (project) generalLines.push(`פרויקט: ${project}`);
  if (site) generalLines.push(`אתר: ${site}`);
  if (municipality) generalLines.push(`רשות מקומית: ${municipality}`);
  if (workDesc) generalLines.push(`תיאור העבודה: ${workDesc}`);
  if (signDate) generalLines.push(`תאריך חתימה: ${signDate}`);
  if (startDate) generalLines.push(`מועד התחלה: ${startDate}`);
  if (months) generalLines.push(`מספר חודשים: ${months}`);
  if (expectedEnd) generalLines.push(`מועד סיום צפוי: ${expectedEnd}`);
  if (status) generalLines.push(`סטטוס: ${status}`);

  const partiesLines: string[] = [];
  partiesLines.push("צד א':");
  partiesLines.push(uiState.partyA.summary ? uiState.partyA.summary : "—");
  partiesLines.push("");
  partiesLines.push("צד ב':");
  partiesLines.push(uiState.partyB.summary ? uiState.partyB.summary : "—");

  const costLines: string[] = [];
  costLines.push(uiState.cost.summary ? uiState.cost.summary : "—");

  // custom fields preview (optional block)
  const customLines: string[] = [];
  const c = uiState.custom;
  const pairs: Array<[string, string]> = [
    ["שדה נוסף 1", c.customField1],
    ["שדה נוסף 2", c.customField2],
    ["שדה נוסף 3", c.customField3],
    ["שדה נוסף 4", c.customField4],
    ["שדה נוסף 5", c.customField5],
    ["שדה נוסף 6", c.customField6],
    ["שדה נוסף 7", c.customField7],
    ["שדה נוסף 8", c.customField8],
  ];
  pairs.forEach(([label, val]) => { if (val) customLines.push(`${label}: ${val}`); });

  setText("readonlyGeneral", generalLines.length ? generalLines.join("\n") : "—");
  setText("readonlyParties", partiesLines.join("\n"));
  setText("readonlyCost", costLines.join("\n"));

  if (document.getElementById("readonlyCustom")) {
    setText("readonlyCustom", customLines.length ? customLines.join("\n") : "—");
  }
}

/* =========
   Parties wiring
   ========= */
function wirePartiesUI() {
  const companyASelect = document.getElementById("companyASelect") as HTMLSelectElement | null;
  if (companyASelect) {
    companyASelect.addEventListener("change", () => {
      const title = getSelectValue("companyASelect");
      const c = companiesByTitle.get(title);
      uiState.partyA.companyName = title;
      uiState.partyA.address = c ? (c.address || "") : "";
      uiState.partyA.hp = c ? (c.hp || "") : "";
      setInputValue("companyAAddressInput", uiState.partyA.address);
      setInputValue("companyAHpInput", uiState.partyA.hp);
      refreshReadonly();
    });
  }

  const partyANameEl = document.getElementById("partyANameInput") as HTMLInputElement | null;
  if (partyANameEl) {
    partyANameEl.addEventListener("input", () => {
      uiState.partyA.contactName = getInputValue("partyANameInput");
      refreshReadonly();
    });
  }

  const partyANamePercent = document.getElementById("partyANamePercentInput") as HTMLInputElement | null;
  if (partyANamePercent) {
    partyANamePercent.addEventListener("input", () => {
      uiState.partyA.namePercent = getInputValue("partyANamePercentInput");
      refreshReadonly();
    });
  }

  const partyAAddBtn = document.getElementById("partyAAddBtn") as HTMLButtonElement | null;
  if (partyAAddBtn) {
    partyAAddBtn.addEventListener("click", () => {
      // Task 5: append instead of overwrite
      const newA = buildPartyASummary();
      uiState.partyA.summary = uiState.partyA.summary ? uiState.partyA.summary + "\n" + newA : newA;
      refreshPartyPreviews();
      refreshReadonly();
    });
  }

  const partyAClearBtn = document.getElementById("partyAClearBtn") as HTMLButtonElement | null;
  if (partyAClearBtn) {
    partyAClearBtn.addEventListener("click", () => {
      uiState.partyA.companyName = "";
      uiState.partyA.address = "";
      uiState.partyA.hp = "";
      uiState.partyA.contactName = ""; // ✅ NEW
      uiState.partyA.namePercent = "";
      uiState.partyA.summary = "";

      const sel = document.getElementById("companyASelect") as HTMLSelectElement | null;
      if (sel) sel.value = "";
      setInputValue("companyAAddressInput", "");
      setInputValue("companyAHpInput", "");
      setInputValue("partyANameInput", ""); // ✅ NEW
      setInputValue("partyANamePercentInput", "");
      refreshPartyPreviews();
      refreshReadonly();
    });
  }

  const chk = document.getElementById("showSuppliersCheckbox") as HTMLInputElement | null;
  if (chk) {
    chk.addEventListener("change", () => {
      uiState.partyB.useSuppliers = !!chk.checked;
      toggleHidden("partyBManualWrap", uiState.partyB.useSuppliers);
      toggleHidden("partyBSuppliersWrap", !uiState.partyB.useSuppliers);
      refreshReadonly();
    });
  }

  ["partyBCompanyNameInput", "partyBAddressInput", "partyBHpInput"].forEach(id => {
    const el = document.getElementById(id) as HTMLInputElement | null;
    if (!el) return;
    el.addEventListener("input", () => {
      uiState.partyB.manualCompany = getInputValue("partyBCompanyNameInput");
      uiState.partyB.manualAddress = getInputValue("partyBAddressInput");
      uiState.partyB.manualHp = getInputValue("partyBHpInput");
      refreshReadonly();
    });
  });

  const supplierTypeSelect = document.getElementById("supplierTypeSelect") as HTMLSelectElement | null;
  if (supplierTypeSelect) {
    supplierTypeSelect.addEventListener("change", () => {
      const type = getSelectValue("supplierTypeSelect");
      uiState.partyB.supplierType = type;
      uiState.partyB.supplierName = "";
      uiState.partyB.supplierAddress = "";
      setInputValue("supplierAddressInput", "");

      const filtered = suppliersAll
        .filter(s => (s.type || "").toString().trim() === type)
        .map(s => ({ value: s.title, label: s.title }))
        .sort((a, b) => a.label.localeCompare(b.label, "he"));

      fillSelectByIdPairs("supplierSelect", filtered, "— בחר/י ספק —");
      refreshReadonly();
    });
  }

  const supplierSelect = document.getElementById("supplierSelect") as HTMLSelectElement | null;
  if (supplierSelect) {
    supplierSelect.addEventListener("change", () => {
      const name = getSelectValue("supplierSelect");
      uiState.partyB.supplierName = name;

      const s = suppliersAll.find(x => x.title === name);
      uiState.partyB.supplierAddress = s ? (s.address || "") : "";
      setInputValue("supplierAddressInput", uiState.partyB.supplierAddress);
      refreshReadonly();
    });
  }

  const partyBAddBtn = document.getElementById("partyBAddBtn") as HTMLButtonElement | null;
  if (partyBAddBtn) {
    partyBAddBtn.addEventListener("click", () => {
      // Task 5: append instead of overwrite
      const newB = buildPartyBSummary();
      uiState.partyB.summary = uiState.partyB.summary ? uiState.partyB.summary + "\n" + newB : newB;
      // Task 5: accumulate supplier name (comma-separated) separately for cntSupllierName tag
      if (uiState.partyB.useSuppliers && uiState.partyB.supplierName) {
        committedSupplierNames = committedSupplierNames
          ? committedSupplierNames + ", " + uiState.partyB.supplierName
          : uiState.partyB.supplierName;
      }
      refreshPartyPreviews();
      refreshReadonly();
    });
  }

  const partyBClearBtn = document.getElementById("partyBClearBtn") as HTMLButtonElement | null;
  if (partyBClearBtn) {
    partyBClearBtn.addEventListener("click", () => {
      uiState.partyB.summary = "";
      committedSupplierNames = ""; // Task 5: reset supplier name accumulator
      uiState.partyB.manualCompany = "";
      uiState.partyB.manualAddress = "";
      uiState.partyB.manualHp = "";
      uiState.partyB.supplierType = "";
      uiState.partyB.supplierName = "";
      uiState.partyB.supplierAddress = "";

      setInputValue("partyBCompanyNameInput", "");
      setInputValue("partyBAddressInput", "");
      setInputValue("partyBHpInput", "");
      setInputValue("supplierAddressInput", "");

      const st = document.getElementById("supplierTypeSelect") as HTMLSelectElement | null;
      if (st) st.value = "";

      const ss = document.getElementById("supplierSelect") as HTMLSelectElement | null;
      if (ss) {
        ss.innerHTML = "";
        ss.append(new Option("— בחר/י סוג ספק קודם —", ""));
      }

      refreshPartyPreviews();
      refreshReadonly();
    });
  }
}

/* =========
   Load dropdowns + caches
   ========= */
async function loadLookups() {
  try {
    const token = await getGraphToken();
    const siteId = await getSiteId(token);

    setSelectDisabled("statusSelect", true, "— טוען סטטוס… —");
    const statusListId = await getListId(siteId, token, STATUS_LIST_DISPLAY_NAME, STATUS_LIST_SERVER_RELATIVE_URL);
    const statusValues = await getListItemsByField(siteId, statusListId, token, STATUS_FIELD_NAME);
    fillSelectById("statusSelect", statusValues, "— בחר/י סטטוס —");
    setSelectDisabled("statusSelect", false);

    setSelectDisabled("projectSelect", true, "— טוען פרויקטים… —");
    const projectsListId = await getListId(siteId, token, PROJECTS_LIST_DISPLAY_NAME, PROJECTS_LIST_SERVER_RELATIVE_URL);
    const projectValues = await getListItemsByField(siteId, projectsListId, token, PROJECTS_FIELD_NAME);
    fillSelectById("projectSelect", projectValues, "— בחר/י פרויקט —");
    setSelectDisabled("projectSelect", false);

    setSelectDisabled("templateSelect", true, "— טוען תבניות… —");
    const templatesListId = await getListId(siteId, token, TEMPLATES_LIST_DISPLAY_NAME, TEMPLATES_LIST_SERVER_RELATIVE_URL);
    const templateValues = await getListItemsByField(siteId, templatesListId, token, TEMPLATES_FIELD_NAME);
    fillSelectById("templateSelect", templateValues, "— בחר/י תבנית —");
    setSelectDisabled("templateSelect", false);

    setSelectDisabled("siteSelect", true, "— טוען אתרים… —");
    const sitesListId = await getListId(siteId, token, SITES_LIST_DISPLAY_NAME, SITES_LIST_SERVER_RELATIVE_URL);
    const siteValues = await getListItemsByField(siteId, sitesListId, token, SITES_FIELD_NAME);
    fillSelectById("siteSelect", siteValues, "— בחר/י אתר —");
    setSelectDisabled("siteSelect", false);

    setSelectDisabled("municipalitySelect", true, "— טוען רשויות… —");
    const munListId = await getListId(siteId, token, MUNICIPALITIES_LIST_DISPLAY_NAME, MUNICIPALITIES_LIST_SERVER_RELATIVE_URL);
    const munValues = await getListItemsByField(siteId, munListId, token, MUNICIPALITIES_FIELD_NAME);
    fillSelectById("municipalitySelect", munValues, "— בחר/י רשות —");
    setSelectDisabled("municipalitySelect", false);

    setSelectDisabled("costIndexTypeSelect", true, "— טוען סוגי מדדים… —");
    const idxListId = await getListId(siteId, token, INDEX_TYPES_LIST_DISPLAY_NAME, INDEX_TYPES_LIST_SERVER_RELATIVE_URL);
    const idxValues = await getListItemsByField(siteId, idxListId, token, INDEX_TYPES_FIELD_NAME);
    fillSelectById("costIndexTypeSelect", idxValues, "— בחר/י סוג מדד —");
    setSelectDisabled("costIndexTypeSelect", false);

    setSelectDisabled("companyASelect", true, "— טוען חברות… —");
    const companiesListId = await getListId(siteId, token, COMPANIES_LIST_DISPLAY_NAME, COMPANIES_LIST_SERVER_RELATIVE_URL);
    const companyItems = await getListItemsFields(siteId, companiesListId, token, [
      COMPANIES_FIELDS.title,
      COMPANIES_FIELDS.address,
      COMPANIES_FIELDS.hp
    ]);

    companiesByTitle = new Map<string, Company>();
    const companyTitles: string[] = [];
    companyItems.forEach(it => {
      const title = (it.fields?.[COMPANIES_FIELDS.title] ?? "").toString().trim();
      if (!title) return;
      const address = (it.fields?.[COMPANIES_FIELDS.address] ?? "").toString().trim();
      const hp = (it.fields?.[COMPANIES_FIELDS.hp] ?? "").toString().trim();
      companiesByTitle.set(title, { title, address, hp });
      companyTitles.push(title);
    });
    fillSelectById("companyASelect", Array.from(new Set(companyTitles)).sort((a, b) => a.localeCompare(b, "he")), "— בחר/י חברה —");
    setSelectDisabled("companyASelect", false);

    setSelectDisabled("supplierTypeSelect", true, "— טוען סוגי ספקים… —");
    const supTypeListId = await getListId(siteId, token, SUPPLIER_TYPES_LIST_DISPLAY_NAME, SUPPLIER_TYPES_LIST_SERVER_RELATIVE_URL);
    const supTypes = await getListItemsByField(siteId, supTypeListId, token, SUPPLIER_TYPES_FIELD_NAME);
    fillSelectById("supplierTypeSelect", supTypes, "— בחר/י סוג ספק —");
    setSelectDisabled("supplierTypeSelect", false);

    const suppliersListId = await getListId(siteId, token, SUPPLIERS_LIST_DISPLAY_NAME, SUPPLIERS_LIST_SERVER_RELATIVE_URL);
    const supplierItems = await getListItemsFields(siteId, suppliersListId, token, [
      SUPPLIERS_FIELDS.title,
      SUPPLIERS_FIELDS.address,
      SUPPLIERS_FIELDS.type
    ]);

    suppliersAll = supplierItems
      .map(it => ({
        title: (it.fields?.[SUPPLIERS_FIELDS.title] ?? "").toString().trim(),
        address: (it.fields?.[SUPPLIERS_FIELDS.address] ?? "").toString().trim(),
        type: (it.fields?.[SUPPLIERS_FIELDS.type] ?? "").toString().trim()
      }))
      .filter(s => !!s.title);

    const supplierSelect = document.getElementById("supplierSelect") as HTMLSelectElement | null;
    if (supplierSelect) {
      supplierSelect.innerHTML = "";
      supplierSelect.append(new Option("— בחר/י סוג ספק קודם —", ""));
    }

  } catch (e: any) {
    console.error("Lookups load error:", e);
    alert("לא ניתן לטעון נתוני רשימות (סטטוס/פרויקט/תבניות/אתרים/חברות/ספקים/מדדים). בדקי הרשאות ושמות.");
  }
}

/* =========
   File identity helpers
   ========= */
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

function toShareIdFromWebUrl(webUrl: string): string {
  const bytes = new TextEncoder().encode(webUrl);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  const b64 = btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return `u!${b64}`;
}

async function getListItemIdByWebUrl(token: string, webUrl: string): Promise<string | null> {
  const shareId = toShareIdFromWebUrl(webUrl);
  const data = await graph<{ id: string; listItem?: { id: string } }>(
    `/shares/${encodeURIComponent(shareId)}/driveItem?$expand=listItem`, token
  );
  return data.listItem?.id ?? null;
}

function getLibraryNameFromWebUrl(webUrl: string): string | null {
  try {
    const u = new URL(webUrl);
    const parts = u.pathname.replace(/^\/+/, "").split("/");
    if (SITE_IS_UNDER_SITES) {
      if (parts.length >= 3 && parts[0].toLowerCase() === "sites") return decodeURIComponent(parts[2] || "");
    } else {
      if (parts.length >= 2 && parts[0].toLowerCase() === SP_SITE_PATH.toLowerCase()) return decodeURIComponent(parts[1] || "");
    }
    return null;
  } catch { return null; }
}

function toDateInputValue(val: any): string {
  const s = (val ?? "").toString().trim();
  if (!s) return "";
  return s.length >= 10 ? s.substring(0, 10) : s; // YYYY-MM-DD
}

/* =========
   ✅ Apply all loaded fields into UI + uiState
   ========= */
function applyLoadedFieldsToUI(fields: Record<string, any>) {
  // ===== General =====
  const contractNumber = pickField(fields, "ContractNumber");
  // Task 1: prefer the actual SP file version label over the manually-saved value
  const spVersion = pickField(fields, "_UIVersionString", "OData__UIVersionString");
  const contractVersion = spVersion || pickField(fields, "contractVersion");
  if (contractNumber) setInputValue("contractNumberInput", `${contractNumber}`);
  if (contractVersion) setInputValue("contractVersionInput", `${contractVersion}`);

  const template = pickField(fields, "ContractTemplate");
  const project = pickField(fields, "project");
  const site = pickField(fields, "SiteName");
  const municipality = pickField(fields, "Municipality");
  const workDescription = pickField(fields, "WorkDescription");
  const signDate = pickField(fields, "SignDate0");
  const startDate = pickField(fields, "StartDate");
  const months = pickField(fields, "DurationMonths");
  const expectedEnd = pickField(fields, "ExpectedEndDate");
  const status = pickField(fields, "status"); // ✅ כפי שאמרת

  if (template) setSelectValue("templateSelect", `${template}`);
  if (project) setSelectValue("projectSelect", `${project}`);
  if (site) setSelectValue("siteSelect", `${site}`);
  if (municipality) setSelectValue("municipalitySelect", `${municipality}`);
  if (workDescription) setInputValue("workDescriptionInput", `${workDescription}`);

  if (signDate) setInputValue("signDateInput", toDateInputValue(signDate));
  if (startDate) setInputValue("startDateInput", toDateInputValue(startDate));
  if (months) setInputValue("monthsInput", `${months}`);
  if (expectedEnd) setInputValue("expectedEndDateInput", toDateInputValue(expectedEnd));
  if (status) setSelectValue("statusSelect", `${status}`);

  // ===== Party A =====
  const partyACompanyName = pickField(fields, "PartyACompanyName");
  const partyACompanyAddress = pickField(fields, "PartyACompanyAddress");
  const partyACompanyHp = pickField(fields, "PartyACompanyHp");
  const partyAContactNamePercent = pickField(fields, "PartyAContactNamePercent");

  // ✅ NEW: Party A Name from library
  const partyAName = pickField(fields, "partyAName");

  if (partyACompanyName) {
    uiState.partyA.companyName = `${partyACompanyName}`.trim();
    setSelectValue("companyASelect", uiState.partyA.companyName);

    const fromCache = companiesByTitle.get(uiState.partyA.companyName);
    uiState.partyA.address = `${partyACompanyAddress ?? fromCache?.address ?? ""}`.trim();
    uiState.partyA.hp = `${partyACompanyHp ?? fromCache?.hp ?? ""}`.trim();

    setInputValue("companyAAddressInput", uiState.partyA.address);
    setInputValue("companyAHpInput", uiState.partyA.hp);
  } else {
    if (partyACompanyAddress) uiState.partyA.address = `${partyACompanyAddress}`.trim();
    if (partyACompanyHp) uiState.partyA.hp = `${partyACompanyHp}`.trim();
    setInputValue("companyAAddressInput", uiState.partyA.address);
    setInputValue("companyAHpInput", uiState.partyA.hp);
  }

  if (partyAName) {
    uiState.partyA.contactName = `${partyAName}`.trim();
    setInputValue("partyANameInput", uiState.partyA.contactName);
  }

  if (partyAContactNamePercent) {
    uiState.partyA.namePercent = `${partyAContactNamePercent}`.trim();
    setInputValue("partyANamePercentInput", uiState.partyA.namePercent);
  }

  uiState.partyA.summary = buildPartyASummary();

  // ===== Party B =====
  const partyBMode = pickField(fields, "PartyBMode"); // "Supplier" / "Manual"
  const partyBManualCompanyName = pickField(fields, "PartyBManualCompanyName");
  const partyBManualAddress = pickField(fields, "PartyBManualAddress");
  const partyBManualHp = pickField(fields, "PartyBManualHp");

  const partyBSupplierType = pickField(fields, "PartyBSupplierType");
  const partyBSupplierName = pickField(fields, "PartyBSupplierName");
  const partyBSupplierAddress = pickField(fields, "PartyBSupplierAddress");

  const useSuppliers = `${partyBMode ?? ""}`.toLowerCase() === "supplier";
  uiState.partyB.useSuppliers = useSuppliers;

  const chk = document.getElementById("showSuppliersCheckbox") as HTMLInputElement | null;
  if (chk) chk.checked = useSuppliers;
  toggleHidden("partyBManualWrap", useSuppliers);
  toggleHidden("partyBSuppliersWrap", !useSuppliers);

  if (!useSuppliers) {
    uiState.partyB.manualCompany = `${partyBManualCompanyName ?? ""}`.trim();
    uiState.partyB.manualAddress = `${partyBManualAddress ?? ""}`.trim();
    uiState.partyB.manualHp = `${partyBManualHp ?? ""}`.trim();

    setInputValue("partyBCompanyNameInput", uiState.partyB.manualCompany);
    setInputValue("partyBAddressInput", uiState.partyB.manualAddress);
    setInputValue("partyBHpInput", uiState.partyB.manualHp);

    uiState.partyB.supplierType = "";
    uiState.partyB.supplierName = "";
    uiState.partyB.supplierAddress = "";
    setSelectValue("supplierTypeSelect", "");
    setSelectValue("supplierSelect", "");
    setInputValue("supplierAddressInput", "");
  } else {
    uiState.partyB.supplierType = `${partyBSupplierType ?? ""}`.trim();
    uiState.partyB.supplierName = `${partyBSupplierName ?? ""}`.trim();
    uiState.partyB.supplierAddress = `${partyBSupplierAddress ?? ""}`.trim();
    committedSupplierNames = uiState.partyB.supplierName; // Task 5: restore accumulator from saved value

    setSelectValue("supplierTypeSelect", uiState.partyB.supplierType);

    const filtered = suppliersAll
      .filter(s => (s.type || "").toString().trim() === uiState.partyB.supplierType)
      .map(s => ({ value: s.title, label: s.title }))
      .sort((a, b) => a.label.localeCompare(b.label, "he"));

    fillSelectByIdPairs("supplierSelect", filtered, "— בחר/י ספק —");
    setSelectValue("supplierSelect", uiState.partyB.supplierName);

    if (!uiState.partyB.supplierAddress && uiState.partyB.supplierName) {
      const s = suppliersAll.find(x => x.title === uiState.partyB.supplierName);
      uiState.partyB.supplierAddress = s ? (s.address || "") : "";
    }
    setInputValue("supplierAddressInput", uiState.partyB.supplierAddress);

    uiState.partyB.manualCompany = "";
    uiState.partyB.manualAddress = "";
    uiState.partyB.manualHp = "";
    setInputValue("partyBCompanyNameInput", "");
    setInputValue("partyBAddressInput", "");
    setInputValue("partyBHpInput", "");
  }

  uiState.partyB.summary = buildPartyBSummary();

  const recipient = pickField(fields, "recipient");
  const otherSides = pickField(fields, "otherSides");
  if (recipient) uiState.partyA.summary = `${recipient}`.trim();
  if (otherSides) uiState.partyB.summary = `${otherSides}`.trim();

  // ===== Cost (אם העמודות קיימות בספרייה) =====
  uiState.cost.compMethod = `${pickField(fields, "CostCompMethod") ?? ""}`.trim();
  uiState.cost.contractScope = `${pickField(fields, "CostContractScope") ?? ""}`.trim();
  uiState.cost.currency = `${pickField(fields, "CostCurrency") ?? ""}`.trim();
  uiState.cost.indexType = `${pickField(fields, "CostIndexType") ?? ""}`.trim();
  uiState.cost.baseIndexDate = toDateInputValue(pickField(fields, "CostBaseIndexDate"));
  uiState.cost.indexMode = `${pickField(fields, "CostIndexMode") ?? ""}`.trim();
  uiState.cost.indexPoints = `${pickField(fields, "CostIndexPoints") ?? ""}`.trim();
  uiState.cost.paymentTerms = `${pickField(fields, "CostPaymentTerms") ?? ""}`.trim();

  if (uiState.cost.compMethod) setSelectValue("costCompMethodSelect", uiState.cost.compMethod);
  if (uiState.cost.contractScope) setInputValue("costContractScopeInput", uiState.cost.contractScope);
  if (uiState.cost.currency) setSelectValue("costCurrencySelect", uiState.cost.currency);
  if (uiState.cost.indexType) setSelectValue("costIndexTypeSelect", uiState.cost.indexType);
  if (uiState.cost.baseIndexDate) setInputValue("costBaseIndexDateInput", uiState.cost.baseIndexDate);
  if (uiState.cost.indexMode) setCostIndexMode(uiState.cost.indexMode);
  if (uiState.cost.indexPoints) setInputValue("costIndexPointsInput", uiState.cost.indexPoints);
  if (uiState.cost.paymentTerms) setSelectValue("costPaymentTermsSelect", uiState.cost.paymentTerms);

  uiState.cost.summary = buildCostSummary();

  // Custom fields 1..8 from library columns customField1..customField8
  (["1","2","3","4","5","6","7","8"] as const).forEach(n => {
    const key = `customField${n}`;
    const val = pickField(fields, key);
    if (val !== undefined && val !== null) {
      (uiState.custom as any)[key] = `${val}`.trim();
      setInputValue(`${key}Input`, (uiState.custom as any)[key]);
    }
  });

  // ===== Refresh UI =====
  refreshPartyPreviews();
  refreshCostPreview();
  refreshReadonly();
}

/* =========
   ✅ Load from library -> loads ALL fields (per saveToHelper)
   ========= */
async function loadFromLibraryIntoUI() {
  try {
    const absUrl = await getCurrentDocumentUrl();
    if (!absUrl) return;

    const token = await getGraphToken();
    const shareId = toShareIdFromWebUrl(absUrl);

    const data = await graph<{
      listItem?: { fields?: Record<string, any> };
    }>(
      `/shares/${encodeURIComponent(shareId)}/driveItem?$expand=listItem($expand=fields)`,
      token
    );

    const fields = data?.listItem?.fields;
    if (!fields) return;

    applyLoadedFieldsToUI(fields);
  } catch (e) {
    console.warn("loadFromLibraryIntoUI failed:", e);
  }
}

/* =========
   Tasks 2 / 3 / 4 – new-document defaults
   ========= */

// Task 2: if templateSelect is still empty, read cntTemplateName CC from the document
async function tryFillTemplateFromDocument(): Promise<void> {
  if (getSelectValue("templateSelect")) return;
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(TAGS.template);
      ccs.load("items/text");
      await context.sync();
      console.log("[tryFillTemplateFromDocument] CC count:", ccs.items.length);
      if (ccs.items.length > 0) {
        const raw = (ccs.items[0].text || "").trim();
        console.log("[tryFillTemplateFromDocument] raw CC text:", JSON.stringify(raw));

        // skip placeholder values like [תבנית]
        if (!raw || (raw.startsWith("[") && raw.endsWith("]"))) {
          console.log("[tryFillTemplateFromDocument] skipping placeholder or empty value");
          return;
        }

        const normalize = (s: string) =>
          s.trim()
           .toLowerCase()
           .replace(/\u00a0/g, " ")   // non-breaking space → regular space
           .replace(/[\r\n]+/g, " ")  // line breaks → space
           .replace(/\s+/g, " ");     // multiple spaces → single space

        const normalizedRaw = normalize(raw);
        console.log("[tryFillTemplateFromDocument] normalized CC text:", JSON.stringify(normalizedRaw));

        const sel = document.getElementById("templateSelect") as HTMLSelectElement | null;
        if (!sel) { console.warn("[tryFillTemplateFromDocument] templateSelect not found"); return; }

        const allOptions = Array.from(sel.options).filter(o => o.value);
        console.log("[tryFillTemplateFromDocument] available options:",
          allOptions.map(o => ({ value: o.value, text: o.text }))
        );

        const match = allOptions.find(o =>
          normalize(o.value) === normalizedRaw ||
          normalize(o.text)  === normalizedRaw
        );

        if (match) {
          console.log("[tryFillTemplateFromDocument] matched option:", { value: match.value, text: match.text });
          setSelectValue("templateSelect", match.value);
          refreshReadonly();
        } else {
          console.warn("[tryFillTemplateFromDocument] no matching option found for:", JSON.stringify(normalizedRaw));
        }
      } else {
        console.warn("[tryFillTemplateFromDocument] no CC with tag", TAGS.template, "found in document");
      }
    });
  } catch (e) {
    console.error("[tryFillTemplateFromDocument] error:", e);
  }
}

// Task 3: if contractNumberInput is still empty, use the document file name from the Office API URL
async function tryFillContractNumberFromFileName(): Promise<void> {
  if (getInputValue("contractNumberInput")) return;
  try {
    const url = await getCurrentDocumentUrl();
    if (!url) return;
    const raw = url.split("/").pop() || "";
    const decoded = decodeURIComponent(raw.split("?")[0]);
    // strip file extension (e.g. ".docx")
    const name = decoded.replace(/\.[^.]+$/, "");
    if (name) {
      setInputValue("contractNumberInput", name);
      refreshReadonly();
    }
  } catch { /* ignore */ }
}

// Task 4: if statusSelect is still empty, default to "חדש" (only if it exists in the loaded options)
function tryFillDefaultStatus(): void {
  if (getSelectValue("statusSelect")) return;
  const sel = document.getElementById("statusSelect") as HTMLSelectElement | null;
  if (!sel) return;
  const exists = Array.from(sel.options).some(o => o.value === "חדש");
  if (exists) {
    setSelectValue("statusSelect", "חדש");
    refreshReadonly();
  }
}

/* =========
   Save to Helper list
   ========= */
async function saveToHelper(fields: Record<string, any>): Promise<string> {
  const token = await getGraphToken();
  const siteId = await getSiteId(token);
  const helperListId = await getListId(siteId, token, HELPER_LIST_DISPLAY_NAME, HELPER_LIST_SERVER_RELATIVE_URL);
  const created = await createListItem(siteId, helperListId, token, fields);
  return created.id;
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

// ================================
// ✅ Fields tab: UI wiring
// ================================
function renderFieldsList(filterText = "") {
  const root = document.getElementById("fieldsList");
  if (!root) return;

  const q = (filterText || "").trim().toLowerCase();
  root.innerHTML = "";

  const list = FIELD_CATALOG.filter(f =>
    !q ||
    f.label.toLowerCase().includes(q) ||
    f.tag.toLowerCase().includes(q)
  );

  list.forEach((f) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "mini-btn primary field-pill";
    btn.style.textAlign = "right";
    btn.style.display = "flex";
    btn.style.justifyContent = "space-between";
    btn.style.gap = "10px";
    btn.style.alignItems = "center";

    const left = document.createElement("span");
    left.textContent = f.label;

    const right = document.createElement("code");
    right.textContent = f.tag;
    right.style.opacity = "0.75";
    right.style.fontWeight = "700";
    right.style.fontSize = "0.85rem";

    btn.appendChild(left);
    btn.appendChild(right);

    btn.addEventListener("click", async () => {
      try {
        await insertFieldAtCursor(f.tag, f.label);
        const lbl = document.getElementById("item-subject");
        if (lbl) lbl.textContent = `נוסף שדה: ${f.label}`;
      } catch (e: any) {
        console.error("insertFieldAtCursor failed:", e);
        alert("שגיאה בהוספת שדה למסמך: " + (e?.message || "לא ידועה"));
      }
    });

    root.appendChild(btn);
  });

  if (!list.length) {
    const p = document.createElement("div");
    p.className = "muted";
    p.textContent = "לא נמצאו שדות.";
    root.appendChild(p);
  }
}

function wireFieldsTabUI() {

  const listRoot = document.getElementById("fieldsList");
  if (!listRoot) return;

  const inp = document.getElementById("fieldSearchInput") as HTMLInputElement | null;
  if (inp) inp.addEventListener("input", () => renderFieldsList(inp.value || ""));

  const insertAllBtn = document.getElementById("insertAllFieldsBtn") as HTMLButtonElement | null;
  if (insertAllBtn) {
    insertAllBtn.addEventListener("click", async () => {
      try {
        for (const f of FIELD_CATALOG) {
          await insertFieldAtCursor(f.tag, f.label);
        }
        const lbl = document.getElementById("item-subject");
        if (lbl) lbl.textContent = "כל השדות נוספו למסמך.";
      } catch (e: any) {
        console.error("insertAllFields failed:", e);
        alert("שגיאה בהוספת כל השדות: " + (e?.message || "לא ידועה"));
      }
    });
  }

  const dbgBtn = document.getElementById("debugCCBtn") as HTMLButtonElement | null;
  if (dbgBtn) dbgBtn.addEventListener("click", () => debugListContentControls());

  renderFieldsList("");
}

/* =========
   Word update
   ========= */
// ================================
// ✅ Fields tab: insert CC at cursor
// ================================
async function insertFieldAtCursor(tag: string, label: string) {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    const cc = range.insertContentControl();

    cc.tag = tag;     // internal name
    cc.title = label; // Hebrew display

    // @ts-ignore
    cc.appearance = "BoundingBox";

    // initial value
    cc.insertText(`[${label}]`, Word.InsertLocation.replace);
    cc.font.color = "#00B0F0";

    await context.sync();
  });
}

async function updateDocumentFields(fields: Record<string, string | undefined>) {
  await Word.run(async (context) => {

    const safeEntries = Object.entries(fields)
      .filter(([tag, val]) => !!tag && tag !== "undefined" && tag !== "null" && val !== undefined);

    const tags = safeEntries.map(([tag]) => tag);

    const collections = tags.map(tag => {
      const col = context.document.contentControls.getByTag(tag);
      col.load("items");
      return { tag, col };
    });

    await context.sync();

    const missing: string[] = [];
    let updatedCount = 0;

    for (const { tag, col } of collections) {
      const val = (fields[tag] ?? "").toString();

      // אם אין ערך – לא מעדכנים (כמו אצלך היום)
      if (!val) continue;

      if (col.items.length > 0) {
        col.items.forEach(cc => cc.insertText(val, Word.InsertLocation.replace));
        updatedCount += col.items.length;
      } else {
        missing.push(tag);
      }
    }

    await context.sync();

    const lbl = document.getElementById("item-subject");
    if (lbl) {
      if (missing.length) {
        lbl.textContent =
          `עודכן (${updatedCount}). חסרים במסמך: ${missing.join(", ")}`;
      } else {
        lbl.textContent = `המסמך עודכן (${updatedCount}).`;
      }
    }

    if (missing.length) console.warn("Missing tags (not created):", missing);
  });
}

/* =========
   Buttons
   ========= */
export async function runUpdateDoc() {
  const contractNumber = getInputValue("contractNumberInput");
  const contractVersion = getInputValue("contractVersionInput");

  const template = getSelectValue("templateSelect");
  const project = getSelectValue("projectSelect");
  const site = getSelectValue("siteSelect");

  const municipality = getSelectValue("municipalitySelect");
  const workDescription = getInputValue("workDescriptionInput");

  const signDate = getInputValue("signDateInput");
  const startDate = getInputValue("startDateInput");
  const months = getInputValue("monthsInput");
  const expectedEndDate = getInputValue("expectedEndDateInput");

  const status = getSelectValue("statusSelect");

  const recipient = uiState.partyA.summary || "";
  const otherSides = uiState.partyB.summary || "";

  // לוודא cost.summary מעודכן
  if (!uiState.cost.summary) {
    refreshCostStateFromUI();
    uiState.cost.summary = buildCostSummary();
    refreshCostPreview();
  }

  // ==== TODO mappings from existing panel fields ====
  const localAuth = municipality;                       // cntLocalAuth <- רשות מקומית
  const tzadAPercent = uiState.partyA.namePercent || ""; // cmtTzadAPercent <- צד א שם ואחוז
  const madadTypeTitle = uiState.cost.indexType || "";   // cntMadadTypeTitle <- סוג מדד
  const isKnownTitle = uiState.cost.indexMode || "";     // cntIsKnownTitle <- מדד בגין/ידוע
  const madadBase = uiState.cost.baseIndexDate || "";    // cntMadadBase <- מדד בסיס (אצלך תאריך)
  const madadPoints = uiState.cost.indexPoints || "";    // cntMadadPoints <- נקודות מדד
  const jobDesc = workDescription;                       // cntJobDesc <- תיאור עבודה

  try {
    const fieldMap: Record<string, string | undefined> = {
      [TAGS.contractNumber]: contractNumber,
      [TAGS.contractVersion]: contractVersion,

      [TAGS.template]: template,
      [TAGS.project]: project,
      [TAGS.site]: site,

      [TAGS.municipality]: municipality,
      [TAGS.workDescription]: workDescription,

      [TAGS.signDate]: signDate,
      [TAGS.startDate]: startDate,
      [TAGS.months]: months,
      [TAGS.expectedEndDate]: expectedEndDate,

      [TAGS.status]: status,

      [TAGS.supplier]: committedSupplierNames || uiState.partyB.supplierName, // Task 5

      [TAGS.recipient]: recipient,
      [TAGS.otherSides]: otherSides,

      [TAGS.costCompMethod]: uiState.cost.compMethod,
      [TAGS.costContractScope]: uiState.cost.contractScope,
      [TAGS.costCurrency]: uiState.cost.currency,
      [TAGS.costIndexType]: uiState.cost.indexType,
      [TAGS.costBaseIndexDate]: uiState.cost.baseIndexDate,
      [TAGS.costIndexMode]: uiState.cost.indexMode,
      [TAGS.costIndexPoints]: uiState.cost.indexPoints,
      [TAGS.costPaymentTerms]: uiState.cost.paymentTerms,

      // ✅ TODO fields – same panel inputs
      [TAGS.localAuth]: localAuth,
      [TAGS.tzadAPercent]: tzadAPercent,
      [TAGS.madadTypeTitle]: madadTypeTitle,
      [TAGS.isKnownTitle]: isKnownTitle,
      [TAGS.madadBase]: madadBase,
      [TAGS.madadPoints]: madadPoints,
      [TAGS.jobDesc]: jobDesc,

      [TAGS.partyAName]: uiState.partyA.contactName || undefined,

      [TAGS.customField1]: uiState.custom.customField1 || undefined,
      [TAGS.customField2]: uiState.custom.customField2 || undefined,
      [TAGS.customField3]: uiState.custom.customField3 || undefined,
      [TAGS.customField4]: uiState.custom.customField4 || undefined,
      [TAGS.customField5]: uiState.custom.customField5 || undefined,
      [TAGS.customField6]: uiState.custom.customField6 || undefined,
      [TAGS.customField7]: uiState.custom.customField7 || undefined,
      [TAGS.customField8]: uiState.custom.customField8 || undefined,
    };

    await updateDocumentFields(fieldMap);

    const lbl = document.getElementById("item-subject");
    if (lbl) lbl.textContent = "המסמך עודכן בהצלחה.";
  } catch (e: any) {
    console.error("runUpdateDoc error:", e);
    alert("שגיאה בעדכון המסמך: " + (e?.message || "לא ידועה"));
  }
}

export async function runSaveSystem() {
  refreshCostStateFromUI();
  if (!uiState.cost.summary) uiState.cost.summary = buildCostSummary();

  const contractNumber = getInputValue("contractNumberInput");
  const contractVersion = getInputValue("contractVersionInput");
  const template = getSelectValue("templateSelect");
  const project = getSelectValue("projectSelect");
  const site = getSelectValue("siteSelect");
  const municipality = getSelectValue("municipalitySelect");
  const workDescription = getInputValue("workDescriptionInput");
  const signDate = getInputValue("signDateInput");
  const startDate = getInputValue("startDateInput");
  const months = getInputValue("monthsInput");
  const expectedEndDate = getInputValue("expectedEndDateInput");
  const status = getSelectValue("statusSelect");

  const partyACompanyName = uiState.partyA.companyName;
  const partyACompanyAddress = uiState.partyA.address;
  const partyACompanyHp = uiState.partyA.hp;

  const partyAName = uiState.partyA.contactName;

  const partyAContactNamePercent = uiState.partyA.namePercent;

  const partyBMode = uiState.partyB.useSuppliers ? "Supplier" : "Manual";
  const partyBManualCompanyName = uiState.partyB.manualCompany;
  const partyBManualAddress = uiState.partyB.manualAddress;
  const partyBManualHp = uiState.partyB.manualHp;

  const partyBSupplierType = uiState.partyB.supplierType;
  const partyBSupplierName = uiState.partyB.supplierName;
  const partyBSupplierAddress = uiState.partyB.supplierAddress;

  const recipient = uiState.partyA.summary || "";
  const otherSides = uiState.partyB.summary || "";

  try {
    const absUrl = await getCurrentDocumentUrl();
    if (!absUrl) { alert("לא מזוהה כתובת למסמך. שמרי את המסמך ב-SharePoint ונסי שוב."); return; }

    const token = await getGraphToken();
    const itemId = await getListItemIdByWebUrl(token, absUrl);
    if (!itemId) { alert("לא ניתן להביא את מזהה הפריט של המסמך (List Item ID)."); return; }

    const libName = getLibraryNameFromWebUrl(absUrl) || "";

    await saveToHelper({
      Title: itemId,
      libName: libName || undefined,

      ContractNumber: contractNumber || undefined,
      contractVersion: contractVersion || undefined,
      ContractTemplate: template || undefined,
      project: project || undefined,
      siteName: site || undefined,
      Municipality: municipality || undefined,
      WorkDescription: workDescription || undefined,
      signDate: signDate || undefined,
      StartDate: startDate || undefined,
      DurationMonths: months || undefined,
      ExpectedEndDate: expectedEndDate || undefined,
      status: status || undefined,
      supplierName: (committedSupplierNames || uiState.partyB.supplierName) || undefined, // Task 5


      PartyACompanyName: partyACompanyName || undefined,
      PartyACompanyAddress: partyACompanyAddress || undefined,
      PartyACompanyHp: partyACompanyHp || undefined,
      partyAName: partyAName || undefined,

      PartyAContactNamePercent: partyAContactNamePercent || undefined,

      PartyBMode: partyBMode || undefined,
      PartyBManualCompanyName: partyBManualCompanyName || undefined,
      PartyBManualAddress: partyBManualAddress || undefined,
      PartyBManualHp: partyBManualHp || undefined,
      PartyBSupplierType: partyBSupplierType || undefined,
      PartyBSupplierName: partyBSupplierName || undefined,
      PartyBSupplierAddress: partyBSupplierAddress || undefined,

      recipient: recipient || undefined,
      otherSides: otherSides || undefined,

      CostCompMethod: uiState.cost.compMethod || undefined,
      CostContractScope: uiState.cost.contractScope || undefined,
      CostCurrency: uiState.cost.currency || undefined,
      CostIndexType: uiState.cost.indexType || undefined,
      CostBaseIndexDate: uiState.cost.baseIndexDate || undefined,
      CostIndexMode: uiState.cost.indexMode || undefined,
      CostIndexPoints: uiState.cost.indexPoints || undefined,
      CostPaymentTerms: uiState.cost.paymentTerms || undefined,

      // ✅ NEW: Custom fields 1..8 saved to helper (and then to library by flow)
      customField1: uiState.custom.customField1 || undefined,
      customField2: uiState.custom.customField2 || undefined,
      customField3: uiState.custom.customField3 || undefined,
      customField4: uiState.custom.customField4 || undefined,
      customField5: uiState.custom.customField5 || undefined,
      customField6: uiState.custom.customField6 || undefined,
      customField7: uiState.custom.customField7 || undefined,
      customField8: uiState.custom.customField8 || undefined,
    });

    await Word.run(async (ctx) => { await ctx.document.save(); });
    showCloseDocMessage();
  } catch (e: any) {
    console.error("runSaveSystem error:", e);
    alert("שגיאה בשמירה במערכת: " + (e?.message || "לא ידועה"));
  }
}

export async function debugListContentControls() {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag,title");
    await context.sync();

    const rows = ccs.items.map((cc) => ({
      tag: (cc.tag || "").toString(),
      title: (cc.title || "").toString(),
    }));

    console.table(rows);

    const uniqTags = Array.from(new Set(rows.map(r => r.tag).filter(Boolean))).sort();
    const uniqTitles = Array.from(new Set(rows.map(r => r.title).filter(Boolean))).sort();

    console.log("Unique TAGS:", uniqTags);
    console.log("Unique TITLES:", uniqTitles);
  });
}


function wireExtrasUI() {
  const clearBtn = document.getElementById("extrasClearBtn") as HTMLButtonElement | null;
  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      for (let i = 1; i <= 8; i++) {
        const key = `customField${i}` as const;
        (uiState.custom as any)[key] = "";                 
        setInputValue(`${key}Input`, "");
      }
      refreshReadonly();
    });
  }

  for (let i = 1; i <= 8; i++) {
    const key = `customField${i}` as const;
    const id = `${key}Input`;
    const el = document.getElementById(id) as HTMLInputElement | null;
    if (!el) continue;

    el.addEventListener("input", () => {
      (uiState.custom as any)[key] = getInputValue(id);     
      refreshReadonly();
    });
  }
}


/* =========
   Bootstrap
   ========= */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {

    console.log("******* VERS 17 ************");

    (document.getElementById("sideload-msg") as HTMLElement).style.display = "none";
    (document.getElementById("app-body") as HTMLElement).style.display = "block";

    const btnUpdate = document.getElementById("runUpdateDoc");
    if (btnUpdate) (btnUpdate as HTMLDivElement).onclick = runUpdateDoc;

    const btnSave = document.getElementById("runSaveSystem");
    if (btnSave) (btnSave as HTMLDivElement).onclick = runSaveSystem;

    wireDates();
    wirePartiesUI();
    wireCostUI();
    wireExtrasUI();
    wireFieldsTabUI();

    [
      "contractNumberInput", "contractVersionInput", "templateSelect", "projectSelect", "siteSelect",
      "municipalitySelect", "workDescriptionInput", "signDateInput", "startDateInput", "monthsInput",
      "statusSelect",

      "partyANameInput",
      "customField1Input", "customField2Input", "customField3Input", "customField4Input",
      "customField5Input", "customField6Input", "customField7Input", "customField8Input"
    ].forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener("change", refreshReadonly);
      el.addEventListener("input", refreshReadonly);
    });

    loadLookups().then(async () => {
      refreshPartyPreviews();
      refreshCostStateFromUI();
      refreshCostPreview();
      refreshReadonly();

      // ✅ Load ALL fields from the library item
      await loadFromLibraryIntoUI();

      // Tasks 2 / 3 / 4: fill defaults only if the above left them empty (= new document)
      await tryFillTemplateFromDocument();
      await tryFillContractNumberFromFileName();
      tryFillDefaultStatus();

      debugListContentControls();
    });
  } else {
    (document.getElementById("sideload-msg") as HTMLElement).style.display = "block";
    (document.getElementById("app-body") as HTMLElement).style.display = "none";
  }
});
