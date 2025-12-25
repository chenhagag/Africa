import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";

/* =========
   Config
   ========= */
const SP_HOSTNAME = "africaisrael.sharepoint.com";
const SP_SITE_PATH = "ContractsNew";
const SITE_IS_UNDER_SITES = false;

/* =========
   LOOKUP LISTS (עדכני כאן לפי שמות אצלכם)
   ========= */

// Status dropdown source
const STATUS_LIST_DISPLAY_NAME = "ContractStatus";
const STATUS_FIELD_NAME = "Title";

// Projects dropdown source
const PROJECTS_LIST_DISPLAY_NAME = "projects";
const PROJECTS_FIELD_NAME = "Title";

// Templates dropdown source (תבניות)
const TEMPLATES_LIST_DISPLAY_NAME = "ContractTemplates"; // <-- עדכני לפי הרשימה בפועל
const TEMPLATES_FIELD_NAME = "Title";

// Sites dropdown source (אתרים)
const SITES_LIST_DISPLAY_NAME = "Sites"; // <-- עדכני לפי הרשימה בפועל
const SITES_FIELD_NAME = "Title";

// Municipalities dropdown source (רשויות מקומיות)
const MUNICIPALITIES_LIST_DISPLAY_NAME = "רשויות מקומיות"; // <-- עדכני אם השם שונה
const MUNICIPALITIES_FIELD_NAME = "Title";

// Companies list (לצד א')
const COMPANIES_LIST_DISPLAY_NAME = "חברות"; // <-- עדכני לפי הרשימה בפועל
const COMPANIES_FIELDS = {
  title: "Title",
  address: "Address", // <-- internal name
  hp: "HP"            // <-- internal name (ח.פ)
};

// Suppliers types list
const SUPPLIER_TYPES_LIST_DISPLAY_NAME = "סוגי ספקים"; // <-- עדכני לפי הרשימה בפועל
const SUPPLIER_TYPES_FIELD_NAME = "Title";

// Suppliers list
const SUPPLIERS_LIST_DISPLAY_NAME = "Suppliers"; // כבר היה אצלך
const SUPPLIERS_FIELDS = {
  title: "Title",
  address: "Address",         // <-- internal name
  type: "SupplierType"        // <-- internal name שמחזיק את סוג הספק (טקסט/lookup title)
};

/* =========
   URLs for list resolving
   ========= */
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
const COMPANIES_LIST_SERVER_RELATIVE_URL = listUrl(COMPANIES_LIST_DISPLAY_NAME);
const SUPPLIER_TYPES_LIST_SERVER_RELATIVE_URL = listUrl(SUPPLIER_TYPES_LIST_DISPLAY_NAME);
const SUPPLIERS_LIST_SERVER_RELATIVE_URL = listUrl(SUPPLIERS_LIST_DISPLAY_NAME);

// Helper list target
const HELPER_LIST_DISPLAY_NAME = "FieldsUpdateHelper";
const HELPER_LIST_SERVER_RELATIVE_URL = SITE_IS_UNDER_SITES
  ? `/sites/${SP_SITE_PATH}/Lists/${HELPER_LIST_DISPLAY_NAME}`
  : `/${SP_SITE_PATH}/Lists/${HELPER_LIST_DISPLAY_NAME}`;

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

/* =========
   Local UI state (לצפייה בלבד + צדדים)
   ========= */
const uiState = {
  partyA: {
    companyName: "",
    address: "",
    hp: "",
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
  }
};

/* =========
   Lookup caches
   ========= */
type Company = { title: string; address: string; hp: string };
type Supplier = { title: string; address: string; type: string };

let companiesByTitle = new Map<string, Company>();
let suppliersAll: Supplier[] = [];

/* =========
   Date calc (מועד סיום צפוי)
   ========= */
function addMonthsToDateISO(startIso: string, months: number): string {
  if (!startIso || !months || months <= 0) return "";
  const [y, m, d] = startIso.split("-").map(Number);
  if (!y || !m || !d) return "";
  const dt = new Date(Date.UTC(y, m - 1, d));
  const targetMonth = dt.getUTCMonth() + months;
  dt.setUTCMonth(targetMonth);

  // שמירה כ-YYYY-MM-DD
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
   Parties behavior
   ========= */
function buildPartyASummary() {
  const lines: string[] = [];
  if (uiState.partyA.companyName) lines.push(`חברה: ${uiState.partyA.companyName}`);
  if (uiState.partyA.address) lines.push(`כתובת: ${uiState.partyA.address}`);
  if (uiState.partyA.hp) lines.push(`ח.פ: ${uiState.partyA.hp}`);
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

function refreshPartyPreviews() {
  const a = uiState.partyA.summary || "לא הוזן צד א' עדיין.";
  const b = uiState.partyB.summary || "לא הוזן צד ב' עדיין.";
  setText("partyAPreview", a);
  setText("partyBPreview", b);
  refreshReadonly();
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

  generalLines.push("פרטים כלליים:");
  if (contractNumber) generalLines.push(`• מספר חוזה: ${contractNumber}`);
  if (contractVersion) generalLines.push(`• גרסת חוזה: ${contractVersion}`);
  if (template) generalLines.push(`• תבנית: ${template}`);
  if (project) generalLines.push(`• פרויקט: ${project}`);
  if (site) generalLines.push(`• אתר: ${site}`);
  if (municipality) generalLines.push(`• רשות מקומית: ${municipality}`);
  if (workDesc) generalLines.push(`• תיאור עבודה: ${workDesc}`);
  if (signDate) generalLines.push(`• תאריך חתימה: ${signDate}`);
  if (startDate) generalLines.push(`• מועד התחלה: ${startDate}`);
  if (months) generalLines.push(`• מספר חודשים: ${months}`);
  if (expectedEnd) generalLines.push(`• מועד סיום צפוי: ${expectedEnd}`);
  if (status) generalLines.push(`• סטטוס: ${status}`);

  const partiesLines: string[] = [];
  partiesLines.push("צדדים בחוזה:");
  partiesLines.push("צד א':");
  partiesLines.push(uiState.partyA.summary ? uiState.partyA.summary : "—");
  partiesLines.push("");
  partiesLines.push("צד ב':");
  partiesLines.push(uiState.partyB.summary ? uiState.partyB.summary : "—");

  setText("readonlyGeneral", generalLines.join("\n"));
  setText("readonlyParties", partiesLines.join("\n"));
}

function wirePartiesUI() {
  // צד א' - בחירת חברה => מילוי כתובת וח.פ
  const companyASelect = document.getElementById("companyASelect") as HTMLSelectElement | null;
  if (companyASelect) {
    companyASelect.addEventListener("change", () => {
      const title = getSelectValue("companyASelect");
      const c = companiesByTitle.get(title);
      uiState.partyA.companyName = title;
      uiState.partyA.address = c?.address || "";
      uiState.partyA.hp = c?.hp || "";
      setInputValue("companyAAddressInput", uiState.partyA.address);
      setInputValue("companyAHpInput", uiState.partyA.hp);
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
      uiState.partyA.summary = buildPartyASummary();
      refreshPartyPreviews();
    });
  }

  const partyAClearBtn = document.getElementById("partyAClearBtn") as HTMLButtonElement | null;
  if (partyAClearBtn) {
    partyAClearBtn.addEventListener("click", () => {
      uiState.partyA = { companyName: "", address: "", hp: "", namePercent: "", summary: "" };
      // clear fields
      const sel = document.getElementById("companyASelect") as HTMLSelectElement | null;
      if (sel) sel.value = "";
      setInputValue("companyAAddressInput", "");
      setInputValue("companyAHpInput", "");
      setInputValue("partyANamePercentInput", "");
      refreshPartyPreviews();
    });
  }

  // צד ב' - toggle ספקים
  const chk = document.getElementById("showSuppliersCheckbox") as HTMLInputElement | null;
  if (chk) {
    chk.addEventListener("change", () => {
      uiState.partyB.useSuppliers = !!chk.checked;
      toggleHidden("partyBManualWrap", uiState.partyB.useSuppliers);
      toggleHidden("partyBSuppliersWrap", !uiState.partyB.useSuppliers);
      refreshReadonly();
    });
  }

  // צד ב' ידני
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

  // צד ב' ספקים - סוג ספק => סינון ספקים
  const supplierTypeSelect = document.getElementById("supplierTypeSelect") as HTMLSelectElement | null;
  if (supplierTypeSelect) {
    supplierTypeSelect.addEventListener("change", () => {
      const type = getSelectValue("supplierTypeSelect");
      uiState.partyB.supplierType = type;
      uiState.partyB.supplierName = "";
      uiState.partyB.supplierAddress = "";
      setInputValue("supplierAddressInput", "");

      // fill suppliers by type
      const filtered = suppliersAll
        .filter(s => (s.type || "").toString().trim() === type)
        .map(s => ({ value: s.title, label: s.title }))
        .sort((a, b) => a.label.localeCompare(b.label, "he"));

      fillSelectByIdPairs("supplierSelect", filtered, "— בחר/י ספק —");
      refreshReadonly();
    });
  }

  // בחירת ספק => שתילת כתובת
  const supplierSelect = document.getElementById("supplierSelect") as HTMLSelectElement | null;
  if (supplierSelect) {
    supplierSelect.addEventListener("change", () => {
      const name = getSelectValue("supplierSelect");
      uiState.partyB.supplierName = name;
      const s = suppliersAll.find(x => x.title === name);
      uiState.partyB.supplierAddress = s?.address || "";
      setInputValue("supplierAddressInput", uiState.partyB.supplierAddress);
      refreshReadonly();
    });
  }

  const partyBAddBtn = document.getElementById("partyBAddBtn") as HTMLButtonElement | null;
  if (partyBAddBtn) {
    partyBAddBtn.addEventListener("click", () => {
      uiState.partyB.summary = buildPartyBSummary();
      refreshPartyPreviews();
    });
  }

  const partyBClearBtn = document.getElementById("partyBClearBtn") as HTMLButtonElement | null;
  if (partyBClearBtn) {
    partyBClearBtn.addEventListener("click", () => {
      uiState.partyB.summary = "";
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
    });
  }
}

/* =========
   load dropdowns + caches
   ========= */
async function loadLookups() {
  console.log("******** VERS GENERAL+PARTIES **************");

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

    // Templates
    setSelectDisabled("templateSelect", true, "— טוען תבניות… —");
    const templatesListId = await getListId(siteId, token, TEMPLATES_LIST_DISPLAY_NAME, TEMPLATES_LIST_SERVER_RELATIVE_URL);
    const templateValues = await getListItemsByField(siteId, templatesListId, token, TEMPLATES_FIELD_NAME);
    fillSelectById("templateSelect", templateValues, "— בחר/י תבנית —");
    setSelectDisabled("templateSelect", false);

    // Sites
    setSelectDisabled("siteSelect", true, "— טוען אתרים… —");
    const sitesListId = await getListId(siteId, token, SITES_LIST_DISPLAY_NAME, SITES_LIST_SERVER_RELATIVE_URL);
    const siteValues = await getListItemsByField(siteId, sitesListId, token, SITES_FIELD_NAME);
    fillSelectById("siteSelect", siteValues, "— בחר/י אתר —");
    setSelectDisabled("siteSelect", false);

    // Municipalities
    setSelectDisabled("municipalitySelect", true, "— טוען רשויות… —");
    const munListId = await getListId(siteId, token, MUNICIPALITIES_LIST_DISPLAY_NAME, MUNICIPALITIES_LIST_SERVER_RELATIVE_URL);
    const munValues = await getListItemsByField(siteId, munListId, token, MUNICIPALITIES_FIELD_NAME);
    fillSelectById("municipalitySelect", munValues, "— בחר/י רשות —");
    setSelectDisabled("municipalitySelect", false);

    // Companies (Side A) - need address+hp cache
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
    fillSelectById("companyASelect", Array.from(new Set(companyTitles)).sort((a,b)=>a.localeCompare(b,"he")), "— בחר/י חברה —");
    setSelectDisabled("companyASelect", false);

    // Supplier Types
    setSelectDisabled("supplierTypeSelect", true, "— טוען סוגי ספקים… —");
    const supTypeListId = await getListId(siteId, token, SUPPLIER_TYPES_LIST_DISPLAY_NAME, SUPPLIER_TYPES_LIST_SERVER_RELATIVE_URL);
    const supTypes = await getListItemsByField(siteId, supTypeListId, token, SUPPLIER_TYPES_FIELD_NAME);
    fillSelectById("supplierTypeSelect", supTypes, "— בחר/י סוג ספק —");
    setSelectDisabled("supplierTypeSelect", false);

    // Suppliers cache (title+address+type)
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

    // supplierSelect starts with a hint
    const supplierSelect = document.getElementById("supplierSelect") as HTMLSelectElement | null;
    if (supplierSelect) {
      supplierSelect.innerHTML = "";
      supplierSelect.append(new Option("— בחר/י סוג ספק קודם —", ""));
    }

  } catch (e: any) {
    console.error("Lookups load error:", e);
    alert("לא ניתן לטעון נתוני רשימות (סטטוס/פרויקט/תבניות/אתרים/חברות/ספקים). בדקי הרשאות ושמות רשימות/עמודות.");
  }
}

/* =========
   Word actions (השארתי את הפונקציות, אבל כרגע שומרות רק recipient/otherSides + שדות כלליים)
   ========= */
async function updateDocumentFields(fields: {
  // General
  contractNumber?: string;
  contractVersion?: string;
  template?: string;
  project?: string;
  site?: string;
  municipality?: string;
  workDescription?: string;
  signDate?: string;
  startDate?: string;
  months?: string;
  expectedEndDate?: string;
  status?: string;

  // Parties summaries mapped to existing tags
  recipient?: string;   // צד א' (summary)
  otherSides?: string;  // צד ב' (summary)
}) {
  await Word.run(async (context) => {
    const tags = [
      "contractNumber",
      "contractVersion",
      "template",
      "project",
      "site",
      "municipality",
      "workDescription",
      "signDate",
      "startDate",
      "months",
      "expectedEndDate",
      "status",
      "recipient",
      "otherSides"
    ] as const;

    const valuesByTag: Record<string, string | undefined> = {
      contractNumber: fields.contractNumber,
      contractVersion: fields.contractVersion,
      template: fields.template,
      project: fields.project,
      site: fields.site,
      municipality: fields.municipality,
      workDescription: fields.workDescription,
      signDate: fields.signDate,
      startDate: fields.startDate,
      months: fields.months,
      expectedEndDate: fields.expectedEndDate,
      status: fields.status,
      recipient: fields.recipient,
      otherSides: fields.otherSides
    };

    const collections = tags.map(tag => {
      const col = context.document.contentControls.getByTag(tag);
      col.load("items");
      return { tag, col };
    });

    await context.sync();

    for (const { tag, col } of collections) {
      const val = (valuesByTag[tag] || "").toString();
      if (!val) continue;

      if (col.items.length > 0) {
        col.items.forEach(cc => cc.insertText(val, Word.InsertLocation.replace));
      } else {
        const range = context.document.getSelection();
        const cc = range.insertContentControl();
        cc.tag = tag;
        cc.title = tag;
        cc.insertText(val, Word.InsertLocation.replace);
      }
    }

    await context.sync();
  });
}

/* =========
   Helper list save (כאן שמרתי מינימום כדי לא לשבור את הסכימה הקיימת אצלך)
   ========= */
async function saveToHelper(fields: {
  titleItemId: string;
  libName?: string;

  // keep old columns (if exist)
  recipient: string;
  otherSides: string;
  status: string;

  // new-ish (אם קיימים אצלכם בעמודות helper – אפשר להוסיף; אם לא, אפשר להשאיר רק מה שיש)
  contractNumber?: string;
  contractVersion?: string;
  template?: string;
  project?: string;
  site?: string;
  municipality?: string;
  workDescription?: string;
  signDate?: string;
  startDate?: string;
  months?: string;
  expectedEndDate?: string;
}): Promise<string> {
  const token = await getGraphToken();
  const siteId = await getSiteId(token);
  const helperListId = await getListId(siteId, token, HELPER_LIST_DISPLAY_NAME, HELPER_LIST_SERVER_RELATIVE_URL);

  const helperFields: Record<string, any> = {
    Title: fields.titleItemId,
    recipient: fields.recipient,
    otherSides: fields.otherSides,
    status: fields.status
  };

  // נסיון לשמור גם שדות חדשים אם קיימים ברשימה (אם אין— Graph יחזיר שגיאה; במקרה כזה תסירי/נשנה לשמות נכונים)
  if (fields.contractNumber) helperFields.contractNumber = fields.contractNumber;
  if (fields.contractVersion) helperFields.contractVersion = fields.contractVersion;
  if (fields.template) helperFields.template = fields.template;
  if (fields.project) helperFields.project = fields.project;
  if (fields.site) helperFields.site = fields.site;
  if (fields.municipality) helperFields.municipality = fields.municipality;
  if (fields.workDescription) helperFields.workDescription = fields.workDescription;
  if (fields.signDate) helperFields.signDate = fields.signDate;
  if (fields.startDate) helperFields.startDate = fields.startDate;
  if (fields.months) helperFields.months = fields.months;
  if (fields.expectedEndDate) helperFields.expectedEndDate = fields.expectedEndDate;

  if (fields.libName) helperFields.libName = fields.libName;

  const created = await createListItem(siteId, helperListId, token, helperFields);
  return created.id;
}

/* =========
   Existing helpers for file identity + url
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

/* =========
   Buttons
   ========= */
export async function runUpdateDoc() {
  // general
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

  // parties (summaries)
  const recipient = uiState.partyA.summary || "";
  const otherSides = uiState.partyB.summary || "";

  if (![contractNumber, contractVersion, template, project, site, municipality, workDescription, signDate, startDate, months, expectedEndDate, status, recipient, otherSides].some(Boolean)) {
    alert("יש למלא לפחות שדה אחד לעדכון במסמך.");
    return;
  }

  try {
    await updateDocumentFields({
      contractNumber,
      contractVersion,
      template,
      project,
      site,
      municipality,
      workDescription,
      signDate,
      startDate,
      months,
      expectedEndDate,
      status,
      recipient,
      otherSides
    });

    const lbl = document.getElementById("item-subject");
    if (lbl) lbl.textContent = "המסמך עודכן בהצלחה.";
  } catch (e: any) {
    console.error("runUpdateDoc error:", e);
    alert("שגיאה בעדכון המסמך: " + (e?.message || "לא ידועה"));
  }
}

export async function runSaveSystem() {
  // general
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

  const recipient = uiState.partyA.summary || "נתון חסר";
  const otherSides = uiState.partyB.summary || "נתון חסר";

  try {
    await updateDocumentFields({
      contractNumber: contractNumber || undefined,
      contractVersion: contractVersion || undefined,
      template: template || undefined,
      project: project || undefined,
      site: site || undefined,
      municipality: municipality || undefined,
      workDescription: workDescription || undefined,
      signDate: signDate || undefined,
      startDate: startDate || undefined,
      months: months || undefined,
      expectedEndDate: expectedEndDate || undefined,
      status: status || undefined,
      recipient: recipient || undefined,
      otherSides: otherSides || undefined
    });

    const absUrl = await getCurrentDocumentUrl();
    if (!absUrl) { alert("לא מזוהה כתובת למסמך. שמרי את המסמך ב-SharePoint ונסי שוב."); return; }

    const token = await getGraphToken();
    const itemId = await getListItemIdByWebUrl(token, absUrl);
    if (!itemId) { alert("לא ניתן להביא את מזהה הפריט של המסמך (List Item ID)."); return; }

    const libName = getLibraryNameFromWebUrl(absUrl) || undefined;

    await saveToHelper({
      titleItemId: itemId,
      libName,
      recipient,
      otherSides,
      status: status || "נתון חסר",
      contractNumber: contractNumber || undefined,
      contractVersion: contractVersion || undefined,
      template: template || undefined,
      project: project || undefined,
      site: site || undefined,
      municipality: municipality || undefined,
      workDescription: workDescription || undefined,
      signDate: signDate || undefined,
      startDate: startDate || undefined,
      months: months || undefined,
      expectedEndDate: expectedEndDate || undefined
    });

    await Word.run(async (ctx) => { await ctx.document.save(); });

    showCloseDocMessage();
  } catch (e: any) {
    console.error("runSaveSystem error:", e);
    alert("שגיאה בשמירה במערכת: " + (e?.message || "לא ידועה"));
  }
}

/* =========
   Bootstrap
   ========= */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    (document.getElementById("sideload-msg") as HTMLElement).style.display = "none";
    // חשוב: block ולא flex, כדי שהפריסה (grid) תעבוד יציב
    (document.getElementById("app-body") as HTMLElement).style.display = "block";

    const btnUpdate = document.getElementById("runUpdateDoc");
    if (btnUpdate) (btnUpdate as HTMLDivElement).onclick = runUpdateDoc;

    const btnSave = document.getElementById("runSaveSystem");
    if (btnSave) (btnSave as HTMLDivElement).onclick = runSaveSystem;

    wireDates();
    wirePartiesUI();

    // refresh readonly on any general change
    [
      "contractNumberInput","contractVersionInput","templateSelect","projectSelect","siteSelect",
      "municipalitySelect","workDescriptionInput","signDateInput","startDateInput","monthsInput",
      "statusSelect"
    ].forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener("change", refreshReadonly);
      el.addEventListener("input", refreshReadonly);
    });

    loadLookups().then(() => {
      refreshPartyPreviews();
      refreshReadonly();
    });
  } else {
    (document.getElementById("sideload-msg") as HTMLElement).style.display = "block";
    (document.getElementById("app-body") as HTMLElement).style.display = "none";
  }
});
