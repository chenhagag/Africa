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

const TEMPLATES_LIST_DISPLAY_NAME = "ContractTemplates";
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

  setText("readonlyGeneral", generalLines.length ? generalLines.join("\n") : "—");
  setText("readonlyParties", partiesLines.join("\n"));
  setText("readonlyCost", costLines.join("\n"));
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
      refreshReadonly();
    });
  }

  const partyAClearBtn = document.getElementById("partyAClearBtn") as HTMLButtonElement | null;
  if (partyAClearBtn) {
    partyAClearBtn.addEventListener("click", () => {
      uiState.partyA.companyName = "";
      uiState.partyA.address = "";
      uiState.partyA.hp = "";
      uiState.partyA.namePercent = "";
      uiState.partyA.summary = "";

      const sel = document.getElementById("companyASelect") as HTMLSelectElement | null;
      if (sel) sel.value = "";
      setInputValue("companyAAddressInput", "");
      setInputValue("companyAHpInput", "");
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
      uiState.partyB.summary = buildPartyBSummary();
      refreshPartyPreviews();
      refreshReadonly();
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
  const contractVersion = pickField(fields, "contractVersion"); // ✅ כפי שאמרת
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

  // אם קיימים שדות summary מוכנים בספרייה — נעדיף אותם
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

/* =========
   Word update
   ========= */
async function updateDocumentFields(fields: Record<string, string | undefined>) {
  await Word.run(async (context) => {
    const tags = Object.keys(fields);

    const collections = tags.map(tag => {
      const col = context.document.contentControls.getByTag(tag);
      col.load("items");
      return { tag, col };
    });

    await context.sync();

    for (const { tag, col } of collections) {
      const val = (fields[tag] || "").toString();
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

  if (!uiState.cost.summary) {
    refreshCostStateFromUI();
    uiState.cost.summary = buildCostSummary();
    refreshCostPreview();
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
      otherSides,
      costCompMethod: uiState.cost.compMethod,
      costContractScope: uiState.cost.contractScope,
      costCurrency: uiState.cost.currency,
      costIndexType: uiState.cost.indexType,
      costBaseIndexDate: uiState.cost.baseIndexDate,
      costIndexMode: uiState.cost.indexMode,
      costIndexPoints: uiState.cost.indexPoints,
      costPaymentTerms: uiState.cost.paymentTerms
    });

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

      PartyACompanyName: partyACompanyName || undefined,
      PartyACompanyAddress: partyACompanyAddress || undefined,
      PartyACompanyHp: partyACompanyHp || undefined,
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
      CostPaymentTerms: uiState.cost.paymentTerms || undefined
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

    console.log("******* VERS 8 ************");

    (document.getElementById("sideload-msg") as HTMLElement).style.display = "none";
    (document.getElementById("app-body") as HTMLElement).style.display = "block";

    const btnUpdate = document.getElementById("runUpdateDoc");
    if (btnUpdate) (btnUpdate as HTMLDivElement).onclick = runUpdateDoc;

    const btnSave = document.getElementById("runSaveSystem");
    if (btnSave) (btnSave as HTMLDivElement).onclick = runSaveSystem;

    wireDates();
    wirePartiesUI();
    wireCostUI();

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

    loadLookups().then(async () => {
      refreshPartyPreviews();
      refreshCostStateFromUI();
      refreshCostPreview();
      refreshReadonly();

      // ✅ Load ALL fields from the library item
      await loadFromLibraryIntoUI();
    });
  } else {
    (document.getElementById("sideload-msg") as HTMLElement).style.display = "block";
    (document.getElementById("app-body") as HTMLElement).style.display = "none";
  }
});
