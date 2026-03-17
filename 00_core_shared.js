/*
 * Bloc 0 - Coeur partage
 * Contient constantes globales, mappings, cache runtime et helpers transverses.
 */

/* global Charts, DriveApp, FormApp, HtmlService, Logger, PropertiesService, ScriptApp, Session, SpreadsheetApp, Utilities */

// Configuration globale du deployement.
const DEPLOYMENT_CONFIG = {
  rootFolderName: "GESTION_STOCK_DEPLOY",
  timezone: "Europe/Paris",
  locale: "fr_FR",
  sites: ["JLL"],
  adminEditors: ["seboulaugnac@gmail.com"],
  dashboardViewers: [],
};

const SYSTEM_CODE_VERSION = "2026.03.17.6";
const SYSTEM_COMPONENT_VERSIONS = {
  "00_core_shared.gs": "2026.03.17.6",
  "10_input_forms.gs": "2026.03.17.5",
  "20_raw_pipeline.gs": "2026.03.17.5",
  "30_domain_engine.gs": "2026.03.17.5",
  "40_dashboard_pilotage.gs": "2026.03.17.6",
  "NativeFormDialog.html": "2026.03.17.5",
  "NativeButtonsPanel.html": "2026.03.17.5",
  "appsscript.json": "2026.03.17.5",
};
const FORM_ID_PROPERTY_PREFIX = "FORM_ID__";
const DASHBOARD_ID_PROPERTY = "DASHBOARD_ID";
const USER_DASHBOARD_ID_PROPERTY = "USER_DASHBOARD_ID";
const CONTROLLER_DASHBOARD_ID_PROPERTY = "CONTROLLER_DASHBOARD_ID";
const ROLE_ADMIN = "ADMIN";
const ROLE_USER = "USER";
const ROLE_CONTROLLER = "CONTROLEUR";
const LIVE_FORM_MAX_CHOICES = 250;
const LIVE_FORM_REFRESH_HANDLER = "refreshAllFormsLiveChoices_";
const LAYOUT_VERSIONS = {
  GLOBAL_PILOTAGE: "8",
  USER_PILOTAGE: "8",
  CONTROLLER_PILOTAGE: "8",
  KPI_OVERVIEW: "3",
};
const NATIVE_FORM_TYPES = {
  FIRE_MOVEMENT: "fire_movement",
  FIRE_REPLENISHMENT: "fire_replenishment",
  FIRE_ITEM_CREATE: "fire_item_create",
  PHARMA_MOVEMENT: "pharma_movement",
  PHARMA_INVENTORY: "pharma_inventory",
  PHARMA_ITEM_CREATE: "pharma_item_create",
};
const FORM_RUNTIME_MODE = "NATIVE_HTML";
const DASHBOARD_USER_VISIBLE_TABS = [
  "SYSTEM_PILOTAGE",
  "STOCK_CONSOLIDE",
  "ALERTS_CONSOLIDE",
  "LOTS_CONSOLIDE",
  "PURCHASE_CONSOLIDE",
];
const FIRE_USER_VISIBLE_TABS = ["STOCK_VIEW", "ALERTS", "PURCHASE_REQUESTS"];
const PHARMA_USER_VISIBLE_TABS = ["STOCK_VIEW_PHARMACY", "ALERTS_PHARMACY", "LOTS_PHARMACY", "PURCHASE_REQUESTS_PHARMACY"];
const RUNTIME_CACHE = {
  accessProfiles: {},
  spreadsheetsById: {},
  spreadsheetsByUrl: {},
  formsById: {},
  formsByUrl: {},
  filesById: {},
};

const DASHBOARD_TABS = {
  CONFIG_SOURCES: [
    "SourceKey",
    "Module",
    "SiteKey",
    "SheetUrl",
    "StockRange",
    "AlertsRange",
    "LotsRange",
    "PurchaseRange",
    "Enabled",
  ],
  FORM_LINKS: ["FormKey", "Label", "FormUrl", "Module", "SiteKey", "IsActive"],
  NAVIGATION: ["Block", "Label", "TargetType", "Target"],
  KPI_OVERVIEW: ["KPIKey", "KPILabel", "Value", "Target", "Comment"],
  ACTIONS_RAPIDES: ["Label", "Link"],
  STOCK_CONSOLIDE: [
    "ItemID",
    "ItemName",
    "Category",
    "Unit",
    "Threshold",
    "CurrentQty",
    "GapToThreshold",
    "Status",
    "LastMovementAt",
    "SupplierID",
    "Module",
    "SiteKey",
  ],
  ALERTS_CONSOLIDE: [
    "AlertID",
    "Module",
    "SiteKey",
    "AlertType",
    "Severity",
    "ItemID",
    "ItemName",
    "Message",
    "CreatedAt",
    "Status",
    "Owner",
  ],
  LOTS_CONSOLIDE: [
    "LotID",
    "ItemID",
    "SiteKey",
    "LotNumber",
    "ExpiryDate",
    "QtyAvailable",
    "SupplierID",
    "ExpiryStatus",
  ],
  PURCHASE_CONSOLIDE: [
    "RequestID",
    "Module",
    "SiteKey",
    "ItemID",
    "ItemName",
    "SupplierID",
    "RequestedQty",
    "Priority",
    "Status",
    "RequestedBy",
    "RequestedAt",
    "ValidatedStatus",
  ],
  SYSTEM_PILOTAGE: ["Block", "Metric", "Value", "Link"],
  ACCESS_CONTROL: ["UserEmail", "Role", "AllowedModules", "AllowedSites", "CanUseDebug", "CanSeeAudit", "UpdatedAt"],
  REQUEST_VALIDATION: ["RequestID", "Decision", "ValidatedBy", "ValidatedAt", "Comment"],
  DEPLOYMENT_LOG: ["Type", "SiteKey", "Name", "URL", "ID", "CreatedAt"],
};

const USER_DASHBOARD_TABS = {
  FORM_LINKS: ["FormKey", "Label", "FormUrl", "Module", "SiteKey", "IsActive"],
  KPI_OVERVIEW: DASHBOARD_TABS.KPI_OVERVIEW.slice(),
  STOCK_CONSOLIDE: DASHBOARD_TABS.STOCK_CONSOLIDE.slice(),
  ALERTS_CONSOLIDE: DASHBOARD_TABS.ALERTS_CONSOLIDE.slice(),
  LOTS_CONSOLIDE: DASHBOARD_TABS.LOTS_CONSOLIDE.slice(),
  PURCHASE_CONSOLIDE: DASHBOARD_TABS.PURCHASE_CONSOLIDE.slice(),
  SYSTEM_PILOTAGE: ["Block", "Metric", "Value", "Link"],
};

const USER_DASHBOARD_VISIBLE_TABS = [
  "SYSTEM_PILOTAGE",
  "KPI_OVERVIEW",
  "STOCK_CONSOLIDE",
  "ALERTS_CONSOLIDE",
  "LOTS_CONSOLIDE",
  "PURCHASE_CONSOLIDE",
];

const CONTROLLER_DASHBOARD_TABS = {
  FORM_LINKS: ["FormKey", "Label", "FormUrl", "Module", "SiteKey", "IsActive"],
  KPI_OVERVIEW: DASHBOARD_TABS.KPI_OVERVIEW.slice(),
  STOCK_CONSOLIDE: DASHBOARD_TABS.STOCK_CONSOLIDE.slice(),
  ALERTS_CONSOLIDE: DASHBOARD_TABS.ALERTS_CONSOLIDE.slice(),
  LOTS_CONSOLIDE: DASHBOARD_TABS.LOTS_CONSOLIDE.slice(),
  PURCHASE_CONSOLIDE: DASHBOARD_TABS.PURCHASE_CONSOLIDE.slice(),
  SYSTEM_PILOTAGE: ["Block", "Metric", "Value", "Link"],
  INVENTORY_CONTROL: ["ItemID", "ItemName", "CurrentQty", "Threshold", "GapToThreshold", "Status", "Module", "SiteKey"],
};

const CONTROLLER_DASHBOARD_VISIBLE_TABS = [
  "SYSTEM_PILOTAGE",
  "KPI_OVERVIEW",
  "INVENTORY_CONTROL",
  "STOCK_CONSOLIDE",
  "ALERTS_CONSOLIDE",
  "LOTS_CONSOLIDE",
  "PURCHASE_CONSOLIDE",
];

const FIRE_TABS = {
  ITEMS: [
    "ItemID",
    "Module",
    "SiteKey",
    "ItemName",
    "Category",
    "Unit",
    "MinThreshold",
    "IsActive",
    "SupplierID",
    "StorageZone",
    "LastUpdatedAt",
  ],
  SUPPLIERS: [
    "SupplierID",
    "SupplierName",
    "ContactName",
    "Email",
    "Phone",
    "Address",
    "IsActive",
    "LastUpdatedAt",
  ],
  MOVEMENTS: [
    "MovementID",
    "Timestamp",
    "ItemID",
    "Module",
    "SiteKey",
    "MovementType",
    "QuantityDelta",
    "UnitCost",
    "Reason",
    "ActorEmail",
    "DocumentRef",
    "Comment",
  ],
  STOCK_VIEW: [
    "ItemID",
    "ItemName",
    "Category",
    "Unit",
    "Threshold",
    "CurrentQty",
    "GapToThreshold",
    "Status",
    "LastMovementAt",
    "SupplierID",
    "Module",
    "SiteKey",
  ],
  INVENTORY_COUNT: [
    "CountID",
    "CountDate",
    "ItemID",
    "Module",
    "SiteKey",
    "ExpectedQty",
    "CountedQty",
    "DiffQty",
    "CounterEmail",
    "ValidatedBy",
    "ValidationStatus",
    "Comment",
  ],
  PURCHASE_REQUESTS: [
    "RequestID",
    "Module",
    "SiteKey",
    "ItemID",
    "ItemName",
    "SupplierID",
    "RequestedQty",
    "Priority",
    "Status",
    "RequestedBy",
    "RequestedAt",
  ],
  ALERTS: [
    "AlertID",
    "Module",
    "SiteKey",
    "AlertType",
    "Severity",
    "ItemID",
    "ItemName",
    "Message",
    "CreatedAt",
    "Status",
    "Owner",
  ],
};

const PHARMA_TABS = {
  ITEMS_PHARMACY: [
    "ItemID",
    "Module",
    "SiteKey",
    "ItemName",
    "Category",
    "Unit",
    "MinThreshold",
    "IsActive",
    "SupplierID",
    "StorageZone",
    "LastUpdatedAt",
  ],
  SUPPLIERS_PHARMACY: [
    "SupplierID",
    "SupplierName",
    "ContactName",
    "Email",
    "Phone",
    "Address",
    "IsActive",
    "LastUpdatedAt",
  ],
  LOTS_PHARMACY: [
    "LotID",
    "ItemID",
    "SiteKey",
    "LotNumber",
    "ExpiryDate",
    "QtyAvailable",
    "SupplierID",
    "ExpiryStatus",
  ],
  MOVEMENTS_PHARMACY: [
    "MovementID",
    "Timestamp",
    "ItemID",
    "Module",
    "SiteKey",
    "MovementType",
    "QuantityDelta",
    "UnitCost",
    "Reason",
    "ActorEmail",
    "LotNumber",
    "ExpiryDate",
    "DocumentRef",
    "Comment",
  ],
  STOCK_VIEW_PHARMACY: [
    "ItemID",
    "ItemName",
    "Category",
    "Unit",
    "Threshold",
    "CurrentQty",
    "GapToThreshold",
    "Status",
    "LastMovementAt",
    "SupplierID",
    "Module",
    "SiteKey",
  ],
  INVENTORY_COUNT_PHARMACY: [
    "CountID",
    "CountDate",
    "ItemID",
    "Module",
    "SiteKey",
    "ExpectedQty",
    "CountedQty",
    "DiffQty",
    "CounterEmail",
    "ValidatedBy",
    "ValidationStatus",
    "Comment",
  ],
  PURCHASE_REQUESTS_PHARMACY: [
    "RequestID",
    "Module",
    "SiteKey",
    "ItemID",
    "ItemName",
    "SupplierID",
    "RequestedQty",
    "Priority",
    "Status",
    "RequestedBy",
    "RequestedAt",
  ],
  ALERTS_PHARMACY: [
    "AlertID",
    "Module",
    "SiteKey",
    "AlertType",
    "Severity",
    "ItemID",
    "ItemName",
    "Message",
    "CreatedAt",
    "Status",
    "Owner",
  ],
};

const KPI_ROWS = [
  ["TOTAL_ITEMS", "Total items active", "", "", "Count of distinct items"],
  ["LOW_STOCK", "Items under threshold", "", "", "Status SOUS_SEUIL + RUPTURE"],
  ["RUPTURE", "Items in rupture", "", "", "Status RUPTURE"],
  ["PURCHASE_OPEN", "Purchase requests open", "", "", "A_VALIDER + EN_COURS"],
  ["EXPIRY_90D", "Pharmacy lots expiring in 90 days", "", "", "ExpiryDate <= TODAY()+90"],
  ["EXPIRY_30D", "Pharmacy lots expiring in 30 days", "", "", "ExpiryDate <= TODAY()+30"],
  ["EXPIRY_PAST", "Pharmacy lots expired", "", "", "ExpiryDate < TODAY()"],
  ["CURRENT_STOCK_TOTAL", "Current stock total", "", "", "Sum of CurrentQty from STOCK_CONSOLIDE"],
  ["FIRE_STOCK_TOTAL", "Current stock incendie", "", "", "Sum CurrentQty for module incendie"],
  ["PHARMA_STOCK_TOTAL", "Current stock pharmacie", "", "", "Sum CurrentQty for module pharmacie"],
  ["STOCK_OK_RATE", "Stock OK rate (%)", "", "", "Share of items in status OK"],
];

function getCachedSpreadsheetById_(spreadsheetId) {
  const id = String(spreadsheetId || "").trim();
  if (!id) throw new Error("Spreadsheet ID vide.");
  if (RUNTIME_CACHE.spreadsheetsById[id]) return RUNTIME_CACHE.spreadsheetsById[id];
  const spreadsheet = SpreadsheetApp.openById(id);
  RUNTIME_CACHE.spreadsheetsById[id] = spreadsheet;
  return spreadsheet;
}

function getCachedSpreadsheetByUrl_(spreadsheetUrl) {
  const url = String(spreadsheetUrl || "").trim();
  if (!url) throw new Error("Spreadsheet URL vide.");
  if (RUNTIME_CACHE.spreadsheetsByUrl[url]) return RUNTIME_CACHE.spreadsheetsByUrl[url];

  const extractedId = extractIdFromDriveUrl_(url);
  if (extractedId && RUNTIME_CACHE.spreadsheetsById[extractedId]) {
    RUNTIME_CACHE.spreadsheetsByUrl[url] = RUNTIME_CACHE.spreadsheetsById[extractedId];
    return RUNTIME_CACHE.spreadsheetsByUrl[url];
  }

  const spreadsheet = SpreadsheetApp.openByUrl(url);
  const spreadsheetId = String(spreadsheet.getId() || "").trim();
  if (spreadsheetId) RUNTIME_CACHE.spreadsheetsById[spreadsheetId] = spreadsheet;
  RUNTIME_CACHE.spreadsheetsByUrl[url] = spreadsheet;
  return spreadsheet;
}

function getCachedFileById_(fileId) {
  const id = String(fileId || "").trim();
  if (!id) throw new Error("Drive file ID vide.");
  if (RUNTIME_CACHE.filesById[id]) return RUNTIME_CACHE.filesById[id];
  const file = DriveApp.getFileById(id);
  RUNTIME_CACHE.filesById[id] = file;
  return file;
}

function getCachedFormById_(formId) {
  const id = String(formId || "").trim();
  if (!id) throw new Error("Form ID vide.");
  if (RUNTIME_CACHE.formsById[id]) return RUNTIME_CACHE.formsById[id];
  const form = FormApp.openById(id);
  RUNTIME_CACHE.formsById[id] = form;
  return form;
}

function getCachedFormByUrl_(formUrl) {
  const url = String(formUrl || "").trim();
  if (!url) throw new Error("Form URL vide.");
  if (RUNTIME_CACHE.formsByUrl[url]) return RUNTIME_CACHE.formsByUrl[url];

  const extractedId = extractFormIdFromUrl_(url);
  if (extractedId && RUNTIME_CACHE.formsById[extractedId]) {
    RUNTIME_CACHE.formsByUrl[url] = RUNTIME_CACHE.formsById[extractedId];
    return RUNTIME_CACHE.formsByUrl[url];
  }

  const form = FormApp.openByUrl(url);
  const id = String(form.getId() || "").trim();
  if (id) RUNTIME_CACHE.formsById[id] = form;
  RUNTIME_CACHE.formsByUrl[url] = form;
  return form;
}

function openDashboardSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active && isDashboardSpreadsheet_(active)) {
    setStoredDashboardId_(active.getId());
    return active;
  }

  const id = getStoredDashboardId_();
  if (!id) return null;
  return getCachedSpreadsheetById_(id);
}

function resolveAdminDashboard_(actionLabel) {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active && active.getSheetByName("CONFIG_SOURCES")) {
    setStoredDashboardId_(active.getId());
    return active;
  }

  const stored = openDashboardSpreadsheet_();
  if (stored && stored.getSheetByName("CONFIG_SOURCES")) {
    return stored;
  }

  throw new Error(`${actionLabel || "Action"}: ouvrir le DASHBOARD_GLOBAL admin (CONFIG_SOURCES requis).`);
}

function normalizeEmail_(value) {
  return String(value || "").trim().toLowerCase();
}

function parseCsvList_(value) {
  const raw = String(value || "").trim();
  if (!raw || raw === "*") return ["*"];
  return raw
    .split(",")
    .map((item) => String(item || "").trim().toLowerCase())
    .filter((item) => item !== "");
}

function isTruthy_(value) {
  if (value === true) return true;
  const normalized = String(value || "").trim().toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "oui";
}

function normalizeTextKey_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function findHeaderIndex_(normalizedHeaders, tokens) {
  for (let i = 0; i < normalizedHeaders.length; i += 1) {
    const cell = normalizedHeaders[i];
    for (let t = 0; t < tokens.length; t += 1) {
      if (cell.indexOf(tokens[t]) !== -1) return i;
    }
  }
  return -1;
}

function parseNumberOrDefault_(value, defaultValue) {
  const normalized = String(value || "").trim().replace(",", ".");
  if (!normalized) return defaultValue;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : defaultValue;
}

function collectSheetIds_(spreadsheet) {
  const ids = {};
  spreadsheet.getSheets().forEach((sheet) => {
    ids[sheet.getSheetId()] = true;
  });
  return ids;
}

function appendSheetData_(fromSheet, toSheet) {
  const fromLastRow = fromSheet.getLastRow();
  if (fromLastRow < 2) return;

  const fromLastCol = fromSheet.getLastColumn();
  const toLastCol = toSheet.getLastColumn();
  const copyCols = Math.min(fromLastCol, toLastCol);
  if (copyCols < 1) return;

  const values = fromSheet.getRange(2, 1, fromLastRow - 1, copyCols).getValues()
    .filter((row) => row.some((cell) => String(cell || "").trim() !== ""));
  if (!values.length) return;

  const startRow = Math.max(2, toSheet.getLastRow() + 1);
  toSheet.getRange(startRow, 1, values.length, copyCols).setValues(values);
}

function getSystemComponentVersions_() {
  return Object.assign({}, SYSTEM_COMPONENT_VERSIONS);
}

function buildSystemComponentSignature_(versionsMap) {
  const map = versionsMap || getSystemComponentVersions_();
  const keys = Object.keys(map).sort();
  return keys.map((key) => `${key}@${String(map[key] || "").trim()}`).join("|");
}

function getSystemComponentSignature_() {
  return buildSystemComponentSignature_(getSystemComponentVersions_());
}

function parseSystemComponentSignature_(signature) {
  const map = {};
  const raw = String(signature || "").trim();
  if (!raw) return map;
  raw.split("|").forEach((part) => {
    const token = String(part || "").trim();
    if (!token) return;
    const at = token.lastIndexOf("@");
    if (at <= 0) return;
    const key = token.substring(0, at).trim();
    const version = token.substring(at + 1).trim();
    if (!key) return;
    map[key] = version;
  });
  return map;
}

function getSystemComponentDelta_(appliedSignature, currentSignature) {
  const applied = parseSystemComponentSignature_(appliedSignature);
  const current = parseSystemComponentSignature_(currentSignature || getSystemComponentSignature_());
  const keys = {};
  Object.keys(applied).forEach((key) => { keys[key] = true; });
  Object.keys(current).forEach((key) => { keys[key] = true; });

  const delta = [];
  Object.keys(keys).sort().forEach((key) => {
    const from = String(applied[key] || "");
    const to = String(current[key] || "");
    if (from === to) return;
    delta.push({ component: key, from, to });
  });
  return delta;
}

function formatSystemComponentDeltaSummary_(appliedSignature, currentSignature, maxItems) {
  const delta = getSystemComponentDelta_(appliedSignature, currentSignature);
  if (!delta.length) return "Aucun changement composant.";
  const limit = Math.max(1, Number(maxItems || 3));
  const head = delta.slice(0, limit).map((entry) => `${entry.component}: ${entry.from || "none"} -> ${entry.to || "none"}`);
  const suffix = delta.length > limit ? ` (+${delta.length - limit})` : "";
  return `${head.join(" | ")}${suffix}`;
}

function normalizeModuleKey_(value) {
  return normalizeTextKey_(value);
}

function normalizeSiteKey_(value) {
  return normalizeTextKey_(value);
}

function normalizeAccessList_(values, normalizer) {
  const normalize = typeof normalizer === "function" ? normalizer : ((entry) => String(entry || "").trim().toLowerCase());
  const input = Array.isArray(values) ? values : parseCsvList_(values);
  const dedup = {};
  (input || []).forEach((entry) => {
    const raw = String(entry || "").trim();
    if (!raw) return;
    if (raw === "*") {
      dedup["*"] = true;
      return;
    }
    const normalized = normalize(raw);
    if (normalized) dedup[normalized] = true;
  });
  return Object.keys(dedup);
}

function hasAccessWildcard_(values, normalizer) {
  const list = normalizeAccessList_(values, normalizer);
  return list.indexOf("*") >= 0;
}

function hasModuleAccess_(access, module) {
  const moduleKey = normalizeModuleKey_(module);
  if (!moduleKey) return false;
  if (access && access.isAdmin) return true;
  const allowedModules = normalizeAccessList_(access ? access.allowedModules : [], normalizeModuleKey_);
  if (!allowedModules.length) return false;
  if (allowedModules.indexOf("*") >= 0) return true;
  return allowedModules.indexOf(moduleKey) >= 0;
}

function hasSiteAccess_(access, siteKey) {
  const site = String(siteKey || "").trim();
  const siteKeyNormalized = normalizeSiteKey_(site);
  if (!siteKeyNormalized) return false;
  if (access && access.isAdmin) return true;
  const allowedSites = normalizeAccessList_(access ? access.allowedSites : [], normalizeSiteKey_);
  if (!allowedSites.length) return false;
  if (allowedSites.indexOf("*") >= 0) return true;
  return allowedSites.indexOf(siteKeyNormalized) >= 0;
}

function assertModuleAccess_(access, module, actionLabel) {
  const moduleLabel = String(module || "").trim().toLowerCase() || "module";
  if (hasModuleAccess_(access, moduleLabel)) return true;
  throw new Error(`Acces refuse au module ${moduleLabel}.`);
}

function assertSiteAccess_(access, siteKey, actionLabel) {
  const siteLabel = String(siteKey || "").trim() || "site";
  if (hasSiteAccess_(access, siteLabel)) return true;
  throw new Error(`Acces refuse au site ${siteLabel}.`);
}

function canUseNativeDestructiveActions_(access) {
  return !!(access && (access.isAdmin || access.isController));
}

function assertNativeDestructiveActionAccess_(access) {
  if (canUseNativeDestructiveActions_(access)) return true;
  throw new Error("Action reservee aux profils ADMIN / CONTROLEUR.");
}

function hasNativeDeleteIntentPayload_(payload) {
  const safePayload = payload || {};
  const movementType = String(safePayload.movementType || "").trim().toUpperCase();
  if (movementType === "DELETE_ITEM") return true;

  const deleteRaw = safePayload.deleteItems;
  const values = Array.isArray(deleteRaw)
    ? deleteRaw
    : String(deleteRaw || "").split(/[,\n;]+/);
  return values.some((entry) => String(entry || "").trim() !== "");
}

function isLegacyFireInventoryFormKey_(formKey) {
  const key = String(formKey || "").trim().toUpperCase();
  return key.indexOf("FORM_FIRE_INVENTORY_") === 0;
}

function isFireReplenishmentFormType_(formType) {
  const normalizedType = String(formType || "").trim().toLowerCase().replace(/-/g, "_");
  return normalizedType === "fire_replenishment" || normalizedType === "fire_inventory";
}

function rawSheetNameForFireReplenishment_() {
  // Compatibilite legacy: FORM_FIRE_INVENTORY_* reste mappe vers cet onglet RAW.
  return "FIRE_FORM_INVENTORY_RAW";
}
