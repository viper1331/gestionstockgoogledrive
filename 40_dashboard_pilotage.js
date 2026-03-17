/*
 * Bloc 4 - Dashboards / consolidation / pilotage
 * Deploiement, consolides, pilotage, securite, maintenance et debug.
 */

function deployTotalSystem() {
  const folder = getOrCreateFolder_(DEPLOYMENT_CONFIG.rootFolderName);
  const now = new Date();

  const dashboard = createWorkbook_(folder, "DASHBOARD_GLOBAL", DASHBOARD_TABS);
  const sourceRows = [];
  const formRows = [];
  const navRows = [];
  const logRows = [["DASHBOARD", "GLOBAL", dashboard.spreadsheet.getName(), dashboard.url, dashboard.id, now]];

  DEPLOYMENT_CONFIG.sites.forEach((siteKey) => {
    const fire = createFireModule_(folder, siteKey, now);
    const pharma = createPharmaModule_(folder, siteKey, now);
    const formLogType = isNativeFormsMode_() ? "NATIVE_FORM" : "FORM";

    logRows.push(["WORKBOOK", siteKey, fire.workbookName, fire.workbookUrl, fire.workbookId, now]);
    logRows.push(["WORKBOOK", siteKey, pharma.workbookName, pharma.workbookUrl, pharma.workbookId, now]);

    fire.forms.forEach((f) => {
      setStoredFormId_(f.key, f.id);
      logRows.push([formLogType, siteKey, f.label, f.url, f.id, now]);
      formRows.push([f.key, f.label, f.url, "incendie", siteKey, true]);
      navRows.push(["Operations", f.label, "URL", f.url]);
    });

    pharma.forms.forEach((f) => {
      setStoredFormId_(f.key, f.id);
      logRows.push([formLogType, siteKey, f.label, f.url, f.id, now]);
      formRows.push([f.key, f.label, f.url, "pharmacie", siteKey, true]);
      navRows.push(["Operations", f.label, "URL", f.url]);
    });

    navRows.push(["Operations", `Open Incendie ${siteKey}`, "URL", fire.workbookUrl]);
    navRows.push(["Operations", `Open Pharmacie ${siteKey}`, "URL", pharma.workbookUrl]);

    sourceRows.push([
      `INCENDIE_${siteKey}`,
      "incendie",
      siteKey,
      fire.workbookUrl,
      "STOCK_VIEW!A:L",
      "ALERTS!A:K",
      "",
      "PURCHASE_REQUESTS!A:K",
      true,
    ]);

    sourceRows.push([
      `PHARMACIE_${siteKey}`,
      "pharmacie",
      siteKey,
      pharma.workbookUrl,
      "STOCK_VIEW_PHARMACY!A:L",
      "ALERTS_PHARMACY!A:K",
      "LOTS_PHARMACY!A:H",
      "PURCHASE_REQUESTS_PHARMACY!A:K",
      true,
    ]);
  });

  navRows.push(["Control", "Low stock alerts", "SHEET", "ALERTS_CONSOLIDE"]);
  navRows.push(["Control", "Expiry alerts", "SHEET", "LOTS_CONSOLIDE"]);
  navRows.push(["Control", "Purchase queue", "SHEET", "PURCHASE_CONSOLIDE"]);
  navRows.push(["Reporting", "Global stock", "SHEET", "STOCK_CONSOLIDE"]);
  navRows.push(["Reporting", "KPI overview", "SHEET", "KPI_OVERVIEW"]);

  populateDashboard_(dashboard.spreadsheet, sourceRows, formRows, navRows, logRows);
  if (isNativeFormsMode_()) {
    ensureNativeFormLinksFromDashboard_(dashboard.spreadsheet);
  }
  setStoredDashboardId_(dashboard.id);
  ensureAccessControlSheet_(dashboard.spreadsheet);
  const splitSync = propagateGlobalUpdateToSplitDashboards_(dashboard.spreadsheet, now, folder);
  protectDashboard_(dashboard.spreadsheet);
  applyVisibilityMode_(dashboard.spreadsheet, "user");
  focusSystemPilotage_(dashboard.spreadsheet);
  setAppliedCodeVersion_(SYSTEM_CODE_VERSION);
  setAppliedComponentsSignature_(getSystemComponentSignature_());

  applyFileAccess_(dashboard.file, DEPLOYMENT_CONFIG.adminEditors, DEPLOYMENT_CONFIG.dashboardViewers);
  SpreadsheetApp.flush();

  Logger.log(`Dashboard URL: ${dashboard.url}`);
  return {
    dashboardUrl: dashboard.url,
    dashboardId: dashboard.id,
    userDashboardUrl: splitSync.userDashboard.url,
    userDashboardId: splitSync.userDashboard.id,
    controllerDashboardUrl: splitSync.controllerDashboard.url,
    controllerDashboardId: splitSync.controllerDashboard.id,
    sites: DEPLOYMENT_CONFIG.sites.slice(),
    createdAt: now.toISOString(),
  };
}

function onOpen() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const isAdminDashboard = isDashboardSpreadsheet_(spreadsheet);
  const isUserDashboard = isUserDashboardSpreadsheet_(spreadsheet);
  if (!spreadsheet || (!isAdminDashboard && !isUserDashboard)) return;

  clearRuntimeAccessProfileCache_();
  let access = null;

  if (isAdminDashboard) {
    setStoredDashboardId_(spreadsheet.getId());
    try {
      ensureAuditSheetsExist_(spreadsheet);
      ensureAccessControlSheet_(spreadsheet);
      if (isNativeFormsMode_()) {
        ensureNativeFormLinksFromDashboard_(spreadsheet);
      }
    } catch (error) {
      Logger.log(`onOpen setup warning: ${String(error.message || error)}`);
    }
    access = getCurrentUserAccessProfile_(spreadsheet);
    try {
      applyDashboardVisibilityForAccess_(spreadsheet, access);
    } catch (error) {
      Logger.log(`onOpen visibility warning: ${String(error.message || error)}`);
    }
  } else {
    try {
      access = getCurrentUserAccessProfile_(resolveNativeFormsDashboard_());
    } catch (error) {
      access = getCurrentUserAccessProfile_(null);
      Logger.log(`onOpen access fallback warning: ${String(error.message || error)}`);
    }
  }

  focusSystemPilotage_(spreadsheet);
  addOperationsMenu_(access);

  if (isAdminDashboard && access.canUseDebug) {
    addDebugMenuIfDashboard_();
    notifyPendingUpdateIfAny_();
  }
}

function onInstall() {
  onOpen();
}

function onEdit(e) {
  try {
    handlePilotageToggleEdit_(e);
  } catch (error) {
    Logger.log(`onEdit pilotage toggle warning: ${String(error.message || error)}`);
  }
  try {
    const authMode = e && e.authMode ? String(e.authMode) : "";
    if (authMode === String(ScriptApp.AuthMode.FULL)) {
      handlePilotageNativeActionEdit_(e);
    } else {
      handlePilotageNativeActionNoAuth_(e);
    }
  } catch (error) {
    Logger.log(`onEdit pilotage native action warning: ${String(error.message || error)}`);
  }
}

function isDashboardSpreadsheet_(spreadsheet) {
  if (!spreadsheet) return false;
  return spreadsheet.getName().indexOf("DASHBOARD_GLOBAL") === 0
    || (!!spreadsheet.getSheetByName("CONFIG_SOURCES") && !!spreadsheet.getSheetByName("FORM_LINKS"));
}

function isUserDashboardSpreadsheet_(spreadsheet) {
  if (!spreadsheet) return false;
  if (spreadsheet.getSheetByName("CONFIG_SOURCES")) return false;
  return !!spreadsheet.getSheetByName("FORM_LINKS")
    && !!spreadsheet.getSheetByName("SYSTEM_PILOTAGE")
    && !!spreadsheet.getSheetByName("STOCK_CONSOLIDE");
}

function addOperationsMenu_(access) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Operations");
  const formsRefreshLabel = isNativeFormsMode_()
    ? "Synchroniser les formulaires natifs"
    : "Actualiser les choix formulaires";

  const refreshSubMenu = ui.createMenu("Actualisation")
    .addItem("Actualiser le pilotage", "refreshSystemPilotage")
    .addItem("Actualiser les consolides", "refreshDashboardConsolidations")
    .addItem(formsRefreshLabel, "refreshAllFormsLiveChoices_")
    .addItem("Appliquer l'affichage des tableaux", "applyCurrentPilotageTableVisibility_");
  menu.addSubMenu(refreshSubMenu);
  menu.addSubMenu(buildNativeFormsMenu_(ui, access));

  if (access && (access.isAdmin || access.isController)) {
    const requestSubMenu = ui.createMenu("Requetes achats")
      .addItem("Valider les requetes selectionnees", "operationsValiderRequetesSelectionnees_")
      .addItem("Mettre en cours les requetes selectionnees", "operationsMettreEnCoursRequetesSelectionnees_")
      .addItem("Refuser les requetes selectionnees", "operationsRefuserRequetesSelectionnees_");
    menu.addSubMenu(requestSubMenu);
  }

  if (access && access.canUseDebug) {
    menu.addSubMenu(
      ui.createMenu("Acces")
        .addItem("Appliquer la vue selon le role", "applyRoleBasedViewNow_")
    );
  }

  menu.addToUi();
}

function addDebugMenuIfDashboard_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet || !isDashboardSpreadsheet_(spreadsheet)) return;

  const access = getCurrentUserAccessProfile_(spreadsheet);
  if (!access.canUseDebug) return;

  const ui = SpreadsheetApp.getUi();
  const deploymentSubMenu = ui.createMenu("Deploiement et synchronisation")
    .addItem("MAJ en place du systeme", "debugApplyInPlaceUpdate_")
    .addItem("Propager la MAJ aux dashboards split", "debugPushGlobalUpdate_")
    .addItem("Synchroniser dashboard user", "debugSyncUserDashboard_")
    .addItem("Synchroniser dashboard controleur", "debugSyncControllerDashboard_");
  if (access.isAdmin) {
    deploymentSubMenu.addItem("Remise en service complete (Admin)", "debugAdminFullRecovery_");
  }

  ui.createMenu("Debug")
    .addSubMenu(deploymentSubMenu)
    .addSubMenu(
      ui.createMenu("Formulaires et donnees")
        .addItem("Reparer le flux complet", "debugRepairDataFlow_")
        .addItem("Reparer les connexions formulaires", "debugRepairFormConnections_")
        .addItem("Rafraichir UX formulaires", "debugRefreshFormsUx_")
        .addItem("Synchroniser les articles depuis RAW", "debugSyncItemsFromRaw_")
        .addItem("Reappliquer les formules modules", "debugReapplyModuleFormulas_")
        .addItem("Rafraichir les consolides", "debugRefreshConsolidations_")
        .addItem("Rafraichir le pilotage", "debugRefreshSystemPilotage_")
    )
    .addSubMenu(
      ui.createMenu("Triggers et autorisations")
        .addItem("Installer triggers submit formulaire", "debugInstallFormSubmitTriggers_")
        .addItem("Installer trigger refresh live", "debugInstallLiveRefreshTrigger_")
        .addItem("Construire helper IMPORTRANGE", "debugBuildImportRangeHelper_")
    )
    .addSubMenu(
      ui.createMenu("Acces et audits")
        .addItem("Appliquer mode utilisateur", "debugApplyUserMode_")
        .addItem("Appliquer mode admin", "debugApplyAdminMode_")
        .addItem("Ajouter un compte d'acces", "debugAddAccessAccount_")
        .addItem("Audit sante systeme", "debugRunSystemHealthCheck_")
        .addItem("Test flux suppression", "debugRunDeletionFlowTest_")
        .addItem("Audit OPS incendie/pharmacie", "debugAuditOpsModules_")
    )
    .addToUi();
}

function runDebugAction_(label, callback) {
  try {
    const result = callback();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (spreadsheet) spreadsheet.toast(`${label} OK`, "Action", 5);
    Logger.log(`${label} OK: ${JSON.stringify(result)}`);
  } catch (error) {
    const message = String(error && error.message ? error.message : error);
    SpreadsheetApp.getUi().alert(`Erreur (${label}): ${message}`);
    throw error;
  }
}

function debugApplyInPlaceUpdate_() {
  runDebugAction_("MAJ en place du systeme", applyInPlaceSystemUpdate);
}

function debugPushGlobalUpdate_() {
  runDebugAction_("Propagation MAJ dashboards split", pushGlobalUpdateFromDashboard);
}

function debugAdminFullRecovery_() {
  runDebugAction_("Remise en service complete (admin)", runAdminFullRecoveryAfterUpdate_);
}

function debugSyncUserDashboard_() {
  runDebugAction_("Synchronisation dashboard user", syncSplitUserDashboardFromActiveAdmin);
}

function debugSyncControllerDashboard_() {
  runDebugAction_("Synchronisation dashboard controleur", syncSplitControllerDashboardFromActiveAdmin);
}

function debugApplyUserMode_() {
  runDebugAction_("Application mode utilisateur", applyUserVisibilityModeFromDashboard);
}

function debugApplyAdminMode_() {
  runDebugAction_("Application mode admin", applyAdminVisibilityModeFromDashboard);
}

function debugRepairDataFlow_() {
  runDebugAction_("Reparation flux complet", repairDataFlowFromDashboard);
}

function debugRunSystemHealthCheck_() {
  runDebugAction_("Audit sante systeme", runSystemHealthCheckFromDashboard);
}

function debugRunDeletionFlowTest_() {
  runDebugAction_("Test flux suppression", runDeletionFlowTestFromDashboard);
}

function debugAddAccessAccount_() {
  runDebugAction_("Ajout compte acces", addAccessAccountFromPrompt_);
}

function debugRepairFormConnections_() {
  runDebugAction_("Reparation connexions formulaires", repairFormConnectionsFromDashboard);
}

function debugRefreshFormsUx_() {
  runDebugAction_("Rafraichissement UX formulaires", refreshFormsUxFromDashboard);
}

function debugSyncItemsFromRaw_() {
  runDebugAction_("Synchronisation articles depuis RAW", syncAllModuleItemsFromDashboard);
}

function debugReapplyModuleFormulas_() {
  runDebugAction_("Reapplication formules modules", reapplyModuleFormulasFromDashboard);
}

function debugInstallFormSubmitTriggers_() {
  runDebugAction_("Installation triggers submit", installModuleFormSubmitTriggersFromDashboard);
}

function debugInstallLiveRefreshTrigger_() {
  runDebugAction_("Installation trigger refresh live", installLiveFormRefreshTriggerFromDashboard);
}

function debugBuildImportRangeHelper_() {
  runDebugAction_("Construction helper IMPORTRANGE", buildImportRangeAuthorizationHelper);
}

function debugRefreshConsolidations_() {
  runDebugAction_("Rafraichissement consolides", refreshDashboardConsolidations);
}

function debugRefreshSystemPilotage_() {
  runDebugAction_("Rafraichissement system pilotage", refreshSystemPilotage);
}

function debugAuditOpsModules_() {
  runDebugAction_("Audit OPS modules", runOpsModulesAuditFromDashboard);
}

function operationsValiderRequetesSelectionnees_() {
  runDebugAction_("Validation requetes achat", () => applyRequestDecisionFromSelection_("VALIDE"));
}

function operationsMettreEnCoursRequetesSelectionnees_() {
  runDebugAction_("Requetes achat en cours", () => applyRequestDecisionFromSelection_("EN_COURS"));
}

function operationsRefuserRequetesSelectionnees_() {
  runDebugAction_("Requetes achat refusees", () => applyRequestDecisionFromSelection_("REFUSE"));
}

function applyRoleBasedViewNow_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet || !isDashboardSpreadsheet_(spreadsheet)) return;
  const access = getCurrentUserAccessProfile_(spreadsheet);
  applyDashboardVisibilityForAccess_(spreadsheet, access);
  focusSystemPilotage_(spreadsheet);
  SpreadsheetApp.flush();
}

function notifyPendingUpdateIfAny_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) return;

  const appliedVersion = getAppliedCodeVersion_();
  const appliedComponents = getAppliedComponentsSignature_();
  const currentComponents = getSystemComponentSignature_();
  const hasVersionDiff = appliedVersion !== SYSTEM_CODE_VERSION;
  const hasComponentDiff = appliedComponents !== currentComponents;
  if (hasVersionDiff || hasComponentDiff) {
    const componentSummary = formatSystemComponentDeltaSummary_(appliedComponents, currentComponents, 2);
    spreadsheet.toast(
      `Mise a jour disponible: ${appliedVersion || "aucune"} -> ${SYSTEM_CODE_VERSION}. ${componentSummary} Utiliser Debug > Deploiement et synchronisation > MAJ en place du systeme.`,
      "Debug",
      10
    );
  }
}

function ensureRequestValidationSheet_(dashboard) {
  if (!dashboard) return null;
  let sheet = dashboard.getSheetByName("REQUEST_VALIDATION");
  if (!sheet) {
    sheet = dashboard.insertSheet("REQUEST_VALIDATION");
    initSheet_(sheet, DASHBOARD_TABS.REQUEST_VALIDATION);
    return sheet;
  }

  const expected = DASHBOARD_TABS.REQUEST_VALIDATION;
  const width = expected.length;
  const current = sheet.getRange(1, 1, 1, width).getValues()[0].map((value) => String(value || "").trim());
  if (current.join("|") !== expected.join("|")) {
    sheet.clearContents();
    initSheet_(sheet, expected);
  } else {
    sheet.getRange(1, 1, 1, width).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function applyPurchaseValidationFormula_(dashboard) {
  if (!dashboard) return;
  ensureRequestValidationSheet_(dashboard);
  const sheet = dashboard.getSheetByName("PURCHASE_CONSOLIDE");
  if (!sheet) return;
  sheet.getRange("I1").setValue("Status").setFontWeight("bold");
  sheet.getRange("L1").setValue("ValidatedStatus").setFontWeight("bold");
  sheet.getRange("M1").setValue("StatusSource").setFontWeight("bold");
  setFormulaOnSheetIfChanged_(sheet, "I2", '=ARRAYFORMULA(IF(A2:A="";"";IFERROR(VLOOKUP(A2:A;REQUEST_VALIDATION!A:B;2;FALSE);M2:M)))');
  setFormulaOnSheetIfChanged_(sheet, "L2", '=ARRAYFORMULA(IF(A2:A="";"";I2:I))');
  try {
    sheet.hideColumns(13);
  } catch (error) {
    Logger.log(`Hide StatusSource column warning: ${String(error.message || error)}`);
  }
}

function getAppliedCodeVersion_() {
  return String(PropertiesService.getDocumentProperties().getProperty("APPLIED_CODE_VERSION") || "");
}

function setAppliedCodeVersion_(version) {
  PropertiesService.getDocumentProperties().setProperty("APPLIED_CODE_VERSION", String(version || ""));
}

function getAppliedComponentsSignature_() {
  return String(PropertiesService.getDocumentProperties().getProperty("APPLIED_COMPONENTS_SIGNATURE") || "");
}

function setAppliedComponentsSignature_(signature) {
  PropertiesService.getDocumentProperties().setProperty("APPLIED_COMPONENTS_SIGNATURE", String(signature || ""));
}

function setStoredDashboardId_(dashboardId) {
  const id = String(dashboardId || "").trim();
  if (!id) return;
  PropertiesService.getDocumentProperties().setProperty(DASHBOARD_ID_PROPERTY, id);
}

function getStoredDashboardId_() {
  return String(PropertiesService.getDocumentProperties().getProperty(DASHBOARD_ID_PROPERTY) || "").trim();
}

function setStoredUserDashboardId_(dashboardId) {
  const id = String(dashboardId || "").trim();
  if (!id) return;
  PropertiesService.getDocumentProperties().setProperty(USER_DASHBOARD_ID_PROPERTY, id);
}

function getStoredUserDashboardId_() {
  return String(PropertiesService.getDocumentProperties().getProperty(USER_DASHBOARD_ID_PROPERTY) || "").trim();
}

function setStoredControllerDashboardId_(dashboardId) {
  const id = String(dashboardId || "").trim();
  if (!id) return;
  PropertiesService.getDocumentProperties().setProperty(CONTROLLER_DASHBOARD_ID_PROPERTY, id);
}

function getStoredControllerDashboardId_() {
  return String(PropertiesService.getDocumentProperties().getProperty(CONTROLLER_DASHBOARD_ID_PROPERTY) || "").trim();
}

function layoutMarkerPropertyKey_(spreadsheet, markerName) {
  const spreadsheetId = spreadsheet ? String(spreadsheet.getId() || "").trim() : "";
  const marker = String(markerName || "").trim().toUpperCase();
  if (!spreadsheetId || !marker) return "";
  return `LAYOUT__${spreadsheetId}__${marker}`;
}

function getLayoutMarkerVersion_(spreadsheet, markerName) {
  const key = layoutMarkerPropertyKey_(spreadsheet, markerName);
  if (!key) return "";
  return String(PropertiesService.getScriptProperties().getProperty(key) || "").trim();
}

function setLayoutMarkerVersion_(spreadsheet, markerName, version) {
  const key = layoutMarkerPropertyKey_(spreadsheet, markerName);
  const value = String(version || "").trim();
  if (!key || !value) return;
  PropertiesService.getScriptProperties().setProperty(key, value);
}

function isLayoutMarkerCurrent_(spreadsheet, markerName, expectedVersion) {
  const actual = getLayoutMarkerVersion_(spreadsheet, markerName);
  return actual !== "" && actual === String(expectedVersion || "").trim();
}

function normalizeFormulaValue_(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function setFormulaIfChanged_(range, formula) {
  if (!range) return false;
  const target = normalizeFormulaValue_(formula);
  const current = normalizeFormulaValue_(range.getFormula());
  if (target === current) return false;
  range.setFormula(formula);
  return true;
}

function setFormulaOnSheetIfChanged_(sheet, a1, formula) {
  if (!sheet) return false;
  return setFormulaIfChanged_(sheet.getRange(a1), formula);
}

function isKpiOverviewLayoutReady_(sheet) {
  if (!sheet) return false;
  const header = String(sheet.getRange("G1").getDisplayValue() || "").trim();
  const moduleFormula = normalizeFormulaValue_(sheet.getRange("G4").getFormula());
  const statusFormula = normalizeFormulaValue_(sheet.getRange("K4").getFormula());
  const siteFormula = normalizeFormulaValue_(sheet.getRange("G23").getFormula());
  const purchaseStatusFormula = normalizeFormulaValue_(sheet.getRange("K23").getFormula());
  return header.indexOf("Analyse stock") === 0
    && moduleFormula.indexOf("QUERY(STOCK_CONSOLIDE") !== -1
    && statusFormula.indexOf("QUERY(STOCK_CONSOLIDE") !== -1
    && siteFormula.indexOf("QUERY(STOCK_CONSOLIDE") !== -1
    && purchaseStatusFormula.indexOf("QUERY(PURCHASE_CONSOLIDE") !== -1;
}

function isSystemPilotageLayoutReady_(sheet, mode) {
  if (!sheet) return false;
  const leftHeader = String(sheet.getRange("A1").getDisplayValue() || "").trim();
  const rightHeader = String(sheet.getRange("E1").getDisplayValue() || "").trim();
  if (leftHeader !== "CONTROL CENTER") return false;
  if (String(mode || "").toUpperCase() === "GLOBAL") {
    return rightHeader.indexOf("SYSTEM PILOTAGE | Gestion de stock") === 0;
  }
  if (String(mode || "").toUpperCase() === "CONTROLLER") {
    return rightHeader.indexOf("SYSTEM PILOTAGE | Controle inventaire") === 0;
  }
  return rightHeader.indexOf("SYSTEM PILOTAGE | Exploitation quotidienne") === 0;
}

function clearRuntimeAccessProfileCache_() {
  RUNTIME_CACHE.accessProfiles = {};
}

function buildAdminEmailSet_() {
  const set = {};
  (DEPLOYMENT_CONFIG.adminEditors || []).forEach((email) => {
    const normalized = normalizeEmail_(email);
    if (normalized) set[normalized] = true;
  });
  return set;
}

function getCurrentUserEmail_() {
  const activeUserEmail = normalizeEmail_(Session.getActiveUser().getEmail());
  if (activeUserEmail) return activeUserEmail;
  return normalizeEmail_(Session.getEffectiveUser().getEmail());
}

function parseAccessControlListStrict_(value, normalizer) {
  const normalize = typeof normalizer === "function" ? normalizer : ((entry) => String(entry || "").trim().toLowerCase());
  const raw = String(value || "").trim();
  if (!raw) return [];
  if (raw === "*") return ["*"];

  const dedup = {};
  raw.split(",").forEach((token) => {
    const cleaned = String(token || "").trim();
    if (!cleaned) return;
    if (cleaned === "*") {
      dedup["*"] = true;
      return;
    }
    const normalized = normalize(cleaned);
    if (normalized) dedup[normalized] = true;
  });
  return Object.keys(dedup);
}

function buildDefaultAccessProfile_(email, adminSet) {
  const normalizedEmail = normalizeEmail_(email);
  const isAdmin = !!adminSet[normalizedEmail];
  return {
    email: normalizedEmail,
    role: isAdmin ? ROLE_ADMIN : ROLE_USER,
    allowedModules: ["*"],
    allowedSites: ["*"],
    canUseDebug: isAdmin,
    canSeeAudit: isAdmin,
    isAdmin,
    isController: false,
  };
}

function getAccessProfileForEmail_(dashboard, email) {
  const normalizedEmail = normalizeEmail_(email);
  const adminSet = buildAdminEmailSet_();
  const dashboardId = dashboard ? String(dashboard.getId() || "") : "NO_DASHBOARD";
  const cacheKey = `${dashboardId}::${normalizedEmail}`;
  if (RUNTIME_CACHE.accessProfiles[cacheKey]) {
    return RUNTIME_CACHE.accessProfiles[cacheKey];
  }
  const defaultProfile = buildDefaultAccessProfile_(normalizedEmail, adminSet);

  if (!dashboard) {
    RUNTIME_CACHE.accessProfiles[cacheKey] = defaultProfile;
    return defaultProfile;
  }

  const accessSheet = dashboard.getSheetByName("ACCESS_CONTROL");
  if (!accessSheet || accessSheet.getLastRow() < 2 || !normalizedEmail) {
    RUNTIME_CACHE.accessProfiles[cacheKey] = defaultProfile;
    return defaultProfile;
  }
  const rows = accessSheet.getRange(2, 1, accessSheet.getLastRow() - 1, 7).getValues();

  for (let i = 0; i < rows.length; i += 1) {
    const rowEmail = normalizeEmail_(rows[i][0]);
    if (!rowEmail || rowEmail !== normalizedEmail) continue;

    const roleValue = String(rows[i][1] || "").trim().toUpperCase() || ROLE_USER;
    const isAdmin = roleValue === ROLE_ADMIN || !!adminSet[normalizedEmail];
    const isController = !isAdmin && roleValue === ROLE_CONTROLLER;
    const profile = {
      email: normalizedEmail,
      role: isAdmin ? ROLE_ADMIN : (isController ? ROLE_CONTROLLER : ROLE_USER),
      allowedModules: isAdmin ? ["*"] : parseAccessControlListStrict_(rows[i][2], normalizeModuleKey_),
      allowedSites: isAdmin ? ["*"] : parseAccessControlListStrict_(rows[i][3], normalizeSiteKey_),
      canUseDebug: isAdmin ? true : isTruthy_(rows[i][4]),
      canSeeAudit: isAdmin ? true : isTruthy_(rows[i][5]),
      isAdmin,
      isController,
    };
    RUNTIME_CACHE.accessProfiles[cacheKey] = profile;
    return profile;
  }

  RUNTIME_CACHE.accessProfiles[cacheKey] = defaultProfile;
  return defaultProfile;
}

function getCurrentUserAccessProfile_(dashboard) {
  const email = getCurrentUserEmail_();
  return getAccessProfileForEmail_(dashboard, email);
}

function assertAdminOrController_(dashboard, actionLabel) {
  const access = getCurrentUserAccessProfile_(dashboard);
  if (access && (access.isAdmin || access.isController)) return access;
  throw new Error(`${actionLabel || "Action"}: reserve aux profils ADMIN ou CONTROLEUR.`);
}

function assertAdmin_(dashboard, actionLabel) {
  const access = getCurrentUserAccessProfile_(dashboard);
  if (access && access.isAdmin) return access;
  throw new Error(`${actionLabel || "Action"}: reserve au profil ADMIN.`);
}

function collectRequestIdsFromSelection_(spreadsheet) {
  if (!spreadsheet) return [];
  const sheet = spreadsheet.getActiveSheet();
  const range = spreadsheet.getActiveRange();
  if (!sheet || !range) return [];

  const ids = {};
  const pushId = (value) => {
    const token = String(value || "").trim();
    if (!/^REQ-/i.test(token)) return;
    ids[token.toUpperCase()] = token.toUpperCase();
  };

  const values = range.getValues();
  values.forEach((row) => row.forEach((cell) => pushId(cell)));

  const sheetName = String(sheet.getName() || "").toUpperCase();
  if (sheetName.indexOf("PURCHASE") !== -1) {
    for (let r = range.getRow(); r < range.getRow() + range.getNumRows(); r += 1) {
      pushId(sheet.getRange(r, 1).getDisplayValue());
    }
  }

  return Object.keys(ids).map((key) => ids[key]);
}

function parseRequestIdsInput_(value) {
  const raw = String(value || "").trim();
  if (!raw) return [];
  const ids = {};
  raw.split(/[,\n; ]+/).forEach((token) => {
    const current = String(token || "").trim().toUpperCase();
    if (!/^REQ-[A-Z0-9_-]+$/.test(current)) return;
    ids[current] = current;
  });
  return Object.keys(ids).map((key) => ids[key]);
}

function getEligibleRequestStatusesForDecision_(decision) {
  const normalized = String(decision || "").trim().toUpperCase();
  if (normalized === "EN_COURS") return ["A_VALIDER"];
  if (normalized === "VALIDE" || normalized === "REFUSE") return ["A_VALIDER", "EN_COURS"];
  return ["A_VALIDER"];
}

function collectRequestIdsFromPurchaseByStatus_(dashboard, statuses, maxCount) {
  const sheet = dashboard ? dashboard.getSheetByName("PURCHASE_CONSOLIDE") : null;
  if (!sheet || sheet.getLastRow() < 2) return [];

  const accepted = {};
  (statuses || []).forEach((status) => {
    const key = String(status || "").trim().toUpperCase();
    if (key) accepted[key] = true;
  });
  if (!Object.keys(accepted).length) accepted.A_VALIDER = true;

  const limit = Math.max(1, Number(maxCount || 500));
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  const ids = {};
  const ordered = [];

  rows.forEach((row) => {
    const requestId = String(row[0] || "").trim().toUpperCase();
    if (!requestId || !/^REQ-[A-Z0-9_-]+$/.test(requestId)) return;
    const validatedStatus = String(row[11] || "").trim().toUpperCase();
    const rawStatus = String(row[8] || "").trim().toUpperCase();
    const status = validatedStatus || rawStatus;
    if (!accepted[status]) return;
    if (ids[requestId]) return;
    ids[requestId] = true;
    ordered.push(requestId);
  });

  return ordered.slice(0, limit);
}

function resolveRequestIdsForDecision_(adminDashboard, activeSpreadsheet, decision, ui) {
  let requestIds = collectRequestIdsFromSelection_(activeSpreadsheet);
  if (requestIds.length) return requestIds;

  const eligibleStatuses = getEligibleRequestStatusesForDecision_(decision);
  const autoIds = collectRequestIdsFromPurchaseByStatus_(adminDashboard, eligibleStatuses, 500);
  if (!autoIds.length) {
    throw new Error(`Aucune requete eligible detectee (statuts: ${eligibleStatuses.join(", ")}).`);
  }

  const decisionLabel = String(decision || "").trim().toUpperCase();
  const preview = autoIds.slice(0, 20).join(", ");
  const suffix = autoIds.length > 20 ? ` ... (+${autoIds.length - 20})` : "";
  const response = ui.alert(
    "Selection des requetes",
    `Aucune requete detectee dans la selection.\n${autoIds.length} requete(s) chargee(s) automatiquement (${eligibleStatuses.join(" / ")}).\n\nAppliquer la decision ${decisionLabel} a toutes ces requetes ?\n\n${preview}${suffix}`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return [];

  return autoIds;
}

function upsertRequestValidationRows_(sheet, requestIds, decision, actorEmail, comment) {
  if (!sheet || !requestIds || !requestIds.length) return 0;
  const normalizedDecision = String(decision || "").trim().toUpperCase();
  if (!normalizedDecision) return 0;

  const now = new Date();
  const existing = {};
  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    rows.forEach((row, idx) => {
      const requestId = String(row[0] || "").trim().toUpperCase();
      if (requestId) existing[requestId] = idx + 2;
    });
  }

  const rowsToAppend = [];
  requestIds.forEach((requestId) => {
    const id = String(requestId || "").trim().toUpperCase();
    if (!id) return;
    const payload = [id, normalizedDecision, actorEmail || "", now, comment || ""];
    const targetRow = existing[id];
    if (targetRow) {
      sheet.getRange(targetRow, 1, 1, payload.length).setValues([payload]);
    } else {
      rowsToAppend.push(payload);
    }
  });

  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  return requestIds.length;
}

function buildSourceWorkbookUrlMap_(dashboard, moduleFilter) {
  const map = {};
  if (!dashboard) return map;
  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet || sourceSheet.getLastRow() < 2) return map;

  const normalizedModule = normalizeTextKey_(moduleFilter);
  const rows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  rows.forEach((row) => {
    const module = normalizeTextKey_(row[1]);
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];
    if (!siteKey || !workbookUrl || !isTruthy_(enabled)) return;
    if (normalizedModule && module !== normalizedModule) return;
    map[siteKey] = workbookUrl;
  });
  return map;
}

function collectFireReplenishmentRequestsByIds_(dashboard, requestIds) {
  if (!dashboard || !requestIds || !requestIds.length) return [];
  const purchaseSheet = dashboard.getSheetByName("PURCHASE_CONSOLIDE");
  if (!purchaseSheet || purchaseSheet.getLastRow() < 2) return [];

  const wanted = {};
  requestIds.forEach((id) => {
    const key = String(id || "").trim().toUpperCase();
    if (key) wanted[key] = true;
  });
  if (!Object.keys(wanted).length) return [];

  const rows = purchaseSheet.getRange(2, 1, purchaseSheet.getLastRow() - 1, 12).getValues();
  const requests = [];
  rows.forEach((row) => {
    const requestId = String(row[0] || "").trim().toUpperCase();
    const module = normalizeTextKey_(row[1]);
    const siteKey = String(row[2] || "").trim();
    const itemId = String(row[3] || "").trim();
    const itemName = String(row[4] || "").trim();
    const requestedQty = parseNumberOrDefault_(row[6], 0);
    if (!requestId || !wanted[requestId]) return;
    if (module !== "incendie") return;
    if (requestId.indexOf("REQ-INC-FRM-") !== 0) return;
    if (!siteKey || !itemId || requestedQty <= 0) return;
    requests.push({
      requestId,
      siteKey,
      itemId,
      itemName,
      requestedQty,
      requestedBy: String(row[9] || "").trim(),
      requestedAt: row[10],
    });
  });
  return requests;
}

function appendValidatedFireReplenishmentRows_(workbook, requests, validatorEmail, validationComment) {
  if (!workbook || !requests || !requests.length) {
    return { appended: 0, skippedAlreadyApplied: 0 };
  }
  const rawSheet = workbook.getSheetByName("FIRE_FORM_MOVEMENTS_RAW");
  if (!rawSheet) {
    throw new Error(`FIRE_FORM_MOVEMENTS_RAW introuvable dans ${workbook.getName()}.`);
  }

  const lastCol = rawSheet.getLastColumn();
  const headers = rawSheet.getRange(1, 1, 1, lastCol).getValues()[0].map((value) => normalizeTextKey_(value));
  const timestampIdx = findHeaderIndex_(headers, ["horodateur", "timestamp"]);
  const siteIdx = findHeaderIndex_(headers, ["site concerne", "sitekey", "site"]);
  const itemIdx = findItemFieldIndex_(headers);
  const typeIdx = findHeaderIndex_(headers, ["type de mouvement", "movementtype", "movement type"]);
  const qtyIdx = findHeaderIndex_(headers, ["quantite", "quantity"]);
  const reasonIdx = findHeaderIndex_(headers, ["motif", "reason"]);
  const actorIdx = findHeaderIndex_(headers, ["email operateur", "actoremail", "operator email"]);
  const docRefIdx = findHeaderIndex_(headers, ["reference document", "document ref", "documentref"]);
  const commentIdx = findHeaderIndex_(headers, ["commentaire", "comment"]);

  const existingDocRefs = {};
  if (docRefIdx >= 0 && rawSheet.getLastRow() >= 2) {
    const values = rawSheet.getRange(2, docRefIdx + 1, rawSheet.getLastRow() - 1, 1).getValues();
    values.forEach((row) => {
      const key = normalizeTextKey_(row[0]);
      if (key) existingDocRefs[key] = true;
    });
  }

  const now = new Date();
  const rowsToAppend = [];
  let skippedAlreadyApplied = 0;

  requests.forEach((request) => {
    const requestId = String(request.requestId || "").trim().toUpperCase();
    if (!requestId) return;
    const docKey = normalizeTextKey_(requestId);
    if (docKey && existingDocRefs[docKey]) {
      skippedAlreadyApplied += 1;
      return;
    }

    const row = new Array(lastCol).fill("");
    if (timestampIdx >= 0) row[timestampIdx] = now;
    if (siteIdx >= 0) row[siteIdx] = request.siteKey;
    if (itemIdx >= 0) row[itemIdx] = request.itemName ? `${request.itemId} | ${request.itemName}` : request.itemId;
    if (typeIdx >= 0) row[typeIdx] = "IN";
    if (qtyIdx >= 0) row[qtyIdx] = Math.abs(parseNumberOrDefault_(request.requestedQty, 0));
    if (reasonIdx >= 0) row[reasonIdx] = "Reapprovisionnement valide";
    if (actorIdx >= 0) row[actorIdx] = validatorEmail || "";
    if (docRefIdx >= 0) row[docRefIdx] = requestId;
    if (commentIdx >= 0) {
      row[commentIdx] = validationComment
        ? `Validation ${validatorEmail || ""} - ${validationComment}`
        : `Validation ${validatorEmail || ""}`.trim();
    }

    rowsToAppend.push(row);
    if (docKey) existingDocRefs[docKey] = true;
  });

  if (rowsToAppend.length) {
    rawSheet.getRange(rawSheet.getLastRow() + 1, 1, rowsToAppend.length, lastCol).setValues(rowsToAppend);
  }

  return { appended: rowsToAppend.length, skippedAlreadyApplied };
}

function applyValidatedFireReplenishmentsFromDashboard_(dashboard, requestIds, validatorEmail, decision, validationComment) {
  const normalizedDecision = String(decision || "").trim().toUpperCase();
  if (normalizedDecision !== "VALIDE") {
    return { eligibleRequests: 0, appendedMovements: 0, skippedAlreadyApplied: 0, updatedWorkbooks: 0 };
  }

  const requests = collectFireReplenishmentRequestsByIds_(dashboard, requestIds);
  if (!requests.length) {
    return { eligibleRequests: 0, appendedMovements: 0, skippedAlreadyApplied: 0, updatedWorkbooks: 0 };
  }

  const siteWorkbookMap = buildSourceWorkbookUrlMap_(dashboard, "incendie");
  const grouped = {};
  requests.forEach((request) => {
    if (!grouped[request.siteKey]) grouped[request.siteKey] = [];
    grouped[request.siteKey].push(request);
  });

  let appendedMovements = 0;
  let skippedAlreadyApplied = 0;
  let updatedWorkbooks = 0;

  Object.keys(grouped).forEach((siteKey) => {
    const workbookUrl = siteWorkbookMap[siteKey];
    if (!workbookUrl) return;
    try {
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      const result = appendValidatedFireReplenishmentRows_(
        workbook,
        grouped[siteKey],
        validatorEmail,
        validationComment
      );
      appendedMovements += result.appended;
      skippedAlreadyApplied += result.skippedAlreadyApplied;
      if (result.appended > 0) {
        syncFireItems_(workbook);
        applyFireFormulas_(workbook);
        protectSheets_(workbook, ["MOVEMENTS", "INVENTORY_COUNT", "STOCK_VIEW", "ALERTS", "PURCHASE_REQUESTS"]);
        updatedWorkbooks += 1;
      }
    } catch (error) {
      Logger.log(`Reappro validation warning (${siteKey}): ${String(error.message || error)}`);
    }
  });

  return {
    eligibleRequests: requests.length,
    appendedMovements,
    skippedAlreadyApplied,
    updatedWorkbooks,
  };
}

function applyRequestDecisionFromSelection_(decision) {
  const adminDashboard = resolveAdminDashboard_("Validation des requetes");
  const access = assertAdminOrController_(adminDashboard, "Validation des requetes");
  ensureRequestValidationSheet_(adminDashboard);

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const requestIds = resolveRequestIdsForDecision_(adminDashboard, activeSpreadsheet, decision, ui);
  if (!requestIds.length) {
    return { decision: String(decision || "").toUpperCase(), updatedRequests: 0, status: "CANCELLED" };
  }

  const comment = promptText_(
    ui,
    "Commentaire validation (optionnel)",
    `Decision ${String(decision || "").toUpperCase()} pour ${requestIds.length} requete(s).`
  );

  const validationSheet = adminDashboard.getSheetByName("REQUEST_VALIDATION");
  const changed = upsertRequestValidationRows_(validationSheet, requestIds, decision, access.email, comment);
  SpreadsheetApp.flush();
  const replenishment = applyValidatedFireReplenishmentsFromDashboard_(
    adminDashboard,
    requestIds,
    access.email,
    decision,
    comment
  );
  refreshDashboardConsolidations({ skipSplitSync: true });
  refreshSystemPilotage({ skipSplitSync: true });
  propagateGlobalUpdateToSplitDashboards_(adminDashboard, new Date());
  SpreadsheetApp.flush();

  return {
    decision: String(decision || "").toUpperCase(),
    updatedRequests: changed,
    replenishment,
  };
}

function applyDashboardVisibilityForAccess_(dashboard, access) {
  if (!dashboard) return;
  if (access && access.isAdmin) {
    showAllSheets_(dashboard);
  } else {
    const visible = DASHBOARD_USER_VISIBLE_TABS.slice();
    if (access && access.canSeeAudit) {
      visible.push("SYSTEM_HEALTH_AUDIT", "FORM_CONNECTION_AUDIT", "FORM_UX_AUDIT", "DELETE_FLOW_AUDIT", "OPS_MODULE_AUDIT", "REQUEST_VALIDATION");
    }
    setVisibleTabsOnly_(dashboard, visible);
  }
}

function focusSystemPilotage_(spreadsheet) {
  if (!spreadsheet) return;
  const pilotage = spreadsheet.getSheetByName("SYSTEM_PILOTAGE");
  if (!pilotage) return;
  spreadsheet.setActiveSheet(pilotage);
  spreadsheet.setCurrentCell(pilotage.getRange("A1"));
}

function addAccessAccountFromPrompt_() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard || !isDashboardSpreadsheet_(dashboard)) {
    throw new Error("Ouvrez le dashboard global avant d'ajouter un compte.");
  }
  setStoredDashboardId_(dashboard.getId());
  ensureAccessControlSheet_(dashboard);

  const access = getCurrentUserAccessProfile_(dashboard);
  if (!access.isAdmin) {
    throw new Error("Action reservee aux admins.");
  }

  const ui = SpreadsheetApp.getUi();
  const emailInput = promptText_(
    ui,
    "Ajout de compte",
    "Saisir l'email du compte a ajouter (ex: user@domaine.com)"
  );
  if (!emailInput) return { status: "CANCELLED" };

  const email = normalizeEmail_(emailInput);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error(`Email invalide: ${emailInput}`);
  }

  const roleInput = promptText_(
    ui,
    "Role du compte",
    "Saisir le role: ADMIN, CONTROLEUR ou USER"
  );
  if (!roleInput) return { status: "CANCELLED" };

  const role = String(roleInput || "").trim().toUpperCase();
  if (role !== ROLE_ADMIN && role !== ROLE_CONTROLLER && role !== ROLE_USER) {
    throw new Error("Role invalide. Valeurs autorisees: ADMIN, CONTROLEUR, USER.");
  }

  const defaults = buildRoleDefaults_(role);
  const result = upsertAccessControlAccount_(
    dashboard,
    email,
    role,
    defaults.allowedModules,
    defaults.allowedSites,
    defaults.canUseDebug,
    defaults.canSeeAudit
  );
  try {
    const splitSync = propagateGlobalUpdateToSplitDashboards_(dashboard, new Date());
    result.userDashboardUrl = splitSync.userDashboard.url;
    result.controllerDashboardUrl = splitSync.controllerDashboard.url;
  } catch (error) {
    Logger.log(`User dashboard sync after access update failed: ${String(error.message || error)}`);
  }

  ui.alert(`Compte ${result.action}: ${result.email} (${result.role})`);
  return result;
}

function promptText_(ui, title, message) {
  const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return "";
  return String(response.getResponseText() || "").trim();
}

function buildRoleDefaults_(role) {
  if (role === ROLE_ADMIN) {
    return {
      allowedModules: "*",
      allowedSites: "*",
      canUseDebug: true,
      canSeeAudit: true,
    };
  }
  if (role === ROLE_CONTROLLER) {
    return {
      allowedModules: "*",
      allowedSites: "*",
      canUseDebug: false,
      canSeeAudit: true,
    };
  }
  return {
    allowedModules: "*",
    allowedSites: "*",
    canUseDebug: false,
    canSeeAudit: false,
  };
}

function upsertAccessControlAccount_(dashboard, email, role, allowedModules, allowedSites, canUseDebug, canSeeAudit) {
  ensureAccessControlSheet_(dashboard);
  const sheet = dashboard.getSheetByName("ACCESS_CONTROL");
  if (!sheet) throw new Error("ACCESS_CONTROL introuvable.");

  const rowValues = [
    normalizeEmail_(email),
    role,
    allowedModules || "*",
    allowedSites || "*",
    !!canUseDebug,
    !!canSeeAudit,
    new Date(),
  ];

  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, 1, rowValues.length).setValues([rowValues]);
    clearRuntimeAccessProfileCache_();
    return { action: "CREATED", email: rowValues[0], role };
  }

  const emails = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < emails.length; i += 1) {
    const rowEmail = normalizeEmail_(emails[i][0]);
    if (rowEmail !== rowValues[0]) continue;
    const rowNumber = i + 2;
    sheet.getRange(rowNumber, 1, 1, rowValues.length).setValues([rowValues]);
    clearRuntimeAccessProfileCache_();
    return { action: "UPDATED", email: rowValues[0], role };
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, rowValues.length).setValues([rowValues]);
  clearRuntimeAccessProfileCache_();
  return { action: "CREATED", email: rowValues[0], role };
}

function ensureItemCreationFormsFromDashboard_() {
  const dashboard = resolveAdminDashboard_("Creation formulaires nouveaux articles");
  setStoredDashboardId_(dashboard.getId());
  if (isNativeFormsMode_()) {
    const sync = ensureNativeFormLinksFromDashboard_(dashboard);
    return {
      checkedSources: sync.sourcesChecked,
      createdForms: sync.createdRows,
      updatedRows: sync.updatedRows + sync.disabledRows,
    };
  }

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  const formSheet = dashboard.getSheetByName("FORM_LINKS");
  if (!sourceSheet || !formSheet || sourceSheet.getLastRow() < 2) {
    return { checkedSources: 0, createdForms: 0, updatedRows: 0 };
  }

  const sourceRows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  let formRows = formSheet.getLastRow() > 1
    ? formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 6).getValues()
    : [];
  const formMap = {};
  formRows.forEach((row, index) => {
    const key = String(row[0] || "").trim();
    if (!key) return;
    formMap[key] = { rowIndex: index + 2, row };
  });

  const folder = getOrCreateFolder_(DEPLOYMENT_CONFIG.rootFolderName);
  const logSheet = dashboard.getSheetByName("DEPLOYMENT_LOG");
  const rowsToAppend = [];
  let checkedSources = 0;
  let createdForms = 0;
  let updatedRows = 0;

  sourceRows.forEach((row) => {
    const sourceKey = String(row[0] || "").trim();
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];
    if (!sourceKey || !module || !siteKey || !workbookUrl || !isTruthy_(enabled)) return;

    let target = null;
    if (module === "incendie") {
      target = {
        key: `FORM_FIRE_ITEM_CREATE_${siteKey}`,
        label: `Form creation article incendie ${siteKey}`,
        raw: "FIRE_FORM_MOVEMENTS_RAW",
        builder: buildFireItemCreationForm_,
      };
    } else if (module === "pharmacie") {
      target = {
        key: `FORM_PHARMA_ITEM_CREATE_${siteKey}`,
        label: `Form creation article pharmacie ${siteKey}`,
        raw: "PHARMA_FORM_MOVEMENTS_RAW",
        builder: buildPharmaItemCreationForm_,
      };
    }
    if (!target) return;
    checkedSources += 1;

    const existing = formMap[target.key];
    if (!existing) {
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      const formBundle = target.builder(folder, siteKey, workbook);
      attachFormToSpreadsheet_(formBundle.form, workbook, target.raw, true, false);
      setStoredFormId_(target.key, formBundle.id);

      rowsToAppend.push([target.key, target.label, formBundle.url, module, siteKey, true]);
      formMap[target.key] = { rowIndex: -1, row: [target.key, target.label, formBundle.url, module, siteKey, true] };
      createdForms += 1;
      if (logSheet) {
        logSheet.appendRow(["FORM", siteKey, target.label, formBundle.url, formBundle.id, new Date()]);
      }
      return;
    }

    const rowValues = existing.row.slice();
    let changed = false;
    if (String(rowValues[1] || "").trim() !== target.label) {
      rowValues[1] = target.label;
      changed = true;
    }
    if (String(rowValues[3] || "").trim().toLowerCase() !== module) {
      rowValues[3] = module;
      changed = true;
    }
    if (String(rowValues[4] || "").trim() !== siteKey) {
      rowValues[4] = siteKey;
      changed = true;
    }
    if (!isTruthy_(rowValues[5])) {
      rowValues[5] = true;
      changed = true;
    }

    const currentUrl = String(rowValues[2] || "").trim();
    if (!currentUrl) {
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      const formBundle = target.builder(folder, siteKey, workbook);
      attachFormToSpreadsheet_(formBundle.form, workbook, target.raw, true, false);
      setStoredFormId_(target.key, formBundle.id);
      rowValues[2] = formBundle.url;
      changed = true;
      createdForms += 1;
      if (logSheet) {
        logSheet.appendRow(["FORM", siteKey, target.label, formBundle.url, formBundle.id, new Date()]);
      }
    }

    if (changed) {
      formSheet.getRange(existing.rowIndex, 1, 1, 6).setValues([rowValues]);
      formMap[target.key] = { rowIndex: existing.rowIndex, row: rowValues };
      updatedRows += 1;
    }
  });

  if (rowsToAppend.length) {
    formSheet.getRange(formSheet.getLastRow() + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
  }

  return { checkedSources, createdForms, updatedRows };
}

function applyInPlaceSystemUpdate() {
  const dashboard = resolveAdminDashboard_("applyInPlaceSystemUpdate");
  setStoredDashboardId_(dashboard.getId());
  ensureAccessControlSheet_(dashboard);

  const recoveredRaw = recoverRawDataFromBackupsFromDashboard_();
  const itemCreationForms = ensureItemCreationFormsFromDashboard_();
  const forms = repairFormConnectionsFromDashboard();
  const formsUx = refreshFormsUxFromDashboard();
  const syncItems = syncAllModuleItemsFromDashboard();
  const modules = reapplyModuleFormulasFromDashboard();
  const triggers = installModuleFormSubmitTriggersFromDashboard();
  const liveRefresh = installLiveFormRefreshTriggerFromDashboard();
  buildImportRangeAuthorizationHelper();
  const sourcesCount = refreshDashboardConsolidations({ skipSplitSync: true });
  const liveChoices = refreshAllFormsLiveChoices_();
  refreshSystemPilotage({ skipSplitSync: true });
  const splitSync = propagateGlobalUpdateToSplitDashboards_(dashboard, new Date());
  const visibility = applyUserVisibilityModeFromDashboard();
  const health = runSystemHealthCheckFromDashboard();
  const previousComponentsSignature = getAppliedComponentsSignature_();
  const componentsSignature = getSystemComponentSignature_();
  const componentsCount = Object.keys(getSystemComponentVersions_()).length;
  setAppliedCodeVersion_(SYSTEM_CODE_VERSION);
  setAppliedComponentsSignature_(componentsSignature);
  logSystemUpdate_(dashboard, health.errors, previousComponentsSignature);

  return {
    version: SYSTEM_CODE_VERSION,
    componentsCount,
    componentsSignature,
    recoveredRawSheets: recoveredRaw.recoveredSheets,
    recoveredRawRows: recoveredRaw.recoveredRows,
    formsChecked: forms.length - 1,
    itemCreationFormsCreated: itemCreationForms.createdForms,
    itemCreationFormsUpdated: itemCreationForms.updatedRows,
    formsUxUpdated: formsUx.updatedForms,
    itemsAdded: syncItems.totalAdded,
    modulesChecked: modules.length - 1,
    createdTriggers: triggers.createdTriggers,
    liveRefreshTriggers: liveRefresh.createdTriggers,
    liveChoicesUpdated: liveChoices.updatedForms,
    activeSources: sourcesCount,
    userDashboardUrl: splitSync.userDashboard.url,
    controllerDashboardUrl: splitSync.controllerDashboard.url,
    visibilityMode: visibility.mode,
    healthErrors: health.errors,
  };
}

function runAdminFullRecoveryAfterUpdate_() {
  const dashboard = resolveAdminDashboard_("Remise en service complete");
  const access = assertAdmin_(dashboard, "Remise en service complete");
  setStoredDashboardId_(dashboard.getId());

  const update = applyInPlaceSystemUpdate();
  const adminVisibility = applyAdminVisibilityModeFromDashboard();
  const health = runSystemHealthCheckFromDashboard();

  return {
    executedBy: access.email || "",
    version: update.version,
    update,
    finalVisibilityMode: adminVisibility.mode,
    finalHealthErrors: health.errors,
  };
}

function pushGlobalUpdateFromDashboard() {
  const dashboard = resolveAdminDashboard_("pushGlobalUpdateFromDashboard");
  const splitSync = propagateGlobalUpdateToSplitDashboards_(dashboard, new Date());
  return {
    dashboardUrl: dashboard.getUrl(),
    userDashboardUrl: splitSync.userDashboard.url,
    controllerDashboardUrl: splitSync.controllerDashboard.url,
    syncedAt: splitSync.syncedAt,
  };
}

function syncSplitUserDashboardFromActiveAdmin() {
  const adminDashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!adminDashboard || !isDashboardSpreadsheet_(adminDashboard)) {
    throw new Error("Ouvrez le dashboard admin avant de synchroniser le dashboard user.");
  }
  setStoredDashboardId_(adminDashboard.getId());
  const folder = getOrCreateFolder_(DEPLOYMENT_CONFIG.rootFolderName);
  return syncSplitUserDashboard_(folder, adminDashboard, new Date());
}

function syncSplitControllerDashboardFromActiveAdmin() {
  const adminDashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!adminDashboard || !isDashboardSpreadsheet_(adminDashboard)) {
    throw new Error("Ouvrez le dashboard admin avant de synchroniser le dashboard controleur.");
  }
  setStoredDashboardId_(adminDashboard.getId());
  const folder = getOrCreateFolder_(DEPLOYMENT_CONFIG.rootFolderName);
  return syncSplitControllerDashboard_(folder, adminDashboard, new Date());
}

function propagateGlobalUpdateToSplitDashboards_(dashboard, now, folder) {
  const adminDashboard = dashboard && dashboard.getSheetByName("CONFIG_SOURCES")
    ? dashboard
    : resolveAdminDashboard_("propagateGlobalUpdateToSplitDashboards_");
  const targetFolder = folder || getOrCreateFolder_(DEPLOYMENT_CONFIG.rootFolderName);
  const syncedAt = now || new Date();
  const userDashboard = syncSplitUserDashboard_(targetFolder, adminDashboard, syncedAt);
  const controllerDashboard = syncSplitControllerDashboard_(targetFolder, adminDashboard, syncedAt);
  return { userDashboard, controllerDashboard, syncedAt };
}

function syncSplitUserDashboard_(folder, adminDashboard, now) {
  if (!adminDashboard) throw new Error("Dashboard admin introuvable.");
  const targetFolder = folder || getOrCreateFolder_(DEPLOYMENT_CONFIG.rootFolderName);
  const userDashboard = getOrCreateUserDashboard_(targetFolder);
  setStoredUserDashboardId_(userDashboard.spreadsheet.getId());

  ensureWorkbookTabs_(userDashboard.spreadsheet, USER_DASHBOARD_TABS);
  applyUserDashboardMirrorFormulas_(userDashboard.spreadsheet, adminDashboard);
  buildKpiOverviewAnalytics_(userDashboard.spreadsheet);
  buildUserSystemPilotage_(userDashboard.spreadsheet);
  protectUserDashboard_(userDashboard.spreadsheet);
  setVisibleTabsOnly_(userDashboard.spreadsheet, USER_DASHBOARD_VISIBLE_TABS);
  focusSystemPilotage_(userDashboard.spreadsheet);
  applyFileAccess_(userDashboard.file, DEPLOYMENT_CONFIG.adminEditors, DEPLOYMENT_CONFIG.dashboardViewers);
  applyUserDashboardAccessFromAccessControl_(adminDashboard, userDashboard.file);

  const logSheet = adminDashboard.getSheetByName("DEPLOYMENT_LOG");
  if (logSheet) {
    logSheet.appendRow([
      "USER_DASHBOARD",
      "GLOBAL",
      userDashboard.spreadsheet.getName(),
      userDashboard.spreadsheet.getUrl(),
      userDashboard.spreadsheet.getId(),
      now || new Date(),
    ]);
  }

  return {
    id: userDashboard.spreadsheet.getId(),
    url: userDashboard.spreadsheet.getUrl(),
    created: userDashboard.created,
  };
}

function syncSplitControllerDashboard_(folder, adminDashboard, now) {
  if (!adminDashboard) throw new Error("Dashboard admin introuvable.");
  const targetFolder = folder || getOrCreateFolder_(DEPLOYMENT_CONFIG.rootFolderName);
  const controllerDashboard = getOrCreateControllerDashboard_(targetFolder);
  setStoredControllerDashboardId_(controllerDashboard.spreadsheet.getId());

  ensureWorkbookTabs_(controllerDashboard.spreadsheet, CONTROLLER_DASHBOARD_TABS);
  applyUserDashboardMirrorFormulas_(controllerDashboard.spreadsheet, adminDashboard);
  buildKpiOverviewAnalytics_(controllerDashboard.spreadsheet);
  applyControllerInventoryControlFormula_(controllerDashboard.spreadsheet);
  buildControllerSystemPilotage_(controllerDashboard.spreadsheet);
  protectControllerDashboard_(controllerDashboard.spreadsheet);
  setVisibleTabsOnly_(controllerDashboard.spreadsheet, CONTROLLER_DASHBOARD_VISIBLE_TABS);
  focusSystemPilotage_(controllerDashboard.spreadsheet);
  applyFileAccess_(controllerDashboard.file, DEPLOYMENT_CONFIG.adminEditors, []);
  applyControllerDashboardAccessFromAccessControl_(adminDashboard, controllerDashboard.file);
  enforceControllerDashboardPermissions_(adminDashboard, controllerDashboard.file);

  const logSheet = adminDashboard.getSheetByName("DEPLOYMENT_LOG");
  if (logSheet) {
    logSheet.appendRow([
      "CONTROLLER_DASHBOARD",
      "GLOBAL",
      controllerDashboard.spreadsheet.getName(),
      controllerDashboard.spreadsheet.getUrl(),
      controllerDashboard.spreadsheet.getId(),
      now || new Date(),
    ]);
  }

  return {
    id: controllerDashboard.spreadsheet.getId(),
    url: controllerDashboard.spreadsheet.getUrl(),
    created: controllerDashboard.created,
  };
}

function getOrCreateUserDashboard_(folder) {
  const storedId = getStoredUserDashboardId_();
  if (storedId) {
    try {
      const spreadsheet = getCachedSpreadsheetById_(storedId);
      const file = getCachedFileById_(storedId);
      setStoredUserDashboardId_(spreadsheet.getId());
      return { spreadsheet, file, created: false };
    } catch (error) {
      Logger.log(`User dashboard id stale (${storedId}): ${String(error.message || error)}`);
    }
  }

  const created = createWorkbook_(folder, "DASHBOARD_USER", USER_DASHBOARD_TABS);
  setStoredUserDashboardId_(created.id);
  return { spreadsheet: created.spreadsheet, file: created.file, created: true };
}

function getOrCreateControllerDashboard_(folder) {
  const storedId = getStoredControllerDashboardId_();
  if (storedId) {
    try {
      const spreadsheet = getCachedSpreadsheetById_(storedId);
      const file = getCachedFileById_(storedId);
      setStoredControllerDashboardId_(spreadsheet.getId());
      return { spreadsheet, file, created: false };
    } catch (error) {
      Logger.log(`Controller dashboard id stale (${storedId}): ${String(error.message || error)}`);
    }
  }

  const created = createWorkbook_(folder, "DASHBOARD_CONTROLEUR", CONTROLLER_DASHBOARD_TABS);
  setStoredControllerDashboardId_(created.id);
  return { spreadsheet: created.spreadsheet, file: created.file, created: true };
}

function ensureWorkbookTabs_(spreadsheet, tabsMap) {
  const tabNames = Object.keys(tabsMap);
  tabNames.forEach((tabName) => {
    let sheet = spreadsheet.getSheetByName(tabName);
    if (!sheet) sheet = spreadsheet.insertSheet(tabName);
    ensureSheetHeaders_(sheet, tabsMap[tabName]);
  });
}

function ensureSheetHeaders_(sheet, headers) {
  if (!sheet || !headers || !headers.length) return;
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0]
    .map((value) => String(value || "").trim());
  const expected = headers.map((value) => String(value || "").trim());
  if (current.join("|") !== expected.join("|")) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  } else {
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  sheet.setFrozenRows(1);
}

function applyUserDashboardMirrorFormulas_(userDashboard, adminDashboard) {
  const adminUrl = adminDashboard.getUrl();
  if (!adminUrl) throw new Error("URL dashboard admin manquante.");

  const formulas = [
    ["FORM_LINKS", `=IFERROR(QUERY(IMPORTRANGE("${adminUrl}";"FORM_LINKS!A2:F");"select * where Col6=true";0);"")`],
    ["KPI_OVERVIEW", `=IFERROR(QUERY(IMPORTRANGE("${adminUrl}";"KPI_OVERVIEW!A2:E");"select * where Col1 is not null";0);"")`],
    ["STOCK_CONSOLIDE", `=IFERROR(QUERY(IMPORTRANGE("${adminUrl}";"STOCK_CONSOLIDE!A2:L");"select * where Col1 is not null";0);"")`],
    ["ALERTS_CONSOLIDE", `=IFERROR(QUERY(IMPORTRANGE("${adminUrl}";"ALERTS_CONSOLIDE!A2:K");"select * where Col1 is not null";0);"")`],
    ["LOTS_CONSOLIDE", `=IFERROR(QUERY(IMPORTRANGE("${adminUrl}";"LOTS_CONSOLIDE!A2:H");"select * where Col1 is not null";0);"")`],
    ["PURCHASE_CONSOLIDE", `=IFERROR(QUERY(IMPORTRANGE("${adminUrl}";"PURCHASE_CONSOLIDE!A2:L");"select * where Col1 is not null";0);"")`],
  ];

  formulas.forEach((entry) => {
    const sheetName = entry[0];
    const formula = entry[1];
    const sheet = userDashboard.getSheetByName(sheetName);
    if (!sheet) return;
    const changed = setFormulaOnSheetIfChanged_(sheet, "A2", formula);
    if (!changed) return;
    const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
    const maxCols = sheet.getMaxColumns();
    if (maxRows > 1 || maxCols > 1) {
      sheet.getRange(2, 2, maxRows, Math.max(maxCols - 1, 1)).clearContent();
    }
  });
}

function buildUserSystemPilotage_(spreadsheet, options) {
  const sheet = spreadsheet.getSheetByName("SYSTEM_PILOTAGE");
  if (!sheet) return;
  const forceRebuild = !!(options && options.forceRebuild);
  const isLayoutCurrent = !forceRebuild
    && isLayoutMarkerCurrent_(spreadsheet, "SYSTEM_PILOTAGE_USER", LAYOUT_VERSIONS.USER_PILOTAGE)
    && isSystemPilotageLayoutReady_(sheet, "USER");
  if (isLayoutCurrent) {
    setupPilotageToggleControls_(sheet, "USER");
    setupPilotageNativeActionControls_(sheet, "USER");
    applyPilotageSectionVisibility_(sheet, "USER");
    return;
  }

  const totalRows = 56;
  const totalCols = 16;
  if (sheet.getMaxRows() < totalRows) sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < totalCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), totalCols - sheet.getMaxColumns());

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(2);
  sheet.setRowHeights(1, totalRows, 24);

  const widths = [180, 130, 130, 24, 120, 120, 120, 120, 120, 120, 120, 120, 120, 120, 120, 120];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));

  sheet.getRange(1, 1, totalRows, totalCols)
    .setBackground("#071522")
    .setFontColor("#eaf2ff")
    .setFontFamily("Trebuchet MS")
    .setVerticalAlignment("middle");

  sheet.getRange("A1:C56").setBackground("#04101c");
  sheet.getRange("D1:D56").setBackground("#071522");

  sheet.getRange("A1:C1").merge()
    .setValue("CONTROL CENTER")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#0a2844");
  sheet.getRange("A2:C2").merge()
    .setValue("Vue utilisateur")
    .setFontSize(9)
    .setHorizontalAlignment("left")
    .setBackground("#0b2238");

  sheet.getRange("E1:P1").merge()
    .setValue("SYSTEM PILOTAGE | Exploitation quotidienne")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#0d3559");
  sheet.getRange("E2:P2").merge()
    .setValue("Stock, alertes, achats et formulaires (mode utilisateur)")
    .setFontSize(10)
    .setHorizontalAlignment("left")
    .setBackground("#11314f");

  setupPilotageToggleControls_(sheet, "USER");

  stylePanel_(sheet, 4, 20, 1, 3, "Navigation rapide", "#0a1f33", "#11466f");
  stylePanel_(sheet, 22, 38, 1, 3, "Formulaires", "#0a1f33", "#11466f");
  stylePanel_(sheet, 40, 56, 1, 3, "Infos", "#0a1f33", "#11466f");
  sheet.getRange("A23:C23").merge()
    .setValue("Natif HTML: Operations > Formulaires natifs > Ouvrir panneau boutons")
    .setFontSize(8)
    .setHorizontalAlignment("left")
    .setBackground("#0c2f4f");

  stylePanel_(sheet, 4, 12, 5, 16, "Indicateurs globaux", "#0a2236", "#11466f");
  stylePanel_(sheet, 14, 30, 5, 10, "Stock critique", "#0b243b", "#11466f");
  stylePanel_(sheet, 14, 30, 11, 16, "Demandes achat", "#0b243b", "#11466f");
  stylePanel_(sheet, 32, 48, 5, 10, "Alertes", "#0b243b", "#11466f");
  stylePanel_(sheet, 32, 48, 11, 16, "Lots pharmacie", "#0b243b", "#11466f");
  stylePanel_(sheet, 50, 56, 5, 16, "Suivi stock actuel", "#0b243b", "#11466f");

  const gidStock = spreadsheet.getSheetByName("STOCK_CONSOLIDE").getSheetId();
  const gidAlerts = spreadsheet.getSheetByName("ALERTS_CONSOLIDE").getSheetId();
  const gidLots = spreadsheet.getSheetByName("LOTS_CONSOLIDE").getSheetId();
  const gidPurchase = spreadsheet.getSheetByName("PURCHASE_CONSOLIDE").getSheetId();
  const gidKpi = spreadsheet.getSheetByName("KPI_OVERVIEW").getSheetId();

  const navFormulas = [
    `=HYPERLINK("#gid=${gidKpi}";"Ouvrir KPI")`,
    `=HYPERLINK("#gid=${gidStock}";"Ouvrir stock consolide")`,
    `=HYPERLINK("#gid=${gidAlerts}";"Ouvrir alertes")`,
    `=HYPERLINK("#gid=${gidLots}";"Ouvrir lots pharmacie")`,
    `=HYPERLINK("#gid=${gidPurchase}";"Ouvrir achats")`,
  ];
  const navRows = [6, 8, 10, 12, 14];
  navRows.forEach((row, index) => drawPilotageButton_(sheet, row, 1, 3, navFormulas[index], "#1f5f96"));

  const formButtons = isNativeFormsMode_()
    ? [
      '=IFERROR(INDEX(FORM_LINKS!B2:B;1)&" (menu Operations > Formulaires natifs)";"Formulaire #1 indisponible")',
      '=IFERROR(INDEX(FORM_LINKS!B2:B;2)&" (menu Operations > Formulaires natifs)";"Formulaire #2 indisponible")',
      '=IFERROR(INDEX(FORM_LINKS!B2:B;3)&" (menu Operations > Formulaires natifs)";"Formulaire #3 indisponible")',
      '=IFERROR(INDEX(FORM_LINKS!B2:B;4)&" (menu Operations > Formulaires natifs)";"Formulaire #4 indisponible")',
      '=IFERROR(INDEX(FORM_LINKS!B2:B;5)&" (menu Operations > Formulaires natifs)";"Formulaire #5 indisponible")',
      '=IFERROR(INDEX(FILTER(FORM_LINKS!B2:B;LEFT(FORM_LINKS!A2:A;24)="FORM_PHARMA_ITEM_CREATE_");1)&" (menu Operations > Formulaires natifs)";"Form creation article pharmacie indisponible")',
    ]
    : [
      '=IFERROR(HYPERLINK(INDEX(FORM_LINKS!C2:C;1);INDEX(FORM_LINKS!B2:B;1));"Formulaire #1 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FORM_LINKS!C2:C;2);INDEX(FORM_LINKS!B2:B;2));"Formulaire #2 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FORM_LINKS!C2:C;3);INDEX(FORM_LINKS!B2:B;3));"Formulaire #3 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FORM_LINKS!C2:C;4);INDEX(FORM_LINKS!B2:B;4));"Formulaire #4 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FORM_LINKS!C2:C;5);INDEX(FORM_LINKS!B2:B;5));"Formulaire #5 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FILTER(FORM_LINKS!C2:C;LEFT(FORM_LINKS!A2:A;24)="FORM_PHARMA_ITEM_CREATE_");1);INDEX(FILTER(FORM_LINKS!B2:B;LEFT(FORM_LINKS!A2:A;24)="FORM_PHARMA_ITEM_CREATE_");1));"Form creation article pharmacie indisponible")',
    ];
  [24, 27, 30, 33, 36, 38].forEach((row, index) => drawPilotageButton_(sheet, row, 1, 3, formButtons[index], "#296ea8"));
  setupPilotageNativeActionControls_(sheet, "USER");

  drawPilotageCard_(sheet, 5, 5, 8, "Articles actifs", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("TOTAL_ITEMS";KPI_OVERVIEW!A:A;0));0)', "#2d9cdb");
  drawPilotageCard_(sheet, 5, 9, 12, "Sous-seuil", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("LOW_STOCK";KPI_OVERVIEW!A:A;0));0)', "#f9a825");
  drawPilotageCard_(sheet, 5, 13, 16, "Ruptures", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("RUPTURE";KPI_OVERVIEW!A:A;0));0)', "#ef5350");
  drawPilotageCard_(sheet, 9, 5, 8, "Stock actuel total", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("CURRENT_STOCK_TOTAL";KPI_OVERVIEW!A:A;0));0)', "#1e88e5");
  drawPilotageCard_(sheet, 9, 9, 12, "Achats ouverts", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("PURCHASE_OPEN";KPI_OVERVIEW!A:A;0));0)', "#4db6ac");
  drawPilotageCard_(sheet, 9, 13, 16, "Peremption 30j", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("EXPIRY_30D";KPI_OVERVIEW!A:A;0));0)', "#8e24aa");

  sheet.getRange("E16:J16").setValues([["ItemID", "Article", "Stock", "Seuil", "Ecart", "Statut"]]).setFontWeight("bold").setBackground("#14456f");
  sheet.getRange("E17").setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(STOCK_CONSOLIDE!A2:A;STOCK_CONSOLIDE!B2:B;STOCK_CONSOLIDE!F2:F;STOCK_CONSOLIDE!E2:E;STOCK_CONSOLIDE!G2:G;STOCK_CONSOLIDE!H2:H);STOCK_CONSOLIDE!H2:H<>"OK";STOCK_CONSOLIDE!A2:A<>"");5;FALSE);""));14;6)');

  sheet.getRange("K16:P16").setValues([["RequestID", "ItemID", "Article", "Qte", "Priorite", "Statut"]]).setFontWeight("bold").setBackground("#14456f");
  sheet.getRange("K17").setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(PURCHASE_CONSOLIDE!A2:A;PURCHASE_CONSOLIDE!D2:D;PURCHASE_CONSOLIDE!E2:E;PURCHASE_CONSOLIDE!G2:G;PURCHASE_CONSOLIDE!H2:H;PURCHASE_CONSOLIDE!L2:L);PURCHASE_CONSOLIDE!A2:A<>"";(PURCHASE_CONSOLIDE!L2:L="A_VALIDER")+(PURCHASE_CONSOLIDE!L2:L="EN_COURS"));4;FALSE);""));14;6)');

  sheet.getRange("E34:J34").setValues([["AlertID", "Type", "Severite", "ItemID", "Article", "Date"]]).setFontWeight("bold").setBackground("#14456f");
  sheet.getRange("E35").setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(ALERTS_CONSOLIDE!A2:A;ALERTS_CONSOLIDE!D2:D;ALERTS_CONSOLIDE!E2:E;ALERTS_CONSOLIDE!F2:F;ALERTS_CONSOLIDE!G2:G;ALERTS_CONSOLIDE!I2:I);ALERTS_CONSOLIDE!A2:A<>"");6;FALSE);""));14;6)');

  sheet.getRange("K34:P34").setValues([["LotID", "ItemID", "Site", "Lot", "Peremption", "Statut"]]).setFontWeight("bold").setBackground("#14456f");
  sheet.getRange("K35").setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(LOTS_CONSOLIDE!A2:A;LOTS_CONSOLIDE!B2:B;LOTS_CONSOLIDE!C2:C;LOTS_CONSOLIDE!D2:D;LOTS_CONSOLIDE!E2:E;LOTS_CONSOLIDE!H2:H);LOTS_CONSOLIDE!A2:A<>"");5;TRUE);""));14;6)');

  sheet.getRange("A42:C42").merge().setFormula('="Formulaires actifs: "&COUNTA(FORM_LINKS!A2:A)');
  sheet.getRange("A44:C44").merge().setFormula('="Alertes ouvertes: "&COUNTIF(ALERTS_CONSOLIDE!J2:J;"OPEN")');
  sheet.getRange("A46:C46").merge().setFormula('="Achats ouverts: "&(COUNTIFS(PURCHASE_CONSOLIDE!L2:L;"A_VALIDER")+COUNTIFS(PURCHASE_CONSOLIDE!L2:L;"EN_COURS"))');
  sheet.getRange("A48:C48").merge().setFormula('="Lots expires: "&IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("EXPIRY_PAST";KPI_OVERVIEW!A:A;0));0)');
  sheet.getRange("A50:C50").merge().setFormula('="Derniere synchro: "&TEXT(NOW();"dd/mm/yyyy hh:mm")');

  sheet.getRange("E52:N52")
    .setValues([["ItemID", "Article", "Module", "Site", "Stock", "Seuil", "Ecart", "Statut", "Dernier mouv.", "Couv."]])
    .setFontWeight("bold")
    .setBackground("#14456f")
    .setHorizontalAlignment("center");
  sheet.getRange("E53").setFormula(buildStockTrackingArrayFormula_(4));
  drawPilotageButton_(sheet, 53, 15, 16, `=HYPERLINK("#gid=${gidStock}";"Ouvrir stock")`, "#1f5f96");
  drawPilotageButton_(sheet, 55, 15, 16, `=HYPERLINK("#gid=${gidPurchase}";"Ouvrir achats")`, "#1f5f96");

  sheet.getRange("E16:P56").setFontSize(9);
  sheet.getRange("E17:P56").setHorizontalAlignment("left");
  sheet.getRange("M53:M56").setNumberFormat("dd/mm/yyyy hh:mm");
  sheet.getRange("N53:N56").setNumberFormat("0.00");
  sheet.getRange("E16:P16").setHorizontalAlignment("center");
  sheet.getRange("E34:P34").setHorizontalAlignment("center");
  applyPilotageSectionVisibility_(sheet, "USER");
  setLayoutMarkerVersion_(spreadsheet, "SYSTEM_PILOTAGE_USER", LAYOUT_VERSIONS.USER_PILOTAGE);
}

function protectUserDashboard_(spreadsheet) {
  protectSheets_(spreadsheet, [
    "SYSTEM_PILOTAGE",
    "FORM_LINKS",
    "KPI_OVERVIEW",
    "STOCK_CONSOLIDE",
    "ALERTS_CONSOLIDE",
    "LOTS_CONSOLIDE",
    "PURCHASE_CONSOLIDE",
  ]);
}

function applyControllerInventoryControlFormula_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("INVENTORY_CONTROL");
  if (!sheet) return;
  const changed = setFormulaOnSheetIfChanged_(
    sheet,
    "A2",
    '=ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(STOCK_CONSOLIDE!A2:A;STOCK_CONSOLIDE!B2:B;STOCK_CONSOLIDE!F2:F;STOCK_CONSOLIDE!E2:E;STOCK_CONSOLIDE!G2:G;STOCK_CONSOLIDE!H2:H;STOCK_CONSOLIDE!K2:K;STOCK_CONSOLIDE!L2:L);STOCK_CONSOLIDE!A2:A<>"";STOCK_CONSOLIDE!H2:H<>"OK");5;FALSE);""))'
  );
  if (!changed) return;
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  const maxCols = sheet.getMaxColumns();
  if (maxRows > 1 || maxCols > 1) {
    sheet.getRange(2, 2, maxRows, Math.max(maxCols - 1, 1)).clearContent();
  }
}

function buildControllerSystemPilotage_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("SYSTEM_PILOTAGE");
  if (!sheet) return;
  const inventorySheet = spreadsheet.getSheetByName("INVENTORY_CONTROL");
  const isLayoutCurrent = isLayoutMarkerCurrent_(spreadsheet, "SYSTEM_PILOTAGE_CONTROLLER", LAYOUT_VERSIONS.CONTROLLER_PILOTAGE)
    && isSystemPilotageLayoutReady_(sheet, "CONTROLLER");
  if (isLayoutCurrent) {
    setupPilotageToggleControls_(sheet, "CONTROLLER");
    setupPilotageNativeActionControls_(sheet, "CONTROLLER");
    applyPilotageSectionVisibility_(sheet, "CONTROLLER");
    return;
  }

  buildUserSystemPilotage_(spreadsheet, { forceRebuild: true });

  sheet.getRange("A2:C2").merge().setValue("Vue controleur");
  sheet.getRange("E1:P1").merge().setValue("SYSTEM PILOTAGE | Controle inventaire");
  sheet.getRange("E2:P2").merge().setValue("Pilotage controleur: inventaires, ruptures, alertes et formulaires");
  setupPilotageToggleControls_(sheet, "CONTROLLER");
  setupPilotageNativeActionControls_(sheet, "CONTROLLER");
  sheet.getRange("A44:C44").merge().setFormula('="Ecarts inventaire (articles non OK): "&COUNTA(INVENTORY_CONTROL!A2:A)');
  sheet.getRange("K16:P16").setValues([["ItemID", "Article", "Stock", "Seuil", "Ecart", "Statut"]]).setFontWeight("bold").setBackground("#14456f");
  sheet.getRange("K17").setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(INVENTORY_CONTROL!A2:F;""));14;6)');
  sheet.getRange("E53").setFormula(buildStockTrackingArrayFormula_(4));
  sheet.getRange("M53:M56").setNumberFormat("dd/mm/yyyy hh:mm");
  sheet.getRange("N53:N56").setNumberFormat("0.00");
  if (inventorySheet) {
    drawPilotageButton_(sheet, 55, 15, 16, `=HYPERLINK("#gid=${inventorySheet.getSheetId()}";"Ouvrir inventaire")`, "#1f5f96");
  }
  applyPilotageSectionVisibility_(sheet, "CONTROLLER");
  setLayoutMarkerVersion_(spreadsheet, "SYSTEM_PILOTAGE_CONTROLLER", LAYOUT_VERSIONS.CONTROLLER_PILOTAGE);
}

function protectControllerDashboard_(spreadsheet) {
  protectSheets_(spreadsheet, [
    "SYSTEM_PILOTAGE",
    "FORM_LINKS",
    "KPI_OVERVIEW",
    "INVENTORY_CONTROL",
    "STOCK_CONSOLIDE",
    "ALERTS_CONSOLIDE",
    "LOTS_CONSOLIDE",
    "PURCHASE_CONSOLIDE",
  ]);
}

function logSystemUpdate_(dashboard, healthErrors, previousAppliedComponentsSignature) {
  const logSheet = dashboard.getSheetByName("DEPLOYMENT_LOG");
  if (!logSheet) return;
  const componentsSignature = getSystemComponentSignature_();
  const componentsCount = Object.keys(getSystemComponentVersions_()).length;
  const baselineSignature = String(previousAppliedComponentsSignature || getAppliedComponentsSignature_());
  const componentsSummary = formatSystemComponentDeltaSummary_(baselineSignature, componentsSignature, 3);
  logSheet.appendRow([
    "UPDATE",
    "GLOBAL",
    `System update ${SYSTEM_CODE_VERSION}`,
    dashboard.getUrl(),
    dashboard.getId(),
    new Date(),
  ]);
  logSheet.appendRow([
    "COMPONENTS",
    "GLOBAL",
    `Components tracked: ${componentsCount} | ${componentsSummary}`,
    dashboard.getUrl(),
    componentsSignature,
    new Date(),
  ]);
  logSheet.appendRow([
    "HEALTH",
    "GLOBAL",
    `Health errors: ${healthErrors}`,
    dashboard.getUrl(),
    dashboard.getId(),
    new Date(),
  ]);
}

function applyUserVisibilityModeFromDashboard() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard) throw new Error("Aucun classeur actif.");
  return applyVisibilityMode_(dashboard, "user");
}

function applyAdminVisibilityModeFromDashboard() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard) throw new Error("Aucun classeur actif.");
  return applyVisibilityMode_(dashboard, "admin");
}

function applyVisibilityMode_(dashboard, mode) {
  const isUser = mode === "user";
  ensureAuditSheetsExist_(dashboard);
  ensureAccessControlSheet_(dashboard);

  if (isUser) {
    setVisibleTabsOnly_(dashboard, DASHBOARD_USER_VISIBLE_TABS);
  } else {
    showAllSheets_(dashboard);
  }

  let processed = 0;
  forEachActiveSourceWorkbook_(dashboard, (ctx) => {
    const visibleTabs = ctx.module === "pharmacie" ? PHARMA_USER_VISIBLE_TABS : FIRE_USER_VISIBLE_TABS;
    if (isUser) {
      setVisibleTabsOnly_(ctx.workbook, visibleTabs);
    } else {
      showAllSheets_(ctx.workbook);
    }
    processed += 1;
  });

  focusSystemPilotage_(dashboard);

  return { mode, modulesProcessed: processed };
}

function setVisibleTabsOnly_(spreadsheet, visibleTabs) {
  const visibleSet = {};
  visibleTabs.forEach((name) => {
    visibleSet[name] = true;
  });

  const sheets = spreadsheet.getSheets();
  sheets.forEach((sheet) => {
    if (visibleSet[sheet.getName()]) sheet.showSheet();
  });

  const keep = sheets.find((sheet) => visibleSet[sheet.getName()]) || sheets[0];
  spreadsheet.setActiveSheet(keep);

  sheets.forEach((sheet) => {
    if (visibleSet[sheet.getName()]) return;
    if (sheet.getSheetId() === keep.getSheetId()) return;
    sheet.hideSheet();
  });
}

function showAllSheets_(spreadsheet) {
  spreadsheet.getSheets().forEach((sheet) => sheet.showSheet());
}

function forEachActiveSourceWorkbook_(dashboard, callback) {
  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet || sourceSheet.getLastRow() < 2) return;

  const rows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  rows.forEach((row) => {
    const sourceKey = String(row[0] || "").trim();
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];
    if (!sourceKey || !module || !workbookUrl || !isTruthy_(enabled)) return;

    const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
    callback({ sourceKey, module, siteKey, workbook, workbookUrl });
  });
}

function refreshSystemPilotage(options) {
  const spreadsheet = resolveAdminDashboard_("refreshSystemPilotage");
  setStoredDashboardId_(spreadsheet.getId());
  if (!spreadsheet.getSheetByName("SYSTEM_PILOTAGE")) {
    throw new Error("L'onglet SYSTEM_PILOTAGE est introuvable dans le classeur actif.");
  }

  ensureAccessControlSheet_(spreadsheet);
  buildKpiOverviewAnalytics_(spreadsheet);
  buildSystemPilotage_(spreadsheet);
  focusSystemPilotage_(spreadsheet);
  protectDashboard_(spreadsheet);
  let splitSync = null;
  if (!(options && options.skipSplitSync)) {
    splitSync = propagateGlobalUpdateToSplitDashboards_(spreadsheet, new Date());
  }
  SpreadsheetApp.flush();
  Logger.log(`Pilotage refreshed: ${spreadsheet.getUrl()}`);
  return {
    dashboardUrl: spreadsheet.getUrl(),
    userDashboardUrl: splitSync ? splitSync.userDashboard.url : "",
    controllerDashboardUrl: splitSync ? splitSync.controllerDashboard.url : "",
  };
}

function repairFormConnectionsFromDashboard() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard) {
    throw new Error("Aucun classeur actif. Ouvrez le dashboard puis relancez repairFormConnectionsFromDashboard.");
  }
  setStoredDashboardId_(dashboard.getId());
  if (isNativeFormsMode_()) {
    const sync = ensureNativeFormLinksFromDashboard_(dashboard);
    const formSheetNative = dashboard.getSheetByName("FORM_LINKS");
    const formRowsNative = formSheetNative && formSheetNative.getLastRow() > 1
      ? formSheetNative.getRange(2, 1, formSheetNative.getLastRow() - 1, 6).getValues()
      : [];
    const auditNative = [["FormKey", "Status", "Detail", "FormUrl"]];
    formRowsNative.forEach((row) => {
      const formKey = String(row[0] || "").trim();
      const formUrl = String(row[2] || "").trim();
      if (!formKey || !isTruthy_(row[5])) return;
      auditNative.push([
        formKey,
        isNativeFormReference_(formUrl) ? "OK" : "WARN",
        "Mode natif HTML actif. Utiliser Operations > Formulaires natifs (HTML).",
        formUrl,
      ]);
    });
    const auditSheetNative = dashboard.getSheetByName("FORM_CONNECTION_AUDIT") || dashboard.insertSheet("FORM_CONNECTION_AUDIT");
    auditSheetNative.clearContents();
    auditSheetNative.clearFormats();
    auditSheetNative.getRange(1, 1, auditNative.length, auditNative[0].length).setValues(auditNative);
    auditSheetNative.getRange(1, 1, 1, auditNative[0].length).setFontWeight("bold");
    auditSheetNative.autoResizeColumns(1, auditNative[0].length);
    SpreadsheetApp.flush();
    Logger.log(`Native form connection audit generated (${auditNative.length - 1} rows, sync:${JSON.stringify(sync)}).`);
    return auditNative;
  }

  const formSheet = dashboard.getSheetByName("FORM_LINKS");
  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!formSheet || !sourceSheet) {
    throw new Error("FORM_LINKS et/ou CONFIG_SOURCES introuvable(s) dans le dashboard actif.");
  }

  const sourceRows = sourceSheet.getLastRow() > 1
    ? sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues()
    : [];
  let formRows = formSheet.getLastRow() > 1
    ? formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 6).getValues()
    : [];

  const sourceMap = {};
  sourceRows.forEach((row) => {
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];
    if (!module || !siteKey || !workbookUrl) return;
    if (!isTruthy_(enabled)) return;
    sourceMap[`${module}|${siteKey}`] = workbookUrl;
  });

  const audit = [["FormKey", "Status", "Detail", "FormUrl"]];

  formRows.forEach((row) => {
    const formKey = String(row[0] || "").trim();
    const formLabel = String(row[1] || "").trim();
    const formUrl = String(row[2] || "").trim();
    const module = String(row[3] || "").trim().toLowerCase();
    const siteKey = String(row[4] || "").trim();
    const isActive = row[5];

    if (!formKey || !formUrl || !module || !siteKey || !isTruthy_(isActive)) {
      audit.push([formKey || "(vide)", "SKIPPED", "Ligne incomplete ou inactive", formUrl || ""]);
      return;
    }

    const rawSheetName = rawSheetNameForFormKey_(formKey);
    if (!rawSheetName) {
      audit.push([formKey, "ERROR", "FormKey non reconnu pour mapping RAW", formUrl]);
      return;
    }

    const workbookUrl = sourceMap[`${module}|${siteKey}`];
    if (!workbookUrl) {
      audit.push([formKey, "ERROR", `Aucune source active pour ${module}/${siteKey}`, formUrl]);
      return;
    }

    try {
      const form = openManagedForm_(dashboard, formKey, formUrl, siteKey, formLabel);
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      const result = ensureFormDestination_(form, workbook, rawSheetName, true);
      const status = result.changed ? "RECONNECTED" : "OK";
      audit.push([formKey, status, `Destination:${workbook.getId()} | RAW:${rawSheetName}`, formUrl]);
    } catch (error) {
      audit.push([formKey, "ERROR", String(error.message || error), formUrl]);
    }
  });

  const auditSheet = dashboard.getSheetByName("FORM_CONNECTION_AUDIT") || dashboard.insertSheet("FORM_CONNECTION_AUDIT");
  auditSheet.clearContents();
  auditSheet.clearFormats();
  auditSheet.getRange(1, 1, audit.length, audit[0].length).setValues(audit);
  auditSheet.getRange(1, 1, 1, audit[0].length).setFontWeight("bold");
  auditSheet.autoResizeColumns(1, audit[0].length);
  SpreadsheetApp.flush();

  Logger.log(`Form connection audit generated (${audit.length - 1} rows).`);
  return audit;
}

function refreshFormsUxFromDashboard() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard) {
    throw new Error("Aucun classeur actif. Ouvrez le dashboard puis relancez refreshFormsUxFromDashboard.");
  }
  setStoredDashboardId_(dashboard.getId());
  if (isNativeFormsMode_()) {
    const sync = ensureNativeFormLinksFromDashboard_(dashboard);
    const formSheetNative = dashboard.getSheetByName("FORM_LINKS");
    const rowsNative = formSheetNative && formSheetNative.getLastRow() > 1
      ? formSheetNative.getRange(2, 1, formSheetNative.getLastRow() - 1, 6).getValues()
      : [];
    const auditNative = [["FormKey", "Status", "Detail"]];
    rowsNative.forEach((row) => {
      const formKey = String(row[0] || "").trim();
      if (!formKey || !isTruthy_(row[5])) return;
      auditNative.push([formKey, "OK", "Mode natif HTML: UX geree par NativeFormDialog."]);
    });
    const auditSheetNative = dashboard.getSheetByName("FORM_UX_AUDIT") || dashboard.insertSheet("FORM_UX_AUDIT");
    auditSheetNative.clearContents();
    auditSheetNative.clearFormats();
    auditSheetNative.getRange(1, 1, auditNative.length, auditNative[0].length).setValues(auditNative);
    auditSheetNative.getRange(1, 1, 1, auditNative[0].length).setFontWeight("bold");
    auditSheetNative.autoResizeColumns(1, auditNative[0].length);
    SpreadsheetApp.flush();
    return { updatedForms: sync.createdRows + sync.updatedRows };
  }

  const formSheet = dashboard.getSheetByName("FORM_LINKS");
  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!formSheet || !sourceSheet || formSheet.getLastRow() < 2) {
    throw new Error("FORM_LINKS et/ou CONFIG_SOURCES introuvable(s) ou vide(s).");
  }

  const sourceRows = sourceSheet.getRange(2, 1, Math.max(sourceSheet.getLastRow() - 1, 1), 9).getValues();
  const availableSites = sourceRows
    .map((row) => String(row[2] || "").trim())
    .filter((value) => value !== "");
  const sourceMap = {};
  sourceRows.forEach((row) => {
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    if (!module || !siteKey || !workbookUrl) return;
    sourceMap[`${module}|${siteKey}`] = workbookUrl;
  });

  const rows = formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 6).getValues();
  const updatedUrls = rows.map((row) => [row[2]]);
  const updatedLabels = rows.map((row) => [row[1]]);
  const audit = [["FormKey", "Status", "Detail"]];

  rows.forEach((row, index) => {
    const formKey = String(row[0] || "").trim();
    const formLabel = String(row[1] || "").trim();
    const formUrl = String(row[2] || "").trim();
    const module = String(row[3] || "").trim().toLowerCase();
    const siteKey = String(row[4] || "").trim();
    const isActive = row[5];

    if (!formKey || !formUrl || !isTruthy_(isActive)) {
      audit.push([formKey || "(vide)", "SKIPPED", "Ligne inactive ou incomplete"]);
      return;
    }

    try {
      const computedLabel = formLabelFromKey_(formKey, siteKey);
      if (computedLabel) {
        updatedLabels[index][0] = computedLabel;
      }
      const form = openManagedForm_(dashboard, formKey, formUrl, siteKey, formLabel);
      const workbookUrl = sourceMap[`${module}|${siteKey}`];
      const workbook = workbookUrl ? getCachedSpreadsheetByUrl_(workbookUrl) : null;

      applyFriendlyFormDefinition_(form, formKey, siteKey, module, workbook, availableSites);
      const rawSheetName = rawSheetNameForFormKey_(formKey);
      if (workbook && rawSheetName) {
        // En mode MAJ UX, ne jamais forcer un reset de destination pour eviter tout risque de perte de donnees.
        ensureFormDestination_(form, workbook, rawSheetName, true, false);
      }
      const prefilled = buildSitePrefilledUrl_(form, siteKey) || form.getPublishedUrl() || form.getEditUrl();
      updatedUrls[index][0] = prefilled;
      audit.push([formKey, "OK", "Libelles/formulaire mis a jour + URL pre-remplie"]);
    } catch (error) {
      audit.push([formKey, "ERROR", String(error.message || error)]);
    }
  });

  formSheet.getRange(2, 2, updatedLabels.length, 1).setValues(updatedLabels);
  formSheet.getRange(2, 3, updatedUrls.length, 1).setValues(updatedUrls);

  const auditSheet = dashboard.getSheetByName("FORM_UX_AUDIT") || dashboard.insertSheet("FORM_UX_AUDIT");
  auditSheet.clearContents();
  auditSheet.clearFormats();
  auditSheet.getRange(1, 1, audit.length, audit[0].length).setValues(audit);
  auditSheet.getRange(1, 1, 1, audit[0].length).setFontWeight("bold");
  auditSheet.autoResizeColumns(1, audit[0].length);
  SpreadsheetApp.flush();

  return { updatedForms: audit.filter((row, i) => i > 0 && row[1] === "OK").length };
}

function refreshDashboardConsolidations(options) {
  const dashboard = resolveAdminDashboard_("refreshDashboardConsolidations");

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet) {
    throw new Error("CONFIG_SOURCES introuvable dans le dashboard actif.");
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("Aucune source active dans CONFIG_SOURCES.");
  }

  const sourceRows = sourceSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const enabledRows = sourceRows.filter((row) => isTruthy_(row[8]) && String(row[3] || "").trim() !== "");
  if (!enabledRows.length) throw new Error("Aucune source active dans CONFIG_SOURCES.");

  const first = 2;
  const last = lastRow;

  setFormulaOnSheetIfChanged_(dashboard.getSheetByName("STOCK_CONSOLIDE"), "A2", buildConsolidationFormula_(first, last, "E"));
  setFormulaOnSheetIfChanged_(dashboard.getSheetByName("ALERTS_CONSOLIDE"), "A2", buildConsolidationFormula_(first, last, "F"));
  setFormulaOnSheetIfChanged_(dashboard.getSheetByName("LOTS_CONSOLIDE"), "A2", buildConsolidationFormula_(first, last, "G"));
  applyPurchaseConsolidationFormulas_(dashboard, first, last);
  applyPurchaseValidationFormula_(dashboard);
  SpreadsheetApp.flush();

  if (!(options && options.skipSplitSync)) {
    propagateGlobalUpdateToSplitDashboards_(dashboard, new Date());
  }

  Logger.log(`Consolidations refreshed with ${enabledRows.length} active source(s).`);
  return enabledRows.length;
}

function buildImportRangeAuthorizationHelper() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active && isUserDashboardSpreadsheet_(active)) {
    return buildImportRangeAuthorizationHelperForUser_(active);
  }

  const dashboard = resolveAdminDashboard_("buildImportRangeAuthorizationHelper");
  return buildImportRangeAuthorizationHelperForAdmin_(dashboard);
}

function buildImportRangeAuthorizationHelperForAdmin_(dashboard) {
  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    throw new Error("CONFIG_SOURCES vide ou introuvable.");
  }

  const rows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  const activeRows = rows.filter((row) => isTruthy_(row[8]) && String(row[3] || "").trim() !== "");
  if (!activeRows.length) {
    throw new Error("Aucune source active pour autorisation IMPORTRANGE.");
  }

  const helper = dashboard.getSheetByName("IMPORTRANGE_AUTH") || dashboard.insertSheet("IMPORTRANGE_AUTH");
  helper.clearContents();
  helper.clearFormats();
  helper.getRange(1, 1, 1, 5).setValues([["SourceKey", "WorkbookUrl", "Range", "TestFormula", "Status"]]);
  helper.getRange(1, 1, 1, 5).setFontWeight("bold");

  const output = activeRows.map((row) => [row[0], row[3], row[4], "", ""]);
  helper.getRange(2, 1, output.length, 5).setValues(output);

  for (let r = 2; r < output.length + 2; r += 1) {
    helper.getRange(r, 4).setFormula(`=IMPORTRANGE(B${r};C${r})`);
    helper.getRange(r, 5).setFormula(`=SI(ESTERREUR(D${r});"ACTION_REQUISE";"OK")`);
  }

  helper.autoResizeColumns(1, 5);
  SpreadsheetApp.flush();
  Logger.log(`IMPORTRANGE_AUTH prepared (${output.length} source(s)).`);
  return helper.getName();
}

function buildImportRangeAuthorizationHelperForUser_(dashboard) {
  const refs = extractUserImportRangeReferences_(dashboard);
  if (!refs.length) {
    throw new Error("Aucune formule IMPORTRANGE detectee dans le dashboard user.");
  }

  const helper = dashboard.getSheetByName("IMPORTRANGE_AUTH") || dashboard.insertSheet("IMPORTRANGE_AUTH");
  helper.clearContents();
  helper.clearFormats();
  helper.getRange(1, 1, 1, 5).setValues([["SourceKey", "WorkbookUrl", "Range", "TestFormula", "Status"]]);
  helper.getRange(1, 1, 1, 5).setFontWeight("bold");

  const output = refs.map((row) => [row.key, row.url, row.range, "", ""]);
  helper.getRange(2, 1, output.length, 5).setValues(output);

  for (let r = 2; r < output.length + 2; r += 1) {
    helper.getRange(r, 4).setFormula(`=IMPORTRANGE(B${r};C${r})`);
    helper.getRange(r, 5).setFormula(`=SI(ESTERREUR(D${r});"ACTION_REQUISE";"OK")`);
  }

  helper.autoResizeColumns(1, 5);
  SpreadsheetApp.flush();
  Logger.log(`IMPORTRANGE_AUTH (user) prepared (${output.length} source(s)).`);
  return helper.getName();
}

function extractUserImportRangeReferences_(dashboard) {
  const refs = [];
  const seen = {};
  const candidates = [
    ["FORM_LINKS", "A2"],
    ["KPI_OVERVIEW", "A2"],
    ["STOCK_CONSOLIDE", "A2"],
    ["ALERTS_CONSOLIDE", "A2"],
    ["LOTS_CONSOLIDE", "A2"],
    ["PURCHASE_CONSOLIDE", "A2"],
  ];

  candidates.forEach((entry) => {
    const sheetName = entry[0];
    const a1 = entry[1];
    const sheet = dashboard.getSheetByName(sheetName);
    if (!sheet) return;
    const formula = String(sheet.getRange(a1).getFormula() || "");
    const parsed = parseImportRangeFromFormula_(formula);
    if (!parsed) return;
    const key = `${parsed.url}|${parsed.range}`;
    if (seen[key]) return;
    seen[key] = true;
    refs.push({
      key: `USER_${sheetName}`,
      url: parsed.url,
      range: parsed.range,
    });
  });

  return refs;
}

function parseImportRangeFromFormula_(formula) {
  const value = String(formula || "").trim();
  if (!value) return null;
  const match = value.match(/IMPORTRANGE\(\s*"([^"]+)"\s*;\s*"([^"]+)"\s*\)/i);
  if (!match || !match[1] || !match[2]) return null;
  return { url: String(match[1]), range: String(match[2]) };
}

function installModuleFormSubmitTriggersFromDashboard() {
  if (isNativeFormsMode_()) {
    return { createdTriggers: 0, skipped: "NATIVE_HTML" };
  }
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard) {
    throw new Error("Aucun classeur actif. Ouvrez le dashboard puis relancez installModuleFormSubmitTriggersFromDashboard.");
  }

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    throw new Error("CONFIG_SOURCES vide ou introuvable.");
  }

  const sourceRows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  const existing = {};
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() !== "onModuleFormSubmit_") return;
    const id = trigger.getTriggerSourceId();
    if (id) existing[id] = true;
  });

  let created = 0;
  sourceRows.forEach((row) => {
    const workbookUrl = String(row[3] || "").trim();
    if (!workbookUrl || !isTruthy_(row[8])) return;
    const workbookId = extractIdFromDriveUrl_(workbookUrl);
    if (!workbookId || existing[workbookId]) return;

    ScriptApp.newTrigger("onModuleFormSubmit_")
      .forSpreadsheet(workbookId)
      .onFormSubmit()
      .create();
    existing[workbookId] = true;
    created += 1;
  });

  return { createdTriggers: created };
}

function installLiveFormRefreshTriggerFromDashboard() {
  if (isNativeFormsMode_()) {
    return { createdTriggers: 0, existing: 0, skipped: "NATIVE_HTML" };
  }
  const dashboard = openDashboardSpreadsheet_();
  if (!dashboard) throw new Error("Dashboard introuvable pour installer le trigger live.");
  setStoredDashboardId_(dashboard.getId());

  const existing = ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === LIVE_FORM_REFRESH_HANDLER);
  if (existing.length > 0) {
    return { createdTriggers: 0, existing: existing.length };
  }

  ScriptApp.newTrigger(LIVE_FORM_REFRESH_HANDLER)
    .timeBased()
    .everyHours(1)
    .create();
  return { createdTriggers: 1, existing: 0 };
}

function repairDataFlowFromDashboard() {
  return applyInPlaceSystemUpdate();
}

function runOpsModulesAuditFromDashboard() {
  const dashboard = resolveAdminDashboard_("Audit OPS modules");
  setStoredDashboardId_(dashboard.getId());

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    throw new Error("CONFIG_SOURCES vide ou introuvable.");
  }

  const rows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  const audit = [[
    "Timestamp",
    "SourceKey",
    "Module",
    "SiteKey",
    "Status",
    "Detail",
    "ItemsActifs",
    "Ruptures",
    "RequetesOuvertes",
    "LastMovementAt",
    "WorkbookUrl",
  ]];

  let checked = 0;
  let errors = 0;

  rows.forEach((row) => {
    const sourceKey = String(row[0] || "").trim();
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];
    if (!sourceKey || !workbookUrl || !isTruthy_(enabled)) return;
    if (module !== "incendie" && module !== "pharmacie") return;

    checked += 1;
    try {
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      const required = module === "incendie"
        ? ["ITEMS", "MOVEMENTS", "STOCK_VIEW", "INVENTORY_COUNT", "PURCHASE_REQUESTS", "FIRE_FORM_MOVEMENTS_RAW", "FIRE_FORM_INVENTORY_RAW"]
        : ["ITEMS_PHARMACY", "MOVEMENTS_PHARMACY", "STOCK_VIEW_PHARMACY", "INVENTORY_COUNT_PHARMACY", "PURCHASE_REQUESTS_PHARMACY", "PHARMA_FORM_MOVEMENTS_RAW", "PHARMA_FORM_INVENTORY_RAW"];

      const missing = required.filter((name) => !workbook.getSheetByName(name));
      if (missing.length) {
        errors += 1;
        audit.push([new Date(), sourceKey, module, siteKey, "ERROR", `Onglets manquants: ${missing.join(", ")}`, "", "", "", "", workbookUrl]);
        return;
      }

      const itemsSheet = workbook.getSheetByName(module === "incendie" ? "ITEMS" : "ITEMS_PHARMACY");
      const stockSheet = workbook.getSheetByName(module === "incendie" ? "STOCK_VIEW" : "STOCK_VIEW_PHARMACY");
      const purchaseSheet = workbook.getSheetByName(module === "incendie" ? "PURCHASE_REQUESTS" : "PURCHASE_REQUESTS_PHARMACY");
      const movementSheet = workbook.getSheetByName(module === "incendie" ? "MOVEMENTS" : "MOVEMENTS_PHARMACY");

      const itemsActifs = countActiveItems_(itemsSheet);
      const ruptures = stockSheet ? countMatchesInColumn_(stockSheet, 8, "RUPTURE") : 0;
      const openRequests = purchaseSheet ? countAnyMatchesInColumn_(purchaseSheet, 9, ["A_VALIDER", "EN_COURS"]) : 0;
      const lastMovementAt = getMaxDateInColumn_(movementSheet, 2);
      const formulaCheckSheet = stockSheet ? String(stockSheet.getRange("A2").getFormula() || "") : "";
      const status = formulaCheckSheet ? "OK" : "WARN";
      const detail = formulaCheckSheet
        ? "Flux principal present (A2 formula)"
        : "Formule A2 absente dans STOCK_VIEW";
      if (!formulaCheckSheet) errors += 1;

      audit.push([new Date(), sourceKey, module, siteKey, status, detail, itemsActifs, ruptures, openRequests, lastMovementAt || "", workbookUrl]);
    } catch (error) {
      errors += 1;
      audit.push([new Date(), sourceKey, module, siteKey, "ERROR", String(error.message || error), "", "", "", "", workbookUrl]);
    }
  });

  const auditSheet = dashboard.getSheetByName("OPS_MODULE_AUDIT") || dashboard.insertSheet("OPS_MODULE_AUDIT");
  auditSheet.clearContents();
  auditSheet.clearFormats();
  auditSheet.getRange(1, 1, audit.length, audit[0].length).setValues(audit);
  auditSheet.getRange(1, 1, 1, audit[0].length).setFontWeight("bold");
  auditSheet.autoResizeColumns(1, audit[0].length);
  SpreadsheetApp.flush();

  return { checked, errors, rows: Math.max(audit.length - 1, 0) };
}

function runSystemHealthCheckFromDashboard() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard) throw new Error("Aucun classeur actif.");
  setStoredDashboardId_(dashboard.getId());

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  const formSheet = dashboard.getSheetByName("FORM_LINKS");
  if (!sourceSheet || !formSheet) throw new Error("CONFIG_SOURCES et/ou FORM_LINKS introuvable(s).");

  const sourceRows = sourceSheet.getLastRow() > 1
    ? sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues()
    : [];
  let formRows = formSheet.getLastRow() > 1
    ? formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 6).getValues()
    : [];
  const nativeMode = isNativeFormsMode_();

  const sourceMap = {};
  const audit = [["Type", "Key", "Status", "Detail"]];
  let errors = 0;

  if (nativeMode) {
    try {
      const sync = ensureNativeFormLinksFromDashboard_(dashboard);
      formRows = formSheet.getLastRow() > 1
        ? formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 6).getValues()
        : [];
      audit.push(["SYSTEM", "FORM_ENGINE", "OK", `NATIVE_HTML actif | sync links: +${sync.createdRows}/~${sync.updatedRows}/-${sync.disabledRows}`]);
    } catch (error) {
      errors += 1;
      audit.push(["SYSTEM", "FORM_ENGINE", "ERROR", `NATIVE_HTML sync impossible: ${String(error.message || error)}`]);
    }
  }

  sourceRows.forEach((row) => {
    const sourceKey = String(row[0] || "").trim();
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];
    if (!sourceKey || !workbookUrl || !isTruthy_(enabled)) return;

    sourceMap[`${module}|${siteKey}`] = workbookUrl;
    try {
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      const requiredTabs = module === "pharmacie"
        ? ["ITEMS_PHARMACY", "MOVEMENTS_PHARMACY", "STOCK_VIEW_PHARMACY", "PHARMA_FORM_INVENTORY_RAW"]
        : ["ITEMS", "MOVEMENTS", "STOCK_VIEW", "FIRE_FORM_INVENTORY_RAW"];
      const missing = requiredTabs.filter((name) => !workbook.getSheetByName(name));
      if (missing.length) {
        errors += 1;
        audit.push(["SOURCE", sourceKey, "ERROR", `Onglets manquants: ${missing.join(", ")}`]);
      } else {
        audit.push(["SOURCE", sourceKey, "OK", workbook.getName()]);
      }
    } catch (error) {
      errors += 1;
      audit.push(["SOURCE", sourceKey, "ERROR", String(error.message || error)]);
    }
  });

  const triggerSourceIds = {};
  let liveRefreshTriggerExists = false;
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === "onModuleFormSubmit_") {
      const id = trigger.getTriggerSourceId();
      if (id) triggerSourceIds[id] = true;
    }
    if (trigger.getHandlerFunction() === LIVE_FORM_REFRESH_HANDLER) {
      liveRefreshTriggerExists = true;
    }
  });

  if (nativeMode) {
    audit.push(["SYSTEM", "LIVE_FORM_REFRESH", "OK", "Mode natif: refresh live Google Form non requis"]);
  } else {
    audit.push(["SYSTEM", "LIVE_FORM_REFRESH", liveRefreshTriggerExists ? "OK" : "ERROR", liveRefreshTriggerExists ? "Trigger horaire actif" : "Trigger horaire absent"]);
    if (!liveRefreshTriggerExists) errors += 1;
  }

  formRows.forEach((row) => {
    const formKey = String(row[0] || "").trim();
    const formLabel = String(row[1] || "").trim();
    const formUrl = String(row[2] || "").trim();
    const module = String(row[3] || "").trim().toLowerCase();
    const siteKey = String(row[4] || "").trim();
    const isActive = row[5];
    if (!formKey || !formUrl || !isTruthy_(isActive)) return;

    const workbookUrl = sourceMap[`${module}|${siteKey}`];
    if (!workbookUrl) {
      errors += 1;
      audit.push(["FORM", formKey, "ERROR", `Source manquante pour ${module}/${siteKey}`]);
      return;
    }

    if (nativeMode) {
      try {
        const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
        const rawSheetName = rawSheetNameForFormKey_(formKey);
        if (!rawSheetName) {
          errors += 1;
          audit.push(["FORM", formKey, "ERROR", "FormKey non mappe vers un onglet RAW"]);
          return;
        }
        if (!isNativeFormReference_(formUrl)) {
          errors += 1;
          audit.push(["FORM", formKey, "ERROR", "Reference native invalide dans FORM_LINKS"]);
          return;
        }
        if (!workbook.getSheetByName(rawSheetName)) {
          errors += 1;
          audit.push(["FORM", formKey, "ERROR", `Onglet RAW introuvable: ${rawSheetName}`]);
          return;
        }
        audit.push(["FORM", formKey, "OK", `Natif HTML | RAW:${rawSheetName}`]);
      } catch (error) {
        errors += 1;
        audit.push(["FORM", formKey, "ERROR", String(error.message || error)]);
      }
      return;
    }

    try {
      const workbookId = extractIdFromDriveUrl_(workbookUrl);
      const form = openManagedForm_(dashboard, formKey, formUrl, siteKey, formLabel);
      const destinationId = String(form.getDestinationId() || "");
      const responsesState = form.isAcceptingResponses() ? "accepting" : "closed";
      const questions = form.getItems().length;
      let minQuestions = 5;
      if (formKey.indexOf("FORM_FIRE_MOVEMENT_") === 0) minQuestions = 12;
      if (formKey.indexOf("FORM_FIRE_INVENTORY_") === 0) minQuestions = 7;
      if (formKey.indexOf("FORM_FIRE_ITEM_CREATE_") === 0) minQuestions = 12;
      if (formKey.indexOf("FORM_PHARMA_MOVEMENT_") === 0) minQuestions = 14;
      if (formKey.indexOf("FORM_PHARMA_INVENTORY_") === 0) minQuestions = 8;
      if (formKey.indexOf("FORM_PHARMA_ITEM_CREATE_") === 0) minQuestions = 14;
      const hasTrigger = !!triggerSourceIds[workbookId];

      if (destinationId !== workbookId) {
        errors += 1;
        audit.push(["FORM", formKey, "ERROR", `Destination mismatch ${destinationId} != ${workbookId}`]);
      } else if (questions < minQuestions) {
        errors += 1;
        audit.push(["FORM", formKey, "ERROR", `Questionnaire incomplet (${questions})`]);
      } else if (!hasTrigger) {
        errors += 1;
        audit.push(["FORM", formKey, "ERROR", `Trigger onModuleFormSubmit_ absent pour ${workbookId}`]);
      } else {
        audit.push(["FORM", formKey, "OK", `questions:${questions} | ${responsesState}`]);
      }
    } catch (error) {
      errors += 1;
      audit.push(["FORM", formKey, "ERROR", String(error.message || error)]);
    }
  });

  const consolidatedTabs = ["STOCK_CONSOLIDE", "ALERTS_CONSOLIDE", "LOTS_CONSOLIDE", "PURCHASE_CONSOLIDE"];
  consolidatedTabs.forEach((tab) => {
    const cell = dashboard.getSheetByName(tab).getRange("A2");
    const value = String(cell.getDisplayValue() || "");
    const formula = String(cell.getFormula() || "");
    if (!formula) {
      errors += 1;
      audit.push(["DASHBOARD", tab, "ERROR", "Formule absente en A2"]);
      return;
    }
    if (value.indexOf("#") === 0) {
      errors += 1;
      audit.push(["DASHBOARD", tab, "ERROR", `Valeur A2: ${value}`]);
      return;
    }
    audit.push(["DASHBOARD", tab, "OK", `A2:${value || "(vide)"}`]);
  });

  const userDashboardId = getStoredUserDashboardId_();
  if (!userDashboardId) {
    errors += 1;
    audit.push(["USER_DASHBOARD", "GLOBAL", "ERROR", "Dashboard user non configure"]);
  } else {
    try {
      const userDashboard = getCachedSpreadsheetById_(userDashboardId);
      const requiredTabs = ["SYSTEM_PILOTAGE", "FORM_LINKS", "KPI_OVERVIEW", "STOCK_CONSOLIDE", "ALERTS_CONSOLIDE", "LOTS_CONSOLIDE", "PURCHASE_CONSOLIDE"];
      const missing = requiredTabs.filter((name) => !userDashboard.getSheetByName(name));
      if (missing.length) {
        errors += 1;
        audit.push(["USER_DASHBOARD", "GLOBAL", "ERROR", `Onglets manquants: ${missing.join(", ")}`]);
      } else {
        const cell = userDashboard.getSheetByName("STOCK_CONSOLIDE").getRange("A2");
        const value = String(cell.getDisplayValue() || "");
        const formula = String(cell.getFormula() || "");
        if (!formula || value.indexOf("#") === 0) {
          errors += 1;
          audit.push(["USER_DASHBOARD", "GLOBAL", "ERROR", `Sync invalide en STOCK_CONSOLIDE!A2 (${value || "vide"})`]);
        } else {
          audit.push(["USER_DASHBOARD", "GLOBAL", "OK", userDashboard.getUrl()]);
        }
      }
    } catch (error) {
      errors += 1;
      audit.push(["USER_DASHBOARD", "GLOBAL", "ERROR", String(error.message || error)]);
    }
  }

  const controllerDashboardId = getStoredControllerDashboardId_();
  if (!controllerDashboardId) {
    errors += 1;
    audit.push(["CONTROLLER_DASHBOARD", "GLOBAL", "ERROR", "Dashboard controleur non configure"]);
  } else {
    try {
      const controllerDashboard = getCachedSpreadsheetById_(controllerDashboardId);
      const requiredTabs = ["SYSTEM_PILOTAGE", "FORM_LINKS", "KPI_OVERVIEW", "INVENTORY_CONTROL", "STOCK_CONSOLIDE", "ALERTS_CONSOLIDE", "LOTS_CONSOLIDE", "PURCHASE_CONSOLIDE"];
      const missing = requiredTabs.filter((name) => !controllerDashboard.getSheetByName(name));
      if (missing.length) {
        errors += 1;
        audit.push(["CONTROLLER_DASHBOARD", "GLOBAL", "ERROR", `Onglets manquants: ${missing.join(", ")}`]);
      } else {
        const cell = controllerDashboard.getSheetByName("INVENTORY_CONTROL").getRange("A2");
        const value = String(cell.getDisplayValue() || "");
        const formula = String(cell.getFormula() || "");
        if (!formula || value.indexOf("#") === 0) {
          errors += 1;
          audit.push(["CONTROLLER_DASHBOARD", "GLOBAL", "ERROR", `Sync invalide en INVENTORY_CONTROL!A2 (${value || "vide"})`]);
        } else {
          audit.push(["CONTROLLER_DASHBOARD", "GLOBAL", "OK", controllerDashboard.getUrl()]);
        }
      }
    } catch (error) {
      errors += 1;
      audit.push(["CONTROLLER_DASHBOARD", "GLOBAL", "ERROR", String(error.message || error)]);
    }
  }

  const healthSheet = dashboard.getSheetByName("SYSTEM_HEALTH_AUDIT") || dashboard.insertSheet("SYSTEM_HEALTH_AUDIT");
  healthSheet.clearContents();
  healthSheet.clearFormats();
  healthSheet.getRange(1, 1, audit.length, audit[0].length).setValues(audit);
  healthSheet.getRange(1, 1, 1, audit[0].length).setFontWeight("bold");
  healthSheet.autoResizeColumns(1, audit[0].length);
  SpreadsheetApp.flush();

  return { checks: audit.length - 1, errors };
}

function runDeletionFlowTestFromDashboard() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard) throw new Error("Aucun classeur actif.");
  setStoredDashboardId_(dashboard.getId());

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    throw new Error("CONFIG_SOURCES vide ou introuvable.");
  }

  const rows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  const audit = [["SourceKey", "Module", "SiteKey", "RawSheet", "ItemID", "IsActive", "InStockView", "Status", "Detail"]];
  let checked = 0;
  let errors = 0;

  rows.forEach((row) => {
    const sourceKey = String(row[0] || "").trim();
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];
    if (!sourceKey || !module || !workbookUrl || !isTruthy_(enabled)) return;

    try {
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      const context = module === "pharmacie"
        ? { rawSheets: ["PHARMA_FORM_MOVEMENTS_RAW", "PHARMA_FORM_INVENTORY_RAW"], itemsSheet: "ITEMS_PHARMACY", stockSheet: "STOCK_VIEW_PHARMACY" }
        : { rawSheets: ["FIRE_FORM_MOVEMENTS_RAW", "FIRE_FORM_INVENTORY_RAW"], itemsSheet: "ITEMS", stockSheet: "STOCK_VIEW" };

      const itemsMap = buildItemActiveMap_(workbook.getSheetByName(context.itemsSheet));
      const stockMap = buildStockPresenceMap_(workbook.getSheetByName(context.stockSheet));
      let foundDeleteRow = false;

      context.rawSheets.forEach((rawName) => {
        const rawSheet = workbook.getSheetByName(rawName);
        if (!rawSheet || rawSheet.getLastRow() < 2) return;

        const headers = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0]
          .map((value) => normalizeTextKey_(value));
        const rawRows = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, rawSheet.getLastColumn()).getValues();

        rawRows.forEach((rawRow) => {
          const targetIds = extractDeleteTargetIdsFromRawRow_(rawRow, headers);
          if (!targetIds.length) return;
          foundDeleteRow = true;

          targetIds.forEach((itemId) => {
            const key = normalizeTextKey_(itemId);
            const rawActive = Object.prototype.hasOwnProperty.call(itemsMap, key) ? itemsMap[key] : null;
            const isActive = rawActive === null ? "" : isTruthy_(rawActive);
            const inStock = !!stockMap[key];
            let status = "OK";
            let detail = "Suppression coherente";

            if (rawActive === null) {
              status = "WARN";
              detail = "Article introuvable dans ITEMS";
            } else if (isActive) {
              status = "ERROR";
              detail = "Article toujours actif";
            }
            if (inStock) {
              status = "ERROR";
              detail = `${detail} | encore present dans STOCK_VIEW`;
            }
            if (status === "ERROR") errors += 1;
            checked += 1;
            audit.push([sourceKey, module, siteKey, rawName, itemId, String(isActive), String(inStock), status, detail]);
          });
        });
      });

      if (!foundDeleteRow) {
        audit.push([sourceKey, module, siteKey, "(all)", "", "", "", "INFO", "Aucune suppression detectee dans les reponses formulaires."]);
      }
    } catch (error) {
      errors += 1;
      audit.push([sourceKey, module, siteKey, "(open)", "", "", "", "ERROR", String(error.message || error)]);
    }
  });

  const sheet = dashboard.getSheetByName("DELETE_FLOW_AUDIT") || dashboard.insertSheet("DELETE_FLOW_AUDIT");
  sheet.clearContents();
  sheet.clearFormats();
  sheet.getRange(1, 1, audit.length, audit[0].length).setValues(audit);
  sheet.getRange(1, 1, 1, audit[0].length).setFontWeight("bold");
  sheet.autoResizeColumns(1, audit[0].length);
  SpreadsheetApp.flush();

  return { checked, errors, rows: audit.length - 1 };
}

function populateDashboard_(spreadsheet, sourceRows, formRows, navRows, logRows) {
  writeRows_(spreadsheet.getSheetByName("CONFIG_SOURCES"), 2, sourceRows);
  writeRows_(spreadsheet.getSheetByName("FORM_LINKS"), 2, formRows);
  writeRows_(spreadsheet.getSheetByName("NAVIGATION"), 2, navRows);
  writeRows_(spreadsheet.getSheetByName("KPI_OVERVIEW"), 2, KPI_ROWS);
  writeRows_(spreadsheet.getSheetByName("DEPLOYMENT_LOG"), 2, logRows);

  if (sourceRows.length > 0) {
    const first = 2;
    const last = sourceRows.length + 1;
    spreadsheet.getSheetByName("STOCK_CONSOLIDE").getRange("A2").setFormula(buildConsolidationFormula_(first, last, "E"));
    spreadsheet.getSheetByName("ALERTS_CONSOLIDE").getRange("A2").setFormula(buildConsolidationFormula_(first, last, "F"));
    spreadsheet.getSheetByName("LOTS_CONSOLIDE").getRange("A2").setFormula(buildConsolidationFormula_(first, last, "G"));
    applyPurchaseConsolidationFormulas_(spreadsheet, first, last);
  }
  applyPurchaseValidationFormula_(spreadsheet);

  const kpi = spreadsheet.getSheetByName("KPI_OVERVIEW");
  kpi.getRange("C2").setFormula('=COUNTA(UNIQUE(FILTER(STOCK_CONSOLIDE!A2:A;STOCK_CONSOLIDE!A2:A<>"")))');
  kpi.getRange("C3").setFormula('=COUNTIF(STOCK_CONSOLIDE!H2:H;"SOUS_SEUIL")+COUNTIF(STOCK_CONSOLIDE!H2:H;"RUPTURE")');
  kpi.getRange("C4").setFormula('=COUNTIF(STOCK_CONSOLIDE!H2:H;"RUPTURE")');
  kpi.getRange("C5").setFormula('=COUNTIFS(PURCHASE_CONSOLIDE!L2:L;"A_VALIDER")+COUNTIFS(PURCHASE_CONSOLIDE!L2:L;"EN_COURS")');
  kpi.getRange("C6").setFormula('=COUNTIFS(LOTS_CONSOLIDE!E2:E;"<="&TODAY()+90;LOTS_CONSOLIDE!E2:E;">="&TODAY();LOTS_CONSOLIDE!F2:F;">0")');
  kpi.getRange("C7").setFormula('=COUNTIFS(LOTS_CONSOLIDE!E2:E;"<="&TODAY()+30;LOTS_CONSOLIDE!E2:E;">="&TODAY();LOTS_CONSOLIDE!F2:F;">0")');
  kpi.getRange("C8").setFormula('=COUNTIFS(LOTS_CONSOLIDE!E2:E;"<"&TODAY();LOTS_CONSOLIDE!F2:F;">0")');
  kpi.getRange("C9").setFormula('=IFERROR(SUM(FILTER(STOCK_CONSOLIDE!F2:F;STOCK_CONSOLIDE!A2:A<>""));0)');
  kpi.getRange("C10").setFormula('=IFERROR(SUM(FILTER(STOCK_CONSOLIDE!F2:F;STOCK_CONSOLIDE!K2:K="incendie";STOCK_CONSOLIDE!A2:A<>""));0)');
  kpi.getRange("C11").setFormula('=IFERROR(SUM(FILTER(STOCK_CONSOLIDE!F2:F;STOCK_CONSOLIDE!K2:K="pharmacie";STOCK_CONSOLIDE!A2:A<>""));0)');
  kpi.getRange("C12").setFormula('=IFERROR(ROUND(COUNTIF(STOCK_CONSOLIDE!H2:H;"OK")/MAX(COUNTA(FILTER(STOCK_CONSOLIDE!A2:A;STOCK_CONSOLIDE!A2:A<>""));1)*100;1);0)');
  buildKpiOverviewAnalytics_(spreadsheet);

  const actions = spreadsheet.getSheetByName("ACTIONS_RAPIDES");
  actions.getRange("A2").setFormula('=ARRAYFORMULA(IFERROR(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);""))');
  if (isNativeFormsMode_()) {
    actions.getRange("B2").setFormula('=ARRAYFORMULA(IFERROR(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE)&" -> menu Operations > Formulaires natifs (HTML)";""))');
  } else {
    actions.getRange("B2").setFormula('=ARRAYFORMULA(IFERROR(HYPERLINK(FILTER(FORM_LINKS!C2:C;FORM_LINKS!F2:F=TRUE);"Ouvrir formulaire");""))');
  }

  buildSystemPilotage_(spreadsheet);
}

function buildKpiOverviewAnalytics_(spreadsheet) {
  const sheet = spreadsheet ? spreadsheet.getSheetByName("KPI_OVERVIEW") : null;
  if (!sheet) return;
  const isLayoutCurrent = isLayoutMarkerCurrent_(spreadsheet, "KPI_OVERVIEW", LAYOUT_VERSIONS.KPI_OVERVIEW)
    && isKpiOverviewLayoutReady_(sheet);
  if (isLayoutCurrent) {
    refreshKpiOverviewCharts_(sheet, false);
    return;
  }

  const totalRows = 92;
  const totalCols = 22;
  if (sheet.getMaxRows() < totalRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < totalCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), totalCols - sheet.getMaxColumns());
  }

  // Reset uniquement la zone analytics pour eviter les fusions residuelles.
  sheet.getRange(1, 7, totalRows, totalCols - 6).breakApart();
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);

  // Mise en forme bloc KPI principal.
  sheet.getRange("A1:E1")
    .setBackground("#0d3559")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.getRange("A2:E20")
    .setBackground("#f4f9ff")
    .setFontColor("#103552");
  sheet.getRange("A2:A20").setFontWeight("bold");
  sheet.getRange("C2:C20").setNumberFormat("#,##0.00");
  sheet.getRange("C12").setNumberFormat("0.0");
  sheet.getRange("A1:E20").setBorder(true, true, true, true, false, false, "#7aa5c7", SpreadsheetApp.BorderStyle.SOLID);

  // Bloc analyse principale.
  sheet.getRange("G1:N1").merge()
    .setValue("Analyse stock (temps reel)")
    .setBackground("#0d3559")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");
  sheet.getRange("O1:V1").merge()
    .setValue("Risques et flux")
    .setBackground("#0d3559")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  sheet.getRange("G3:I3").setValues([["Module", "Articles", "Stock actuel"]]).setBackground("#14456f").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("G4").setFormula('=ARRAYFORMULA(IFERROR(QUERY(STOCK_CONSOLIDE!A2:K;"select K,count(A),sum(F) where A is not null group by K label count(A) \'\', sum(F) \'\'";0);""))');

  sheet.getRange("K3:L3").setValues([["Statut", "Articles"]]).setBackground("#14456f").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("K4").setFormula('=ARRAYFORMULA(IFERROR(QUERY(STOCK_CONSOLIDE!A2:H;"select H,count(A) where A is not null group by H label count(A) \'\'";0);""))');

  sheet.getRange("O3:P3").setValues([["Priorite", "Demandes"]]).setBackground("#14456f").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("O4").setFormula('=ARRAYFORMULA(IFERROR(QUERY(PURCHASE_CONSOLIDE!A2:L;"select H,count(A) where A is not null and (L=\'A_VALIDER\' or L=\'EN_COURS\') group by H label count(A) \'\'";0);""))');

  sheet.getRange("R3:S3").setValues([["Module", "Demandes ouvertes"]]).setBackground("#14456f").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("R4").setFormula('=ARRAYFORMULA(IFERROR(QUERY(PURCHASE_CONSOLIDE!A2:L;"select B,count(A) where A is not null and (L=\'A_VALIDER\' or L=\'EN_COURS\') group by B label count(A) \'\'";0);""))');

  // Top suivi stock actuel.
  sheet.getRange("G11:P11")
    .setValues([["ItemID", "Article", "Module", "Site", "Stock", "Seuil", "Ecart", "Statut", "Dernier mouvement", "Couverture"]])
    .setBackground("#14456f")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.getRange("G12").setFormula(buildStockTrackingArrayFormula_(10));

  // Tendance alertes.
  sheet.getRange("O11:P11").setValues([["Date", "Alertes"]]).setBackground("#14456f").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("O12").setFormula('=ARRAYFORMULA(HSTACK(TODAY()-SEQUENCE(7;1;6;-1);IFERROR(COUNTIFS(ALERTS_CONSOLIDE!I2:I;">="&TODAY()-SEQUENCE(7;1;6;-1);ALERTS_CONSOLIDE!I2:I;"<"&TODAY()-SEQUENCE(7;1;5;-1));0)))');
  sheet.getRange("O12:O18").setNumberFormat("dd/mm");

  // Analyse sites + demandes + lots.
  sheet.getRange("G22:I22").setValues([["Site", "Stock actuel", "Articles"]]).setBackground("#14456f").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("G23").setFormula('=ARRAYFORMULA(IFERROR(QUERY(STOCK_CONSOLIDE!A2:L;"select L,sum(F),count(A) where A is not null group by L label sum(F) \'\', count(A) \'\'";0);""))');

  sheet.getRange("K22:L22").setValues([["Statut demande", "Demandes"]]).setBackground("#14456f").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("K23").setFormula('=ARRAYFORMULA(IFERROR(QUERY(PURCHASE_CONSOLIDE!A2:L;"select L,count(A) where A is not null group by L label count(A) \'\'";0);""))');

  sheet.getRange("N22:O22").setValues([["Etat lot", "Nombre lots"]]).setBackground("#14456f").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("N23").setFormula('=ARRAYFORMULA(IFERROR(QUERY(LOTS_CONSOLIDE!A2:H;"select H,count(A) where A is not null group by H label count(A) \'\'";0);""))');

  sheet.getRange("R22:S22").setValues([["Fenetre", "Volume"]]).setBackground("#14456f").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("R23").setValue("Peremption 30j");
  sheet.getRange("R24").setValue("Peremption 90j");
  sheet.getRange("R25").setValue("Deja expires");
  sheet.getRange("S23").setFormula('=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("EXPIRY_30D";KPI_OVERVIEW!A:A;0));0)');
  sheet.getRange("S24").setFormula('=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("EXPIRY_90D";KPI_OVERVIEW!A:A;0));0)');
  sheet.getRange("S25").setFormula('=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("EXPIRY_PAST";KPI_OVERVIEW!A:A;0));0)');

  sheet.getRange("G3:V32").setFontSize(9);
  sheet.getRange("G4:V32").setBackground("#f7fbff");
  sheet.getRange("P12:P21").setNumberFormat("0.00");
  sheet.getRange("P12:P21").setHorizontalAlignment("center");
  sheet.getRange("S23:S25").setNumberFormat("#,##0");

  refreshKpiOverviewCharts_(sheet, true);
  setLayoutMarkerVersion_(spreadsheet, "KPI_OVERVIEW", LAYOUT_VERSIONS.KPI_OVERVIEW);
}

function refreshKpiOverviewCharts_(sheet, forceRebuild) {
  if (!sheet) return;
  const charts = sheet.getCharts();
  if (!forceRebuild && charts && charts.length >= 6) return;
  charts.forEach((chart) => sheet.removeChart(chart));

  const moduleChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("G3:I8"))
    .setPosition(36, 7, 0, 0)
    .setOption("title", "Stock actuel par module")
    .setOption("legend", { position: "none" })
    .setOption("backgroundColor", "#f7fbff")
    .setNumHeaders(1)
    .build();
  sheet.insertChart(moduleChart);

  const statusChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange("K3:L8"))
    .setPosition(36, 12, 0, 0)
    .setOption("title", "Repartition des statuts stock")
    .setOption("backgroundColor", "#f7fbff")
    .setNumHeaders(1)
    .build();
  sheet.insertChart(statusChart);

  const priorityChart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange("O3:P8"))
    .setPosition(36, 17, 0, 0)
    .setOption("title", "Demandes ouvertes par priorite")
    .setOption("legend", { position: "none" })
    .setOption("backgroundColor", "#f7fbff")
    .setNumHeaders(1)
    .build();
  sheet.insertChart(priorityChart);

  const trendChart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange("O11:P18"))
    .setPosition(54, 12, 0, 0)
    .setOption("title", "Tendance alertes (7 jours)")
    .setOption("legend", { position: "none" })
    .setOption("backgroundColor", "#f7fbff")
    .setNumHeaders(1)
    .build();
  sheet.insertChart(trendChart);

  const siteChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("G22:I30"))
    .setPosition(54, 7, 0, 0)
    .setOption("title", "Stock actuel par site")
    .setOption("legend", { position: "none" })
    .setOption("backgroundColor", "#f7fbff")
    .setNumHeaders(1)
    .build();
  sheet.insertChart(siteChart);

  const purchaseStatusChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange("K22:L30"))
    .setPosition(54, 17, 0, 0)
    .setOption("title", "Cycle des demandes achat")
    .setOption("backgroundColor", "#f7fbff")
    .setNumHeaders(1)
    .build();
  sheet.insertChart(purchaseStatusChart);
}

function buildSystemPilotage_(spreadsheet, options) {
  const sheet = spreadsheet.getSheetByName("SYSTEM_PILOTAGE");
  if (!sheet) return;
  const forceRebuild = !!(options && options.forceRebuild);
  const isLayoutCurrent = !forceRebuild
    && isLayoutMarkerCurrent_(spreadsheet, "SYSTEM_PILOTAGE_GLOBAL", LAYOUT_VERSIONS.GLOBAL_PILOTAGE)
    && isSystemPilotageLayoutReady_(sheet, "GLOBAL");
  if (isLayoutCurrent) {
    setupPilotageToggleControls_(sheet, "GLOBAL");
    setupPilotageNativeActionControls_(sheet, "GLOBAL");
    applyPilotageSectionVisibility_(sheet, "GLOBAL");
    applyPilotageConditionalStyles_(sheet);
    return;
  }

  const totalRows = 66;
  const totalCols = 20;
  if (sheet.getMaxRows() < totalRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < totalCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), totalCols - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clearContents();
  sheet.clearFormats();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(2);
  sheet.setRowHeights(1, totalRows, 24);

  const widths = [180, 130, 130, 26, 122, 122, 122, 122, 122, 122, 122, 122, 122, 122, 122, 122, 122, 122, 122, 122];
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));

  sheet.getRange(1, 1, totalRows, totalCols)
    .setBackground("#071522")
    .setFontColor("#eaf2ff")
    .setFontFamily("Trebuchet MS")
    .setVerticalAlignment("middle");

  sheet.getRange("A1:C66").setBackground("#04101c");
  sheet.getRange("D1:D66").setBackground("#071522");

  sheet.getRange("A1:C1").merge()
    .setValue("CONTROL CENTER")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#0a2844");
  sheet.getRange("A2:C2").merge()
    .setValue("Dashboard operations")
    .setFontSize(9)
    .setHorizontalAlignment("left")
    .setBackground("#0b2238");

  sheet.getRange("E1:T1").merge()
    .setValue("SYSTEM PILOTAGE | Gestion de stock incendie et pharmacie")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#0d3559");
  sheet.getRange("E2:T2").merge()
    .setValue("Vue centrale temps reel: KPI, formulaires, alertes, achats, lots et modules")
    .setFontSize(10)
    .setHorizontalAlignment("left")
    .setBackground("#11314f");
  sheet.getRange("E3:K3").merge()
    .setFormula('="Derniere mise a jour: "&TEXT(NOW();"dd/mm/yyyy hh:mm")')
    .setFontSize(9)
    .setHorizontalAlignment("left")
    .setBackground("#0f2f4d");

  setupPilotageToggleControls_(sheet, "GLOBAL");

  ensureAuditSheetsExist_(spreadsheet);

  stylePanel_(sheet, 4, 22, 1, 3, "Navigation rapide", "#0a1f33", "#11466f");
  stylePanel_(sheet, 24, 39, 1, 3, "Formulaires", "#0a1f33", "#11466f");
  stylePanel_(sheet, 41, 66, 1, 3, "Etat systeme", "#0a1f33", "#11466f");
  sheet.getRange("A25:C25").merge()
    .setValue("Natif HTML: Operations > Formulaires natifs > Ouvrir panneau boutons")
    .setFontSize(8)
    .setHorizontalAlignment("left")
    .setBackground("#0c2f4f");

  stylePanel_(sheet, 4, 13, 5, 20, "Indicateurs globaux", "#0a2236", "#11466f");
  stylePanel_(sheet, 15, 32, 5, 12, "Stock critique prioritaire", "#0b243b", "#11466f");
  stylePanel_(sheet, 15, 32, 13, 20, "Demandes achat en cours", "#0b243b", "#11466f");
  stylePanel_(sheet, 33, 50, 5, 12, "Alertes recentes", "#0b243b", "#11466f");
  stylePanel_(sheet, 33, 50, 13, 20, "Lots pharmacie a suivre", "#0b243b", "#11466f");
  stylePanel_(sheet, 51, 66, 5, 14, "Suivi stock actuel", "#0b243b", "#11466f");
  stylePanel_(sheet, 51, 66, 15, 20, "Actions et tendances", "#0b243b", "#11466f");

  const gidKpi = spreadsheet.getSheetByName("KPI_OVERVIEW").getSheetId();
  const gidStock = spreadsheet.getSheetByName("STOCK_CONSOLIDE").getSheetId();
  const gidAlerts = spreadsheet.getSheetByName("ALERTS_CONSOLIDE").getSheetId();
  const gidPurchase = spreadsheet.getSheetByName("PURCHASE_CONSOLIDE").getSheetId();
  const gidLots = spreadsheet.getSheetByName("LOTS_CONSOLIDE").getSheetId();
  const gidForms = spreadsheet.getSheetByName("FORM_LINKS").getSheetId();
  const gidHealth = spreadsheet.getSheetByName("SYSTEM_HEALTH_AUDIT").getSheetId();
  const gidImport = spreadsheet.getSheetByName("IMPORTRANGE_AUTH").getSheetId();
  const gidAccess = spreadsheet.getSheetByName("ACCESS_CONTROL").getSheetId();

  const navFormulas = [
    `=HYPERLINK("#gid=${gidKpi}";"Ouvrir KPI global")`,
    `=HYPERLINK("#gid=${gidStock}";"Ouvrir stock consolide")`,
    `=HYPERLINK("#gid=${gidAlerts}";"Ouvrir alertes consolidees")`,
    `=HYPERLINK("#gid=${gidPurchase}";"Ouvrir file achats")`,
    `=HYPERLINK("#gid=${gidLots}";"Ouvrir lots pharmacie")`,
    `=HYPERLINK("#gid=${gidForms}";"Ouvrir registre formulaires")`,
    `=HYPERLINK("#gid=${gidHealth}";"Ouvrir audit sante systeme")`,
    `=HYPERLINK("#gid=${gidImport}";"Ouvrir IMPORTRANGE auth")`,
    `=HYPERLINK("#gid=${gidAccess}";"Ouvrir matrice acces")`,
  ];
  const navRows = [6, 8, 10, 12, 14, 16, 18, 20, 22];
  navRows.forEach((row, index) => {
    drawPilotageButton_(sheet, row, 1, 3, navFormulas[index], "#1f5f96");
  });

  const formButtons = isNativeFormsMode_()
    ? [
      '=IFERROR(INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);1)&" (menu Operations > Formulaires natifs)";"Formulaire #1 indisponible")',
      '=IFERROR(INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);2)&" (menu Operations > Formulaires natifs)";"Formulaire #2 indisponible")',
      '=IFERROR(INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);3)&" (menu Operations > Formulaires natifs)";"Formulaire #3 indisponible")',
      '=IFERROR(INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);4)&" (menu Operations > Formulaires natifs)";"Formulaire #4 indisponible")',
      '=IFERROR(INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);5)&" (menu Operations > Formulaires natifs)";"Formulaire #5 indisponible")',
      '=IFERROR(INDEX(FILTER(FORM_LINKS!B2:B;LEFT(FORM_LINKS!A2:A;24)="FORM_PHARMA_ITEM_CREATE_");1)&" (menu Operations > Formulaires natifs)";"Form creation article pharmacie indisponible")',
    ]
    : [
      '=IFERROR(HYPERLINK(INDEX(FILTER(FORM_LINKS!C2:C;FORM_LINKS!F2:F=TRUE);1);INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);1));"Formulaire #1 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FILTER(FORM_LINKS!C2:C;FORM_LINKS!F2:F=TRUE);2);INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);2));"Formulaire #2 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FILTER(FORM_LINKS!C2:C;FORM_LINKS!F2:F=TRUE);3);INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);3));"Formulaire #3 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FILTER(FORM_LINKS!C2:C;FORM_LINKS!F2:F=TRUE);4);INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);4));"Formulaire #4 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FILTER(FORM_LINKS!C2:C;FORM_LINKS!F2:F=TRUE);5);INDEX(FILTER(FORM_LINKS!B2:B;FORM_LINKS!F2:F=TRUE);5));"Formulaire #5 indisponible")',
      '=IFERROR(HYPERLINK(INDEX(FILTER(FORM_LINKS!C2:C;LEFT(FORM_LINKS!A2:A;24)="FORM_PHARMA_ITEM_CREATE_");1);INDEX(FILTER(FORM_LINKS!B2:B;LEFT(FORM_LINKS!A2:A;24)="FORM_PHARMA_ITEM_CREATE_");1));"Form creation article pharmacie indisponible")',
    ];
  const formRows = [26, 29, 32, 35, 38, 39];
  formRows.forEach((row, index) => {
    drawPilotageButton_(sheet, row, 1, 3, formButtons[index], "#296ea8");
  });
  setupPilotageNativeActionControls_(sheet, "GLOBAL");

  drawPilotageCard_(sheet, 5, 5, 8, "Articles actifs", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("TOTAL_ITEMS";KPI_OVERVIEW!A:A;0));0)', "#2d9cdb");
  drawPilotageCard_(sheet, 5, 9, 12, "Sous-seuil", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("LOW_STOCK";KPI_OVERVIEW!A:A;0));0)', "#f9a825");
  drawPilotageCard_(sheet, 5, 13, 16, "Ruptures", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("RUPTURE";KPI_OVERVIEW!A:A;0));0)', "#ef5350");
  drawPilotageCard_(sheet, 5, 17, 20, "Achats ouverts", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("PURCHASE_OPEN";KPI_OVERVIEW!A:A;0));0)', "#4db6ac");
  drawPilotageCard_(sheet, 9, 5, 8, "Peremption 90j", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("EXPIRY_90D";KPI_OVERVIEW!A:A;0));0)', "#5c6bc0");
  drawPilotageCard_(sheet, 9, 9, 12, "Peremption 30j", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("EXPIRY_30D";KPI_OVERVIEW!A:A;0));0)', "#8e24aa");
  drawPilotageCard_(sheet, 9, 13, 16, "Lots expires", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("EXPIRY_PAST";KPI_OVERVIEW!A:A;0));0)', "#c62828");
  drawPilotageCard_(sheet, 9, 17, 20, "Stock actuel total", '=IFERROR(INDEX(KPI_OVERVIEW!C:C;MATCH("CURRENT_STOCK_TOTAL";KPI_OVERVIEW!A:A;0));0)', "#1e88e5");

  sheet.getRange("A43:C43").merge().setFormula('="Sites actifs: "&IFERROR(COUNTA(UNIQUE(FILTER(CONFIG_SOURCES!C2:C;CONFIG_SOURCES!I2:I=TRUE)));0)');
  sheet.getRange("A45:C45").merge().setFormula('="Sources actives: "&COUNTIF(CONFIG_SOURCES!I2:I;TRUE)');
  sheet.getRange("A47:C47").merge().setFormula('="Formulaires actifs: "&COUNTIF(FORM_LINKS!F2:F;TRUE)');
  sheet.getRange("A49:C49").merge().setFormula('="Alertes ouvertes: "&COUNTIF(ALERTS_CONSOLIDE!J2:J;"OPEN")');
  sheet.getRange("A51:C51").merge().setFormula('="Achats en attente: "&(COUNTIFS(PURCHASE_CONSOLIDE!L2:L;"A_VALIDER")+COUNTIFS(PURCHASE_CONSOLIDE!L2:L;"EN_COURS"))');
  sheet.getRange("A53:C53").merge().setFormula('="Dernier deploiement: "&IFERROR(TEXT(MAX(DEPLOYMENT_LOG!F2:F);"dd/mm/yyyy hh:mm");"N/A")');
  sheet.getRange("A55:C55").merge().setFormula('="Erreurs health check: "&COUNTIF(SYSTEM_HEALTH_AUDIT!C2:C;"ERROR")');
  sheet.getRange("A57:C57").merge().setFormula('="Admins configures: "&COUNTIF(ACCESS_CONTROL!B2:B;"ADMIN")');
  sheet.getRange("A58:C65").merge()
    .setValue("Consignes\n\n1) Ouvrir un formulaire depuis les boutons.\n2) Utiliser les boutons tableaux (ligne 3) pour masquer/afficher les blocs.\n3) Controler ruptures et sous-seuil.\n4) Traiter les demandes achat.\n5) Verifier lots pharmacie.")
    .setWrap(true)
    .setVerticalAlignment("top")
    .setHorizontalAlignment("left")
    .setFontSize(9)
    .setBackground("#0c2f4f");

  sheet.getRange("E17:L17").setValues([["ItemID", "Article", "Stock", "Seuil", "Ecart", "Statut", "Module", "Site"]]);
  sheet.getRange("E17:L17").setFontWeight("bold").setBackground("#14456f");
  sheet.getRange("E18").setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(STOCK_CONSOLIDE!A2:A;STOCK_CONSOLIDE!B2:B;STOCK_CONSOLIDE!F2:F;STOCK_CONSOLIDE!E2:E;STOCK_CONSOLIDE!G2:G;STOCK_CONSOLIDE!H2:H;STOCK_CONSOLIDE!K2:K;STOCK_CONSOLIDE!L2:L);STOCK_CONSOLIDE!H2:H<>"OK";STOCK_CONSOLIDE!A2:A<>"ItemID");5;FALSE);""));15;8)');

  sheet.getRange("M17:T17").setValues([["RequestID", "ItemID", "Article", "Qte", "Priorite", "Statut", "Module", "Date"]]);
  sheet.getRange("M17:T17").setFontWeight("bold").setBackground("#14456f");
  sheet.getRange("M18").setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(PURCHASE_CONSOLIDE!A2:A;PURCHASE_CONSOLIDE!D2:D;PURCHASE_CONSOLIDE!E2:E;PURCHASE_CONSOLIDE!G2:G;PURCHASE_CONSOLIDE!H2:H;PURCHASE_CONSOLIDE!L2:L;PURCHASE_CONSOLIDE!B2:B;PURCHASE_CONSOLIDE!K2:K);PURCHASE_CONSOLIDE!A2:A<>"";PURCHASE_CONSOLIDE!A2:A<>"RequestID";(PURCHASE_CONSOLIDE!L2:L="A_VALIDER")+(PURCHASE_CONSOLIDE!L2:L="EN_COURS"));8;FALSE);""));15;8)');

  sheet.getRange("E35:L35").setValues([["AlertID", "Type", "Severite", "ItemID", "Article", "Message", "Date", "Module"]]);
  sheet.getRange("E35:L35").setFontWeight("bold").setBackground("#14456f");
  sheet.getRange("E36").setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(ALERTS_CONSOLIDE!A2:A;ALERTS_CONSOLIDE!D2:D;ALERTS_CONSOLIDE!E2:E;ALERTS_CONSOLIDE!F2:F;ALERTS_CONSOLIDE!G2:G;ALERTS_CONSOLIDE!H2:H;ALERTS_CONSOLIDE!I2:I;ALERTS_CONSOLIDE!B2:B);ALERTS_CONSOLIDE!A2:A<>"";ALERTS_CONSOLIDE!A2:A<>"AlertID");7;FALSE);""));15;8)');

  sheet.getRange("M35:T35").setValues([["LotID", "ItemID", "Site", "Lot", "Peremption", "Qte", "Fournisseur", "Statut"]]);
  sheet.getRange("M35:T35").setFontWeight("bold").setBackground("#14456f");
  sheet.getRange("M36").setFormula('=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(LOTS_CONSOLIDE!A2:A;LOTS_CONSOLIDE!B2:B;LOTS_CONSOLIDE!C2:C;LOTS_CONSOLIDE!D2:D;LOTS_CONSOLIDE!E2:E;LOTS_CONSOLIDE!F2:F;LOTS_CONSOLIDE!G2:G;LOTS_CONSOLIDE!H2:H);LOTS_CONSOLIDE!A2:A<>"";LOTS_CONSOLIDE!A2:A<>"LotID");5;TRUE);""));15;8)');

  sheet.getRange("E53:N53")
    .setValues([["ItemID", "Article", "Module", "Site", "Stock", "Seuil", "Ecart", "Statut", "Dernier mouv.", "Couv."]])
    .setFontWeight("bold")
    .setBackground("#14456f")
    .setHorizontalAlignment("center");
  sheet.getRange("E54").setFormula(buildStockTrackingArrayFormula_(11));
  sheet.getRange("M54:M64").setNumberFormat("dd/mm/yyyy hh:mm");
  sheet.getRange("N54:N64").setNumberFormat("0.00");

  sheet.getRange("O53:T53").merge().setValue("Tendances").setBackground("#14456f").setFontWeight("bold");
  sheet.getRange("O54:T54").merge().setValue("Evolution alertes (7 jours)").setBackground("#0f2f4d").setFontWeight("bold");
  sheet.getRange("O55:T57").merge().setFormula('=SPARKLINE(IFERROR(COUNTIFS(ALERTS_CONSOLIDE!I2:I;">="&TODAY()-SEQUENCE(7;1;6;-1);ALERTS_CONSOLIDE!I2:I;"<"&TODAY()-SEQUENCE(7;1;5;-1));0))');
  sheet.getRange("O58:T58").merge().setValue("Evolution achats (7 jours)").setBackground("#0f2f4d").setFontWeight("bold");
  sheet.getRange("O59:T61").merge().setFormula('=SPARKLINE(IFERROR(COUNTIFS(PURCHASE_CONSOLIDE!K2:K;">="&TODAY()-SEQUENCE(7;1;6;-1);PURCHASE_CONSOLIDE!K2:K;"<"&TODAY()-SEQUENCE(7;1;5;-1));0))');
  sheet.getRange("O62:T62").merge().setFormula('="Repartition stock (OK / SOUS_SEUIL / RUPTURE): "&COUNTIF(STOCK_CONSOLIDE!H2:H;"OK")&" / "&COUNTIF(STOCK_CONSOLIDE!H2:H;"SOUS_SEUIL")&" / "&COUNTIF(STOCK_CONSOLIDE!H2:H;"RUPTURE")');
  sheet.getRange("O62:T62").setBackground("#103552").setFontWeight("bold");
  drawPilotageButton_(sheet, 63, 15, 20, `=HYPERLINK("#gid=${gidStock}";"Ouvrir stock consolide")`, "#1f5f96");
  drawPilotageButton_(sheet, 64, 15, 20, `=HYPERLINK("#gid=${gidPurchase}";"Ouvrir file achats")`, "#1f5f96");
  drawPilotageButton_(sheet, 65, 15, 20, `=HYPERLINK("#gid=${gidKpi}";"Ouvrir KPI overview")`, "#1f5f96");
  drawPilotageButton_(sheet, 66, 15, 20, `=HYPERLINK("#gid=${gidAlerts}";"Ouvrir alertes consolidees")`, "#1f5f96");

  sheet.getRange("E17:T66").setFontSize(9);
  sheet.getRange("E18:T66").setHorizontalAlignment("left");
  sheet.getRange("T18:T32").setNumberFormat("dd/mm/yyyy hh:mm");
  sheet.getRange("E17:T17").setHorizontalAlignment("center");
  sheet.getRange("M17:T17").setHorizontalAlignment("center");
  sheet.getRange("E35:T35").setHorizontalAlignment("center");
  applyPilotageSectionVisibility_(sheet, "GLOBAL");
  applyPilotageConditionalStyles_(sheet);
  setLayoutMarkerVersion_(spreadsheet, "SYSTEM_PILOTAGE_GLOBAL", LAYOUT_VERSIONS.GLOBAL_PILOTAGE);
}

function buildStockTrackingArrayFormula_(rowsLimit) {
  const maxRows = Math.max(Number(rowsLimit || 1), 1);
  return `=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(SORT(FILTER(HSTACK(STOCK_CONSOLIDE!A2:A;STOCK_CONSOLIDE!B2:B;STOCK_CONSOLIDE!K2:K;STOCK_CONSOLIDE!L2:L;STOCK_CONSOLIDE!F2:F;STOCK_CONSOLIDE!E2:E;STOCK_CONSOLIDE!G2:G;STOCK_CONSOLIDE!H2:H;IF(STOCK_CONSOLIDE!I2:I="";"";STOCK_CONSOLIDE!I2:I);IF(STOCK_CONSOLIDE!E2:E>0;ROUND(STOCK_CONSOLIDE!F2:F/STOCK_CONSOLIDE!E2:E;2);0));STOCK_CONSOLIDE!A2:A<>"";STOCK_CONSOLIDE!A2:A<>"ItemID");7;FALSE;5;TRUE);""));${maxRows};10)`;
}

function drawPilotageCard_(sheet, startRow, startCol, endCol, title, formula, accentColor) {
  const titleRange = sheet.getRange(startRow, startCol, 1, endCol - startCol + 1);
  titleRange.merge();
  titleRange
    .setValue(title)
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("left")
    .setBackground("#133a5d");

  const valueRange = sheet.getRange(startRow + 1, startCol, 2, endCol - startCol + 1);
  valueRange.merge();
  valueRange
    .setFormula(formula)
    .setFontWeight("bold")
    .setFontSize(20)
    .setHorizontalAlignment("center")
    .setBackground("#0d2b46")
    .setFontColor("#ffffff");

  const box = sheet.getRange(startRow, startCol, 3, endCol - startCol + 1);
  box.setBorder(true, true, true, true, false, false, accentColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function drawPilotageButton_(sheet, row, startCol, endCol, formula, backgroundColor) {
  const range = sheet.getRange(row, startCol, 1, endCol - startCol + 1);
  range.merge();
  range
    .setFormula(formula)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setFontSize(9)
    .setFontColor("#ffffff")
    .setBackground(backgroundColor)
    .setBorder(true, true, true, true, false, false, "#6ea9d8", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function stylePanel_(sheet, startRow, endRow, startCol, endCol, title, bodyColor, headerColor) {
  const panel = sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
  panel
    .setBackground(bodyColor)
    .setBorder(true, true, true, true, false, false, "#2a5d8f", SpreadsheetApp.BorderStyle.SOLID);

  const titleRange = sheet.getRange(startRow, startCol, 1, endCol - startCol + 1);
  titleRange.merge();
  titleRange
    .setValue(title)
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("left")
    .setBackground(headerColor)
    .setFontColor("#ffffff");
}

function resolvePilotageMode_(spreadsheet) {
  if (!spreadsheet) return "USER";
  if (spreadsheet.getSheetByName("CONFIG_SOURCES")) return "GLOBAL";
  if (spreadsheet.getSheetByName("INVENTORY_CONTROL")) return "CONTROLLER";
  return "USER";
}

function getPilotageToggleConfig_(mode) {
  const normalized = String(mode || "").toUpperCase();
  if (normalized === "GLOBAL") {
    return [
      { checkboxA1: "M3", labelA1: "N3", label: "Stock + achats", startRow: 15, numRows: 18 },
      { checkboxA1: "P3", labelA1: "Q3", label: "Alertes + lots", startRow: 33, numRows: 18 },
      { checkboxA1: "S3", labelA1: "T3", label: "Suivi + tendances", startRow: 51, numRows: 16 },
    ];
  }
  return [
    { checkboxA1: "H3", labelA1: "I3", label: "Stock + achats", startRow: 14, numRows: 17 },
    { checkboxA1: "L3", labelA1: "M3", label: "Alertes + lots + suivi", startRow: 32, numRows: 25 },
  ];
}

function getPilotageNativeActionConfig_(mode) {
  if (!isNativeFormsMode_()) return [];
  const normalized = String(mode || "").toUpperCase();
  if (normalized === "GLOBAL") {
    return [
      { checkboxA1: "D26", action: "openNativeFireMovementDialog_", label: "Mouvement incendie" },
      { checkboxA1: "D29", action: "openNativeFireReplenishmentDialog_", label: "Demande reappro incendie" },
      { checkboxA1: "D32", action: "openNativeFireItemCreateDialog_", label: "Creation article incendie" },
      { checkboxA1: "D35", action: "openNativePharmaMovementDialog_", label: "Mouvement pharmacie" },
      { checkboxA1: "D38", action: "openNativePharmaInventoryDialog_", label: "Inventaire pharmacie" },
      { checkboxA1: "D39", action: "openNativePharmaItemCreateDialog_", label: "Creation article pharmacie" },
    ];
  }
  return [
    { checkboxA1: "D24", action: "openNativeFireMovementDialog_", label: "Mouvement incendie" },
    { checkboxA1: "D27", action: "openNativeFireReplenishmentDialog_", label: "Demande reappro incendie" },
    { checkboxA1: "D30", action: "openNativeFireItemCreateDialog_", label: "Creation article incendie" },
    { checkboxA1: "D33", action: "openNativePharmaMovementDialog_", label: "Mouvement pharmacie" },
    { checkboxA1: "D36", action: "openNativePharmaInventoryDialog_", label: "Inventaire pharmacie" },
    { checkboxA1: "D38", action: "openNativePharmaItemCreateDialog_", label: "Creation article pharmacie" },
  ];
}

function ensureCheckboxCell_(sheet, a1, defaultValue) {
  const cell = sheet.getRange(a1);
  cell.insertCheckboxes();
  const current = cell.getValue();
  if (typeof current !== "boolean") {
    cell.setValue(!!defaultValue);
  }
  cell.setHorizontalAlignment("center");
  cell.setBackground("#0f2f4d");
  return cell;
}

function setupPilotageNativeActionControls_(sheet, mode) {
  if (!sheet || !isNativeFormsMode_()) return;
  const config = getPilotageNativeActionConfig_(mode);
  if (!config.length) return;

  const rows = config.map((entry) => sheet.getRange(entry.checkboxA1).getRow());
  const headerRow = Math.max(Math.min.apply(null, rows) - 1, 2);
  sheet.getRange(headerRow, 4)
    .setValue("Go")
    .setFontWeight("bold")
    .setFontColor("#dbe9ff")
    .setBackground("#0c2f4f")
    .setHorizontalAlignment("center");

  config.forEach((entry) => {
    const cell = ensureCheckboxCell_(sheet, entry.checkboxA1, false);
    cell.setValue(false);
    cell.setBackground("#103552");
    cell.setBorder(true, true, true, true, false, false, "#6ea9d8", SpreadsheetApp.BorderStyle.SOLID);
    cell.setNote(`Cochez pour ouvrir: ${entry.label}`);
  });
}

function setupPilotageToggleControls_(sheet, mode) {
  if (!sheet) return;
  const config = getPilotageToggleConfig_(mode);
  if (!config.length) return;

  const normalized = String(mode || "").toUpperCase();
  if (normalized === "GLOBAL") {
    sheet.getRange("L3:T3")
      .setBackground("#0f2f4d")
      .setFontColor("#dbe9ff")
      .setFontSize(9)
      .setHorizontalAlignment("left");
    sheet.getRange("L3").setValue("Boutons tableaux:");
  } else {
    sheet.getRange("E3:P3")
      .setBackground("#0f2f4d")
      .setFontColor("#dbe9ff")
      .setFontSize(9)
      .setHorizontalAlignment("left");
    sheet.getRange("E3:G3").merge().setValue("Boutons tableaux:");
  }

  config.forEach((entry) => {
    ensureCheckboxCell_(sheet, entry.checkboxA1, true);
    sheet.getRange(entry.labelA1)
      .setValue(entry.label)
      .setFontWeight("bold")
      .setFontColor("#dbe9ff")
      .setBackground("#0f2f4d")
      .setHorizontalAlignment("left");
  });
}

function applyPilotageSectionVisibility_(sheet, mode) {
  if (!sheet) return;
  const config = getPilotageToggleConfig_(mode);
  config.forEach((entry) => {
    const checked = sheet.getRange(entry.checkboxA1).getValue() === true;
    if (checked) {
      sheet.showRows(entry.startRow, entry.numRows);
    } else {
      sheet.hideRows(entry.startRow, entry.numRows);
    }
  });
}

function isPilotageToggleCell_(sheet, range) {
  if (!sheet || !range || sheet.getName() !== "SYSTEM_PILOTAGE") return false;
  const mode = resolvePilotageMode_(sheet.getParent());
  const a1 = range.getA1Notation();
  const config = getPilotageToggleConfig_(mode);
  return config.some((entry) => entry.checkboxA1 === a1);
}

function handlePilotageToggleEdit_(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (!isPilotageToggleCell_(sheet, e.range)) return;
  const mode = resolvePilotageMode_(sheet.getParent());
  applyPilotageSectionVisibility_(sheet, mode);
}

function resolvePilotageNativeActionEntry_(sheet, range) {
  if (!sheet || !range || sheet.getName() !== "SYSTEM_PILOTAGE") return null;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return null;
  const mode = resolvePilotageMode_(sheet.getParent());
  const a1 = range.getA1Notation();
  const config = getPilotageNativeActionConfig_(mode);
  for (let i = 0; i < config.length; i += 1) {
    if (config[i].checkboxA1 === a1) return config[i];
  }
  return null;
}

function isEditedToTrue_(e) {
  if (!e || !e.range) return false;
  if (typeof e.value !== "undefined") {
    return String(e.value || "").trim().toUpperCase() === "TRUE";
  }
  return e.range.getValue() === true;
}

function runPilotageNativeAction_(actionName) {
  if (actionName === "openNativeFireMovementDialog_") return openNativeFireMovementDialog_();
  if (actionName === "openNativeFireReplenishmentDialog_") return openNativeFireReplenishmentDialog_();
  if (actionName === "openNativeFireItemCreateDialog_") return openNativeFireItemCreateDialog_();
  if (actionName === "openNativePharmaMovementDialog_") return openNativePharmaMovementDialog_();
  if (actionName === "openNativePharmaInventoryDialog_") return openNativePharmaInventoryDialog_();
  if (actionName === "openNativePharmaItemCreateDialog_") return openNativePharmaItemCreateDialog_();
  throw new Error(`Action native inconnue: ${actionName}`);
}

function handlePilotageNativeActionEdit_(e) {
  if (!isNativeFormsMode_()) return;
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  const actionEntry = resolvePilotageNativeActionEntry_(sheet, range);
  if (!actionEntry) return;
  if (!isEditedToTrue_(e)) return;

  range.setValue(false);
  try {
    runPilotageNativeAction_(actionEntry.action);
  } catch (error) {
    const message = String(error && error.message ? error.message : error);
    const spreadsheet = sheet.getParent();
    if (spreadsheet) {
      spreadsheet.toast(`Action ${actionEntry.label}: ${message}`, "Formulaire natif", 8);
    }
    try {
      SpreadsheetApp.getUi().alert(`Action ${actionEntry.label}: ${message}`);
    } catch (uiError) {
      Logger.log(`Pilotage native action UI warning: ${String(uiError.message || uiError)}`);
    }
    throw error;
  }
}

function handlePilotageNativeActionNoAuth_(e) {
  if (!isNativeFormsMode_()) return;
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  const actionEntry = resolvePilotageNativeActionEntry_(sheet, range);
  if (!actionEntry) return;
  if (!isEditedToTrue_(e)) return;

  range.setValue(false);
  try {
    sheet.getParent().toast(
      "Utiliser un bouton script ou le menu Operations > Formulaires natifs (HTML).",
      "Action bloquee (autorisation)",
      6
    );
  } catch (toastError) {
    Logger.log(`Pilotage native no-auth toast warning: ${String(toastError.message || toastError)}`);
  }
}

function applyCurrentPilotageTableVisibility_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) return;
  const sheet = spreadsheet.getSheetByName("SYSTEM_PILOTAGE");
  if (!sheet) return;
  const mode = resolvePilotageMode_(spreadsheet);
  applyPilotageSectionVisibility_(sheet, mode);
  SpreadsheetApp.flush();
}

function applyPilotageConditionalStyles_(sheet) {
  if (!sheet) return;
  const rules = [];

  // Statuts stock critique.
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("RUPTURE")
      .setBackground("#8b1a1a")
      .setFontColor("#ffffff")
      .setRanges([sheet.getRange("J18:J32"), sheet.getRange("L54:L64")])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("SOUS_SEUIL")
      .setBackground("#8c6d1f")
      .setFontColor("#ffffff")
      .setRanges([sheet.getRange("J18:J32"), sheet.getRange("L54:L64")])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("OK")
      .setBackground("#1b5e20")
      .setFontColor("#ffffff")
      .setRanges([sheet.getRange("J18:J32"), sheet.getRange("L54:L64")])
      .build()
  );

  // Priorite achats.
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("HIGH")
      .setBackground("#7a1f32")
      .setFontColor("#ffffff")
      .setRanges([sheet.getRange("Q18:Q32")])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("MEDIUM")
      .setBackground("#5d4037")
      .setFontColor("#ffffff")
      .setRanges([sheet.getRange("Q18:Q32")])
      .build()
  );

  // Severite alertes.
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("HIGH")
      .setBackground("#b71c1c")
      .setFontColor("#ffffff")
      .setRanges([sheet.getRange("G36:G50")])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("MEDIUM")
      .setBackground("#bf6f00")
      .setFontColor("#ffffff")
      .setRanges([sheet.getRange("G36:G50")])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

function ensureAuditSheetsExist_(spreadsheet) {
  const names = [
    "FORM_CONNECTION_AUDIT",
    "FORM_UX_AUDIT",
    "DELETE_FLOW_AUDIT",
    "ITEM_SYNC_AUDIT",
    "MODULE_FORMULA_AUDIT",
    "IMPORTRANGE_AUTH",
    "SYSTEM_HEALTH_AUDIT",
    "OPS_MODULE_AUDIT",
  ];

  names.forEach((name) => {
    if (!spreadsheet.getSheetByName(name)) {
      spreadsheet.insertSheet(name);
    }
  });
}

function ensureAccessControlSheet_(spreadsheet) {
  if (!spreadsheet) return;
  let sheet = spreadsheet.getSheetByName("ACCESS_CONTROL");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("ACCESS_CONTROL");
    initSheet_(sheet, DASHBOARD_TABS.ACCESS_CONTROL);
  } else if (sheet.getLastRow() < 1) {
    initSheet_(sheet, DASHBOARD_TABS.ACCESS_CONTROL);
  } else {
    const headers = sheet.getRange(1, 1, 1, DASHBOARD_TABS.ACCESS_CONTROL.length).getValues()[0];
    const normalized = headers.map((value) => String(value || "").trim());
    const needsReset = normalized.join("|") !== DASHBOARD_TABS.ACCESS_CONTROL.join("|");
    if (needsReset) {
      sheet.clearContents();
      initSheet_(sheet, DASHBOARD_TABS.ACCESS_CONTROL);
    }
  }

  const existing = {};
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    rows.forEach((row) => {
      const email = normalizeEmail_(row[0]);
      if (email) existing[email] = true;
    });
  }

  const rowsToInsert = [];
  (DEPLOYMENT_CONFIG.adminEditors || []).forEach((email) => {
    const normalized = normalizeEmail_(email);
    if (!normalized || existing[normalized]) return;
    rowsToInsert.push([normalized, ROLE_ADMIN, "*", "*", true, true, new Date()]);
  });

  if (rowsToInsert.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, 7).setValues(rowsToInsert);
  }
  clearRuntimeAccessProfileCache_();
}

function applyUserDashboardAccessFromAccessControl_(adminDashboard, userDashboardFile) {
  if (!adminDashboard || !userDashboardFile) return;
  const accessSheet = adminDashboard.getSheetByName("ACCESS_CONTROL");
  if (!accessSheet || accessSheet.getLastRow() < 2) return;

  const currentEditors = {};
  const currentViewers = {};
  userDashboardFile.getEditors().forEach((user) => {
    const email = normalizeEmail_(user.getEmail());
    if (email) currentEditors[email] = true;
  });
  userDashboardFile.getViewers().forEach((user) => {
    const email = normalizeEmail_(user.getEmail());
    if (email) currentViewers[email] = true;
  });

  const rows = accessSheet.getRange(2, 1, accessSheet.getLastRow() - 1, 7).getValues();
  rows.forEach((row) => {
    const email = normalizeEmail_(row[0]);
    const role = String(row[1] || "").trim().toUpperCase();
    if (!email) return;

    try {
      // Les roles operationnels sont editeurs du dashboard user pour piloter les boutons de vue.
      if (role === ROLE_ADMIN || role === ROLE_CONTROLLER || role === ROLE_USER) {
        if (!currentEditors[email]) {
          userDashboardFile.addEditor(email);
          currentEditors[email] = true;
        }
      } else {
        if (!currentEditors[email] && !currentViewers[email]) {
          userDashboardFile.addViewer(email);
          currentViewers[email] = true;
        }
      }
    } catch (error) {
      Logger.log(`Access sync warning (${email}): ${String(error.message || error)}`);
    }
  });
}

function applyControllerDashboardAccessFromAccessControl_(adminDashboard, controllerDashboardFile) {
  if (!adminDashboard || !controllerDashboardFile) return;
  const accessSheet = adminDashboard.getSheetByName("ACCESS_CONTROL");
  if (!accessSheet || accessSheet.getLastRow() < 2) return;

  const currentEditors = {};
  controllerDashboardFile.getEditors().forEach((user) => {
    const email = normalizeEmail_(user.getEmail());
    if (email) currentEditors[email] = true;
  });

  const rows = accessSheet.getRange(2, 1, accessSheet.getLastRow() - 1, 7).getValues();
  rows.forEach((row) => {
    const email = normalizeEmail_(row[0]);
    const role = String(row[1] || "").trim().toUpperCase();
    if (!email) return;

    try {
      if ((role === ROLE_ADMIN || role === ROLE_CONTROLLER) && !currentEditors[email]) {
        controllerDashboardFile.addEditor(email);
        currentEditors[email] = true;
      }
    } catch (error) {
      Logger.log(`Controller access sync warning (${email}): ${String(error.message || error)}`);
    }
  });
}

function collectRoleEmailsFromAccessControl_(adminDashboard, acceptedRoles) {
  const emails = {};
  const roleSet = {};
  (acceptedRoles || []).forEach((role) => {
    roleSet[String(role || "").trim().toUpperCase()] = true;
  });

  (DEPLOYMENT_CONFIG.adminEditors || []).forEach((email) => {
    const normalized = normalizeEmail_(email);
    if (normalized) emails[normalized] = true;
  });

  const accessSheet = adminDashboard ? adminDashboard.getSheetByName("ACCESS_CONTROL") : null;
  if (!accessSheet || accessSheet.getLastRow() < 2) return emails;

  const rows = accessSheet.getRange(2, 1, accessSheet.getLastRow() - 1, 2).getValues();
  rows.forEach((row) => {
    const email = normalizeEmail_(row[0]);
    const role = String(row[1] || "").trim().toUpperCase();
    if (!email || !roleSet[role]) return;
    emails[email] = true;
  });

  return emails;
}

function enforceFilePermissions_(file, allowedEditors, allowedViewers) {
  if (!file) return;
  const ownerEmail = normalizeEmail_(file.getOwner().getEmail());
  const editorSet = allowedEditors || {};
  const viewerSet = allowedViewers || {};

  file.getEditors().forEach((user) => {
    const email = normalizeEmail_(user.getEmail());
    if (!email || email === ownerEmail) return;
    if (editorSet[email]) return;
    try {
      file.removeEditor(user);
    } catch (error) {
      Logger.log(`Permission cleanup editor warning (${email}): ${String(error.message || error)}`);
    }
  });

  file.getViewers().forEach((user) => {
    const email = normalizeEmail_(user.getEmail());
    if (!email || email === ownerEmail) return;
    if (editorSet[email] || viewerSet[email]) return;
    try {
      file.removeViewer(user);
    } catch (error) {
      Logger.log(`Permission cleanup viewer warning (${email}): ${String(error.message || error)}`);
    }
  });
}

function enforceControllerDashboardPermissions_(adminDashboard, controllerDashboardFile) {
  const allowedEditors = collectRoleEmailsFromAccessControl_(adminDashboard, [ROLE_ADMIN, ROLE_CONTROLLER]);
  enforceFilePermissions_(controllerDashboardFile, allowedEditors, {});
}

function buildConsolidationChunksFromConfig_(startRow, endRow, rangeCol) {
  const chunks = [];
  for (let r = startRow; r <= endRow; r += 1) {
    chunks.push(`IFERROR(QUERY(IMPORTRANGE(CONFIG_SOURCES!D${r};CONFIG_SOURCES!${rangeCol}${r});"select * offset 1";0);)`);
  }
  return chunks;
}

function buildConsolidationFormula_(startRow, endRow, rangeCol) {
  const chunks = buildConsolidationChunksFromConfig_(startRow, endRow, rangeCol);
  if (!chunks.length) return '=""';
  return `=IFERROR(QUERY({${chunks.join(";")}};"select * where Col1 is not null";0);"")`;
}

function buildConsolidationFormulaWithSelect_(startRow, endRow, rangeCol, selectClause) {
  const chunks = buildConsolidationChunksFromConfig_(startRow, endRow, rangeCol);
  const clause = String(selectClause || "").trim();
  if (!chunks.length || !clause) return '=""';
  return `=IFERROR(QUERY({${chunks.join(";")}};"${clause}";0);"")`;
}

function applyPurchaseConsolidationFormulas_(dashboard, startRow, endRow) {
  if (!dashboard) return;
  const sheet = dashboard.getSheetByName("PURCHASE_CONSOLIDE");
  if (!sheet) return;

  setFormulaOnSheetIfChanged_(
    sheet,
    "A2",
    buildConsolidationFormulaWithSelect_(
      startRow,
      endRow,
      "H",
      "select Col1,Col2,Col3,Col4,Col5,Col6,Col7,Col8 where Col1 is not null"
    )
  );
  setFormulaOnSheetIfChanged_(
    sheet,
    "J2",
    buildConsolidationFormulaWithSelect_(
      startRow,
      endRow,
      "H",
      "select Col10,Col11 where Col1 is not null"
    )
  );
  setFormulaOnSheetIfChanged_(
    sheet,
    "M2",
    buildConsolidationFormulaWithSelect_(
      startRow,
      endRow,
      "H",
      "select Col9 where Col1 is not null"
    )
  );

}

function protectDashboard_(spreadsheet) {
  protectSheets_(spreadsheet, [
    "STOCK_CONSOLIDE",
    "ALERTS_CONSOLIDE",
    "LOTS_CONSOLIDE",
    "PURCHASE_CONSOLIDE",
    "KPI_OVERVIEW",
    "SYSTEM_PILOTAGE",
    "ACCESS_CONTROL",
    "DEPLOYMENT_LOG",
  ]);
}

function protectSheets_(spreadsheet, names) {
  names.forEach((name) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const protection = protections.length ? protections[0] : sheet.protect();
    for (let i = 1; i < protections.length; i += 1) {
      protections[i].remove();
    }
    protection.setDescription(`Protected formulas: ${name}`);
    const editors = protection.getEditors();
    if (editors.length) protection.removeEditors(editors);
    if (protection.canDomainEdit()) protection.setDomainEdit(false);

    if (name === "SYSTEM_PILOTAGE") {
      const mode = resolvePilotageMode_(spreadsheet);
      const ranges = getPilotageToggleConfig_(mode).map((entry) => sheet.getRange(entry.checkboxA1));
      getPilotageNativeActionConfig_(mode).forEach((entry) => {
        ranges.push(sheet.getRange(entry.checkboxA1));
      });
      protection.setUnprotectedRanges(ranges);
    } else {
      protection.setUnprotectedRanges([]);
    }
  });
}

function writeRows_(sheet, startRow, rows) {
  if (!rows || rows.length === 0) return;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function moveFileToFolder_(file, folder) {
  folder.addFile(file);
  try {
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {
    // No-op.
  }
}

function applyFileAccess_(file, editors, viewers) {
  if (!file) return;
  const existingEditors = {};
  const existingViewers = {};
  file.getEditors().forEach((user) => {
    const email = normalizeEmail_(user.getEmail());
    if (email) existingEditors[email] = true;
  });
  file.getViewers().forEach((user) => {
    const email = normalizeEmail_(user.getEmail());
    if (email) existingViewers[email] = true;
  });

  editors.forEach((email) => {
    const normalized = normalizeEmail_(email);
    if (!normalized || existingEditors[normalized]) return;
    file.addEditor(normalized);
    existingEditors[normalized] = true;
  });
  viewers.forEach((email) => {
    const normalized = normalizeEmail_(email);
    if (!normalized || existingEditors[normalized] || existingViewers[normalized]) return;
    file.addViewer(normalized);
    existingViewers[normalized] = true;
  });
}
