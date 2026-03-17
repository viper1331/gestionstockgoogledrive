/*
 * Bloc 1 - Saisie / formulaires / UI
 * Menus et formulaires natifs/Google Forms + mapping des actions UI.
 */

function buildNativeFormsMenu_(ui, access) {
  const menu = ui.createMenu("Formulaires natifs (HTML)");
  menu.addSubMenu(
    ui.createMenu("Incendie")
      .addItem("Mouvement stock", "openNativeFireMovementDialog_")
      .addItem("Demande reappro", "openNativeFireReplenishmentDialog_")
      .addItem("Creation article", "openNativeFireItemCreateDialog_")
  );
  menu.addSubMenu(
    ui.createMenu("Pharmacie")
      .addItem("Mouvement stock", "openNativePharmaMovementDialog_")
      .addItem("Inventaire", "openNativePharmaInventoryDialog_")
      .addItem("Creation article", "openNativePharmaItemCreateDialog_")
  );
  menu.addSeparator().addItem("Ouvrir panneau boutons", "openNativeButtonsPanel_");
  if (!(access && (access.isAdmin || access.isController))) {
    menu.addSeparator().addItem("Info droits creation article", "showNativeFormsRightsInfo_");
  }
  return menu;
}

function showNativeFormsRightsInfo_() {
  SpreadsheetApp.getUi().alert("Creation article: reservee aux roles ADMIN / CONTROLEUR.");
}

function openNativeFireMovementDialog_() {
  openNativeFormDialog_(NATIVE_FORM_TYPES.FIRE_MOVEMENT, "incendie");
}

function openNativeFireReplenishmentDialog_() {
  openNativeFormDialog_(NATIVE_FORM_TYPES.FIRE_REPLENISHMENT, "incendie");
}

function openNativeFireItemCreateDialog_() {
  openNativeFormDialog_(NATIVE_FORM_TYPES.FIRE_ITEM_CREATE, "incendie");
}

function openNativePharmaMovementDialog_() {
  openNativeFormDialog_(NATIVE_FORM_TYPES.PHARMA_MOVEMENT, "pharmacie");
}

function openNativePharmaInventoryDialog_() {
  openNativeFormDialog_(NATIVE_FORM_TYPES.PHARMA_INVENTORY, "pharmacie");
}

function openNativePharmaItemCreateDialog_() {
  openNativeFormDialog_(NATIVE_FORM_TYPES.PHARMA_ITEM_CREATE, "pharmacie");
}

function btnMouvementIncendie() {
  openNativeFireMovementDialog_();
}

function btnReapproIncendie() {
  openNativeFireReplenishmentDialog_();
}

function btnCreationArticleIncendie() {
  openNativeFireItemCreateDialog_();
}

function btnMouvementPharmacie() {
  openNativePharmaMovementDialog_();
}

function btnInventairePharmacie() {
  openNativePharmaInventoryDialog_();
}

function btnCreationArticlePharmacie() {
  openNativePharmaItemCreateDialog_();
}

function openNativeButtonsPanel_() {
  const html = HtmlService.createTemplateFromFile("NativeButtonsPanel")
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(360)
    .setHeight(460);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openNativeFormDialog_(formType, module) {
  const meta = resolveNativeFormMeta_(formType, module);
  const context = assertNativeFormAccess_(meta.type);
  assertModuleAccess_(context.access, meta.module, "Ouverture formulaire natif");
  const template = HtmlService.createTemplateFromFile("NativeFormDialog");
  template.formType = meta.type;
  template.module = meta.module;
  const html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(620)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, `Formulaire natif | ${meta.title}`);
}

function getNativeFormBootstrap_(formType, module) {
  const meta = resolveNativeFormMeta_(formType, module);
  const context = assertNativeFormAccess_(meta.type);
  assertModuleAccess_(context.access, meta.module, "Formulaire natif");
  const sites = collectNativeSitesForModule_(context.dashboard, meta.module, context.access);
  if (!sites.length) {
    throw new Error("Aucun site autorise pour ce formulaire.");
  }
  const canUseDestructiveActions = canUseNativeDestructiveActions_(context.access);
  let nextItemId = "";
  if (meta.mode === "item_create" && sites.length) {
    const source = resolveSourceWorkbookContext_(
      context.dashboard,
      meta.module,
      sites[0].siteKey,
      context.access,
      "Generation ItemID automatique"
    );
    nextItemId = buildNextAutoItemId_(source.workbook, meta.module);
  }
  const currentUserEmail = getCurrentUserEmail_();
  return {
    formType: meta.type,
    module: meta.module,
    title: meta.title,
    description: meta.description,
    userEmail: currentUserEmail,
    sites,
    defaultSiteKey: sites.length ? sites[0].siteKey : "",
    mode: meta.mode,
    canDeleteItems: canUseDestructiveActions,
    canUseDestructiveActions,
    nextItemId,
  };
}

function getNativeFormChoices_(formType, module, siteKey) {
  const meta = resolveNativeFormMeta_(formType, module);
  const context = assertNativeFormAccess_(meta.type);
  assertModuleAccess_(context.access, meta.module, "Chargement choix formulaire natif");
  const sites = collectNativeSitesForModule_(context.dashboard, meta.module, context.access);
  if (!sites.length) {
    throw new Error("Aucun site autorise pour ce formulaire.");
  }

  const allowedSiteMap = {};
  sites.forEach((entry) => {
    const key = normalizeSiteKey_(entry.siteKey);
    if (key) allowedSiteMap[key] = entry.siteKey;
  });
  const requestedSite = String(siteKey || "").trim();
  let selectedSite = "";
  if (requestedSite) {
    assertSiteAccess_(context.access, requestedSite, "Chargement choix formulaire natif");
    const requestedKey = normalizeSiteKey_(requestedSite);
    if (!allowedSiteMap[requestedKey]) {
      throw new Error(`Acces refuse au site ${requestedSite}.`);
    }
    selectedSite = allowedSiteMap[requestedKey];
  } else {
    selectedSite = sites[0].siteKey;
  }
  if (!selectedSite) {
    throw new Error("Aucun site autorise pour ce formulaire.");
  }

  const source = resolveSourceWorkbookContext_(context.dashboard, meta.module, selectedSite, context.access, "Chargement choix formulaire natif");
  const choices = buildNativeFormChoicesForSite_(source.workbook, meta.module);
  const nextItemId = meta.mode === "item_create" ? buildNextAutoItemId_(source.workbook, meta.module) : "";
  return {
    siteKey: selectedSite,
    itemChoices: choices.itemChoices,
    replenishmentChoices: choices.replenishmentChoices,
    deleteChoices: canUseNativeDestructiveActions_(context.access) ? choices.itemChoices.slice() : [],
    nextItemId,
  };
}

function submitNativeForm_(formType, module, payload) {
  const meta = resolveNativeFormMeta_(formType, module);
  const context = assertNativeFormAccess_(meta.type);
  assertModuleAccess_(context.access, meta.module, "Soumission formulaire natif");
  const safePayload = payload || {};
  const siteKey = String(safePayload.siteKey || "").trim();
  if (!siteKey) {
    throw new Error("Le site est obligatoire.");
  }
  assertSiteAccess_(context.access, siteKey, "Soumission formulaire natif");
  if (hasNativeDeleteIntentPayload_(safePayload) && !canUseNativeDestructiveActions_(context.access)) {
    throw new Error("Action reservee aux profils ADMIN / CONTROLEUR.");
  }

  const source = resolveSourceWorkbookContext_(
    context.dashboard,
    meta.module,
    siteKey,
    context.access,
    "Soumission formulaire natif"
  );
  const actorEmail = getCurrentUserEmail_() || "";
  const submissionPayload = Object.assign({}, safePayload, { sourceWorkbook: source.workbook });
  const submission = buildNativeSubmissionPayload_(meta, submissionPayload, actorEmail, context.access);
  const rowIndex = appendNativeRow_(
    source.workbook,
    submission.rawSheetName,
    submission.requiredHeaders,
    submission.valuesByHeader,
    actorEmail
  );
  const rawSheet = source.workbook.getSheetByName(submission.rawSheetName);
  if (!rawSheet) {
    throw new Error(`Onglet RAW introuvable: ${submission.rawSheetName}`);
  }

  onModuleFormSubmit_({
    source: source.workbook,
    range: rawSheet.getRange(rowIndex, 1, 1, 1),
    namedValues: { Email: [actorEmail], "Adresse e-mail": [actorEmail] },
  });

  return {
    ok: true,
    module: meta.module,
    siteKey,
    rawSheetName: submission.rawSheetName,
    rowIndex,
    workbookName: source.workbook.getName(),
    message: submission.generatedItemId
      ? `Soumission enregistree (${meta.title}). Item cree: ${submission.generatedItemId}.`
      : `Soumission enregistree (${meta.title}) sur ${source.workbook.getName()} / ${submission.rawSheetName}.`,
  };
}

function resolveNativeFormsDashboard_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active && active.getSheetByName("CONFIG_SOURCES")) {
    setStoredDashboardId_(active.getId());
    return active;
  }

  const stored = openDashboardSpreadsheet_();
  if (stored && stored.getSheetByName("CONFIG_SOURCES")) {
    return stored;
  }

  try {
    const folder = getOrCreateFolder_(DEPLOYMENT_CONFIG.rootFolderName);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = String(file.getName() || "");
      if (name.indexOf("DASHBOARD_GLOBAL") !== 0) continue;
      try {
        const candidate = getCachedSpreadsheetById_(file.getId());
        if (candidate && candidate.getSheetByName("CONFIG_SOURCES")) {
          setStoredDashboardId_(candidate.getId());
          return candidate;
        }
      } catch (openError) {
        Logger.log(`Native dashboard lookup warning (${name}): ${String(openError.message || openError)}`);
      }
    }
  } catch (error) {
    Logger.log(`Native dashboard folder scan warning: ${String(error.message || error)}`);
  }
  throw new Error("Dashboard admin introuvable (CONFIG_SOURCES requis). Ouvrez DASHBOARD_GLOBAL puis relancez.");
}

function normalizeNativeFormType_(formType) {
  return String(formType || "").trim().toLowerCase().replace(/-/g, "_");
}

function isNativeFormsMode_() {
  return String(FORM_RUNTIME_MODE || "").trim().toUpperCase() === "NATIVE_HTML";
}

function nativeActionFunctionFromFormKey_(formKey) {
  const key = String(formKey || "").trim().toUpperCase();
  if (key.indexOf("FORM_FIRE_MOVEMENT_") === 0) return "openNativeFireMovementDialog_";
  if (isLegacyFireInventoryFormKey_(key)) return "openNativeFireReplenishmentDialog_";
  if (key.indexOf("FORM_FIRE_ITEM_CREATE_") === 0) return "openNativeFireItemCreateDialog_";
  if (key.indexOf("FORM_PHARMA_MOVEMENT_") === 0) return "openNativePharmaMovementDialog_";
  if (key.indexOf("FORM_PHARMA_INVENTORY_") === 0) return "openNativePharmaInventoryDialog_";
  if (key.indexOf("FORM_PHARMA_ITEM_CREATE_") === 0) return "openNativePharmaItemCreateDialog_";
  return "";
}

function buildNativeFormReference_(formKey, module, siteKey) {
  const action = nativeActionFunctionFromFormKey_(formKey);
  if (!action) return "";
  const modulePart = encodeURIComponent(String(module || "").trim().toLowerCase());
  const sitePart = encodeURIComponent(String(siteKey || "").trim());
  const keyPart = encodeURIComponent(String(formKey || "").trim());
  return `native-action://${action}?module=${modulePart}&site=${sitePart}&key=${keyPart}`;
}

function isNativeFormReference_(formRef) {
  const value = String(formRef || "").trim().toLowerCase();
  return value.indexOf("native-action://") === 0;
}

function buildFormLinkDefinitionsForModuleSite_(module, siteKey) {
  const normalizedModule = String(module || "").trim().toLowerCase();
  const site = String(siteKey || "").trim();
  if (!site) return [];

  let keys = [];
  if (normalizedModule === "incendie") {
    // Compatibilite legacy: la cle publique FORM_FIRE_INVENTORY_* est conservee.
    // Metier reel: flux demande de reappro incendie.
    keys = [
      `FORM_FIRE_MOVEMENT_${site}`,
      `FORM_FIRE_INVENTORY_${site}`,
      `FORM_FIRE_ITEM_CREATE_${site}`,
    ];
  } else if (normalizedModule === "pharmacie") {
    keys = [
      `FORM_PHARMA_MOVEMENT_${site}`,
      `FORM_PHARMA_INVENTORY_${site}`,
      `FORM_PHARMA_ITEM_CREATE_${site}`,
    ];
  } else {
    return [];
  }

  return keys.map((key) => {
    const label = formLabelFromKey_(key, site) || key;
    return {
      key,
      label,
      module: normalizedModule,
      siteKey: site,
      url: isNativeFormsMode_() ? buildNativeFormReference_(key, normalizedModule, site) : "",
      isActive: true,
    };
  });
}

function ensureNativeFormLinksFromDashboard_(dashboard) {
  const adminDashboard = dashboard || resolveAdminDashboard_("Synchronisation FORM_LINKS natifs");
  const sourceSheet = adminDashboard.getSheetByName("CONFIG_SOURCES");
  const formSheet = adminDashboard.getSheetByName("FORM_LINKS");
  if (!sourceSheet || !formSheet) {
    throw new Error("CONFIG_SOURCES et/ou FORM_LINKS introuvable(s).");
  }

  const sourceRows = sourceSheet.getLastRow() > 1
    ? sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues()
    : [];
  const expectedRows = [];
  const expectedKeySet = {};
  sourceRows.forEach((row) => {
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    if (!module || !siteKey || !workbookUrl || !isTruthy_(row[8])) return;
    buildFormLinkDefinitionsForModuleSite_(module, siteKey).forEach((definition) => {
      expectedRows.push(definition);
      expectedKeySet[definition.key] = true;
    });
  });

  const existingRows = formSheet.getLastRow() > 1
    ? formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 6).getValues()
    : [];
  const rowByKey = {};
  existingRows.forEach((row, index) => {
    const key = String(row[0] || "").trim();
    if (!key) return;
    rowByKey[key] = { rowIndex: index + 2, row };
  });

  let createdRows = 0;
  let updatedRows = 0;
  let disabledRows = 0;
  const appendRows = [];

  expectedRows.forEach((definition) => {
    const desired = [
      definition.key,
      definition.label,
      definition.url,
      definition.module,
      definition.siteKey,
      true,
    ];
    const current = rowByKey[definition.key];
    if (!current) {
      appendRows.push(desired);
      createdRows += 1;
      return;
    }

    const currentValues = current.row.map((value) => String(value === undefined || value === null ? "" : value).trim());
    const desiredValues = desired.map((value) => String(value === undefined || value === null ? "" : value).trim());
    if (currentValues.join("|") !== desiredValues.join("|")) {
      formSheet.getRange(current.rowIndex, 1, 1, 6).setValues([desired]);
      updatedRows += 1;
    }
  });

  if (appendRows.length) {
    formSheet.getRange(formSheet.getLastRow() + 1, 1, appendRows.length, 6).setValues(appendRows);
  }

  existingRows.forEach((row, index) => {
    const key = String(row[0] || "").trim();
    if (!key || expectedKeySet[key]) return;
    if (!isTruthy_(row[5])) return;
    formSheet.getRange(index + 2, 6).setValue(false);
    disabledRows += 1;
  });

  return {
    sourcesChecked: sourceRows.length,
    expectedForms: expectedRows.length,
    createdRows,
    updatedRows,
    disabledRows,
  };
}

function resolveNativeFormMeta_(formType, module) {
  const type = normalizeNativeFormType_(formType);
  const normalizedModule = String(module || "").trim().toLowerCase();
  if (type === NATIVE_FORM_TYPES.FIRE_MOVEMENT) {
    return {
      type,
      module: "incendie",
      mode: "movement",
      title: "Mouvement stock incendie",
      description: "Saisie native d'un mouvement stock incendie.",
    };
  }
  if (isFireReplenishmentFormType_(type)) {
    return {
      type: NATIVE_FORM_TYPES.FIRE_REPLENISHMENT,
      module: "incendie",
      mode: "replenishment",
      title: "Demande de reapprovisionnement incendie",
      description: "Demande de reappro incendie (alias legacy FORM_FIRE_INVENTORY_*).",
    };
  }
  if (type === NATIVE_FORM_TYPES.FIRE_ITEM_CREATE) {
    return {
      type,
      module: "incendie",
      mode: "item_create",
      title: "Creation article incendie",
      description: "Creation d'un article incendie avec mouvement IN initial.",
    };
  }
  if (type === NATIVE_FORM_TYPES.PHARMA_MOVEMENT) {
    return {
      type,
      module: "pharmacie",
      mode: "movement",
      title: "Mouvement stock pharmacie",
      description: "Saisie native d'un mouvement stock pharmacie.",
    };
  }
  if (type === NATIVE_FORM_TYPES.PHARMA_INVENTORY) {
    return {
      type,
      module: "pharmacie",
      mode: "inventory",
      title: "Inventaire pharmacie",
      description: "Comptage physique pharmacie avec controleur.",
    };
  }
  if (type === NATIVE_FORM_TYPES.PHARMA_ITEM_CREATE) {
    return {
      type,
      module: "pharmacie",
      mode: "item_create",
      title: "Creation article pharmacie",
      description: "Creation d'un article pharmacie avec mouvement IN initial.",
    };
  }
  if (normalizedModule === "incendie") {
    return resolveNativeFormMeta_(NATIVE_FORM_TYPES.FIRE_MOVEMENT, normalizedModule);
  }
  if (normalizedModule === "pharmacie") {
    return resolveNativeFormMeta_(NATIVE_FORM_TYPES.PHARMA_MOVEMENT, normalizedModule);
  }
  throw new Error(`Type de formulaire natif non supporte: ${formType}`);
}

function assertNativeFormAccess_(formType) {
  const dashboard = resolveNativeFormsDashboard_();
  const access = getCurrentUserAccessProfile_(dashboard);
  const type = normalizeNativeFormType_(formType);
  const itemCreate = type === NATIVE_FORM_TYPES.FIRE_ITEM_CREATE || type === NATIVE_FORM_TYPES.PHARMA_ITEM_CREATE;
  if (itemCreate && !(access.isAdmin || access.isController)) {
    throw new Error("Action reservee aux roles ADMIN / CONTROLEUR.");
  }
  return { dashboard, access };
}

function collectNativeSitesForModule_(dashboard, module, access) {
  const normalizedModule = normalizeModuleKey_(module);
  if (!normalizedModule) return [];
  if (access) {
    assertModuleAccess_(access, normalizedModule, "Chargement sites formulaire natif");
  }
  const sourceSheet = dashboard ? dashboard.getSheetByName("CONFIG_SOURCES") : null;
  const sites = {};
  if (sourceSheet && sourceSheet.getLastRow() >= 2) {
    const rows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
    rows.forEach((row) => {
      const rowModule = normalizeModuleKey_(row[1]);
      const siteKey = String(row[2] || "").trim();
      const enabled = row[8];
      if (!siteKey || rowModule !== normalizedModule || !isTruthy_(enabled)) return;
      if (access && !hasSiteAccess_(access, siteKey)) return;
      sites[siteKey] = true;
    });
  }

  const siteList = Object.keys(sites).sort();
  if (!siteList.length) {
    (DEPLOYMENT_CONFIG.sites || []).forEach((siteKey) => {
      const normalized = String(siteKey || "").trim();
      if (access && normalized && !hasSiteAccess_(access, normalized)) return;
      if (normalized) sites[normalized] = true;
    });
  }

  return Object.keys(sites)
    .sort()
    .map((siteKey) => ({ siteKey, label: siteKey }));
}

function resolveSourceWorkbookContext_(dashboard, module, siteKey, access, actionLabel) {
  const normalizedModule = normalizeModuleKey_(module);
  const requestedSite = String(siteKey || "").trim();
  if (!normalizedModule) {
    throw new Error("Module invalide.");
  }
  if (!requestedSite) {
    throw new Error("Le site est obligatoire.");
  }
  if (access) {
    assertModuleAccess_(access, normalizedModule, actionLabel || "Formulaire natif");
    assertSiteAccess_(access, requestedSite, actionLabel || "Formulaire natif");
  }

  const sourceSheet = dashboard ? dashboard.getSheetByName("CONFIG_SOURCES") : null;
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    throw new Error("CONFIG_SOURCES vide ou introuvable.");
  }

  const rows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  for (let i = 0; i < rows.length; i += 1) {
    const sourceKey = String(rows[i][0] || "").trim();
    const rowModule = normalizeModuleKey_(rows[i][1]);
    const rowSite = String(rows[i][2] || "").trim();
    const workbookUrl = String(rows[i][3] || "").trim();
    const enabled = rows[i][8];
    if (!sourceKey || !workbookUrl || !isTruthy_(enabled)) continue;
    if (rowModule !== normalizedModule || normalizeSiteKey_(rowSite) !== normalizeSiteKey_(requestedSite)) continue;
    return {
      sourceKey,
      module: rowModule,
      siteKey: rowSite,
      workbookUrl,
      workbook: getCachedSpreadsheetByUrl_(workbookUrl),
    };
  }
  throw new Error(`Aucune source active trouvee pour ${normalizedModule}/${requestedSite}.`);
}

function buildNativeFormChoicesForSite_(workbook, module) {
  const context = getModuleItemContext_(module);
  if (!context || !workbook) {
    return { itemChoices: [], replenishmentChoices: [] };
  }
  return {
    itemChoices: getItemChoiceValues_(workbook, context.itemsTab, context.stockTab).slice(0, LIVE_FORM_MAX_CHOICES),
    replenishmentChoices: getReplenishmentItemChoiceValues_(workbook, context.stockTab).slice(0, LIVE_FORM_MAX_CHOICES),
  };
}

function normalizeStringArray_(value) {
  if (Array.isArray(value)) {
    return value
      .map((entry) => String(entry || "").trim())
      .filter((entry) => entry !== "");
  }
  const raw = String(value || "").trim();
  if (!raw) return [];
  return raw
    .split(",")
    .map((entry) => String(entry || "").trim())
    .filter((entry) => entry !== "");
}

function formPropertyKey_(formKey) {
  const normalized = String(formKey || "").trim().toUpperCase();
  return normalized ? `${FORM_ID_PROPERTY_PREFIX}${normalized}` : "";
}

function setStoredFormId_(formKey, formId) {
  const key = formPropertyKey_(formKey);
  const value = String(formId || "").trim();
  if (!key || !value) return;
  PropertiesService.getDocumentProperties().setProperty(key, value);
}

function getStoredFormId_(formKey) {
  const key = formPropertyKey_(formKey);
  if (!key) return "";
  return String(PropertiesService.getDocumentProperties().getProperty(key) || "").trim();
}

function applyFriendlyFormDefinition_(form, formKey, siteKey, module, workbook, availableSites) {
  const key = String(formKey || "").toUpperCase();
  ensureFormEmailCollection_(form);
  resetFormItems_(form);

  if (key.indexOf("FORM_FIRE_MOVEMENT_") === 0) {
    form.setTitle(`Mouvement de stock incendie - ${siteKey}`);
    form.setDescription("Declaration de mouvement (entree, sortie, correction) pour le stock incendie.");
    addSiteQuestion_(form, siteKey, availableSites);
    addItemQuestion_(form, workbook, "ITEMS", "STOCK_VIEW");
    form.addMultipleChoiceItem()
      .setTitle("Type de mouvement")
      .setHelpText("DELETE_ITEM reserve aux profils ADMIN / CONTROLEUR (controle serveur).")
      .setChoiceValues(["IN", "OUT", "CORRECTION", "DELETE_ITEM"])
      .setRequired(true);
    form.addTextItem().setTitle("Quantite").setHelpText("Mettre 0 si l'action est une suppression d'article.").setRequired(true);
    form.addTextItem().setTitle("Cout unitaire").setRequired(false);
    form.addTextItem().setTitle("Motif").setRequired(true);
    form.addTextItem().setTitle("Email operateur").setHelpText("Laisser vide pour utiliser automatiquement l'email du compte connecte.").setRequired(false);
    form.addTextItem().setTitle("Reference document").setRequired(false);
    form.addParagraphTextItem().setTitle("Commentaire").setRequired(false);
    addDeleteItemsQuestion_(form, workbook, "ITEMS", "STOCK_VIEW");
    form.addTextItem().setTitle("Categorie article (optionnel)").setRequired(false);
    form.addTextItem().setTitle("Unite article (optionnel)").setRequired(false);
    form.addTextItem().setTitle("Seuil mini article (optionnel)").setRequired(false);
    return;
  }

  if (isLegacyFireInventoryFormKey_(key)) {
    // Compatibilite legacy: FORM_FIRE_INVENTORY_* correspond au flux metier de reappro incendie.
    form.setTitle(`Demande de reapprovisionnement incendie - ${siteKey}`);
    form.setDescription("Formulaire de demande de reapprovisionnement en cas de rupture/sous-seuil. La demande sera validee par un controleur ou un admin avant increment du stock.");
    addSiteQuestion_(form, siteKey, availableSites);
    addReplenishmentItemQuestion_(form, workbook, "ITEMS", "STOCK_VIEW");
    form.addTextItem().setTitle("Quantite demandee").setHelpText("Quantite a ajouter au stock apres validation.").setRequired(true);
    form.addMultipleChoiceItem().setTitle("Priorite").setChoiceValues(["HIGH", "MEDIUM", "LOW"]).setRequired(true);
    form.addTextItem().setTitle("Motif de la demande").setRequired(true);
    form.addTextItem().setTitle("Email demandeur").setHelpText("Laisser vide pour utiliser automatiquement l'email du compte connecte.").setRequired(false);
    form.addParagraphTextItem().setTitle("Commentaire").setRequired(false);
    return;
  }

  if (key.indexOf("FORM_FIRE_ITEM_CREATE_") === 0) {
    form.setTitle(`Creation nouvel article incendie - ${siteKey}`);
    form.setDescription("Creation d'un nouvel article incendie avec stock initial.");
    addSiteQuestion_(form, siteKey, availableSites);
    addNewItemCreationQuestions_(form, false);
    return;
  }

  if (key.indexOf("FORM_PHARMA_MOVEMENT_") === 0) {
    form.setTitle(`Mouvement stock pharmacie - ${siteKey}`);
    form.setDescription("Declaration de mouvement (entree, sortie, correction) pour la pharmacie.");
    addSiteQuestion_(form, siteKey, availableSites);
    addItemQuestion_(form, workbook, "ITEMS_PHARMACY", "STOCK_VIEW_PHARMACY");
    form.addMultipleChoiceItem()
      .setTitle("Type de mouvement")
      .setHelpText("DELETE_ITEM reserve aux profils ADMIN / CONTROLEUR (controle serveur).")
      .setChoiceValues(["IN", "OUT", "CORRECTION", "DELETE_ITEM"])
      .setRequired(true);
    form.addTextItem().setTitle("Quantite").setHelpText("Mettre 0 si l'action est une suppression d'article.").setRequired(true);
    form.addTextItem().setTitle("Cout unitaire").setRequired(false);
    form.addTextItem().setTitle("Motif").setRequired(true);
    form.addTextItem().setTitle("Email operateur").setHelpText("Laisser vide pour utiliser automatiquement l'email du compte connecte.").setRequired(false);
    form.addTextItem().setTitle("Numero de lot").setRequired(false);
    form.addTextItem().setTitle("Date de peremption").setRequired(false);
    form.addTextItem().setTitle("Reference document").setRequired(false);
    form.addParagraphTextItem().setTitle("Commentaire").setRequired(false);
    addDeleteItemsQuestion_(form, workbook, "ITEMS_PHARMACY", "STOCK_VIEW_PHARMACY");
    form.addTextItem().setTitle("Categorie article (optionnel)").setRequired(false);
    form.addTextItem().setTitle("Unite article (optionnel)").setRequired(false);
    form.addTextItem().setTitle("Seuil mini article (optionnel)").setRequired(false);
    return;
  }

  if (key.indexOf("FORM_PHARMA_ITEM_CREATE_") === 0) {
    form.setTitle(`Creation nouvel article pharmacie - ${siteKey}`);
    form.setDescription("Creation d'un nouvel article pharmacie avec stock initial.");
    addSiteQuestion_(form, siteKey, availableSites);
    addNewItemCreationQuestions_(form, true);
    return;
  }

  if (key.indexOf("FORM_PHARMA_INVENTORY_") === 0) {
    form.setTitle(`Inventaire stock pharmacie - ${siteKey}`);
    form.setDescription("Comptage physique pour mise a jour des quantites pharmacie.");
    addSiteQuestion_(form, siteKey, availableSites);
    addItemQuestion_(form, workbook, "ITEMS_PHARMACY", "STOCK_VIEW_PHARMACY");
    form.addTextItem().setTitle("Quantite comptee").setHelpText("Mettre 0 si l'action est une suppression d'article.").setRequired(true);
    form.addTextItem().setTitle("Email controleur").setHelpText("Laisser vide pour utiliser automatiquement l'email du compte connecte.").setRequired(false);
    form.addParagraphTextItem().setTitle("Commentaire").setRequired(false);
    addDeleteItemsQuestion_(form, workbook, "ITEMS_PHARMACY", "STOCK_VIEW_PHARMACY");
    form.addTextItem().setTitle("Categorie article (optionnel)").setRequired(false);
    form.addTextItem().setTitle("Unite article (optionnel)").setRequired(false);
    form.addTextItem().setTitle("Seuil mini article (optionnel)").setRequired(false);
  }
}

function resetFormItems_(form) {
  const items = form.getItems();
  for (let i = items.length - 1; i >= 0; i -= 1) {
    form.deleteItem(items[i]);
  }
}

function ensureFormEmailCollection_(form) {
  if (!form) return;
  try {
    if (typeof form.setCollectEmail === "function") {
      form.setCollectEmail(true);
    }
  } catch (error) {
    Logger.log(`Collect email setup warning: ${String(error.message || error)}`);
  }
}

function getSiteChoiceValues_(siteKey, availableSites) {
  const rawValues = (availableSites || DEPLOYMENT_CONFIG.sites || [])
    .map((value) => String(value || "").trim())
    .filter((value) => value !== "");
  const values = [];
  rawValues.forEach((value) => {
    if (values.indexOf(value) === -1) values.push(value);
  });
  if (siteKey && values.indexOf(siteKey) === -1) values.unshift(siteKey);
  if (!values.length && siteKey) return [siteKey];
  return values.length ? values : ["JLL"];
}

function getItemChoiceValues_(spreadsheet, itemsTab, stockTab) {
  const stockSheet = stockTab ? spreadsheet.getSheetByName(stockTab) : null;
  if (stockSheet && stockSheet.getLastRow() >= 2) {
    const stockRows = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 8).getValues();
    const withStock = stockRows
      .map((row) => {
        const itemId = String(row[0] || "").trim();
        const itemName = String(row[1] || "").trim();
        const stock = row[5];
        const threshold = row[4];
        if (!itemId) return "";
        return `${itemId} | ${itemName || itemId} | stock:${stock} | seuil:${threshold}`;
      })
      .filter((value) => value !== "");
    if (withStock.length) return withStock;
  }

  const itemsSheet = spreadsheet.getSheetByName(itemsTab);
  if (!itemsSheet || itemsSheet.getLastRow() < 2) return [];
  const rows = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 8).getValues();
  return rows
    .map((row) => {
      const itemId = String(row[0] || "").trim();
      const itemName = String(row[3] || "").trim();
      const isActive = row[7];
      if (!isTruthy_(isActive)) return "";
      if (!itemId) return "";
      return itemName ? `${itemId} | ${itemName}` : itemId;
    })
    .filter((value) => value !== "");
}

function buildSitePrefilledUrl_(form, siteKey) {
  const items = form.getItems();
  if (!items || !items.length || !siteKey) return "";

  const response = form.createResponse();
  const siteResponse = createSiteItemResponse_(items[0], siteKey);
  if (!siteResponse) return "";
  response.withItemResponse(siteResponse);
  return response.toPrefilledUrl();
}

function createSiteItemResponse_(item, siteKey) {
  const type = item.getType();
  if (type === FormApp.ItemType.TEXT) {
    return item.asTextItem().createResponse(siteKey);
  }
  if (type === FormApp.ItemType.LIST) {
    return item.asListItem().createResponse(siteKey);
  }
  if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    return item.asMultipleChoiceItem().createResponse(siteKey);
  }
  return null;
}

function refreshAllFormsLiveChoices_() {
  const dashboard = openDashboardSpreadsheet_();
  if (!dashboard) throw new Error("Dashboard introuvable pour rafraichir les choix des formulaires.");
  if (isNativeFormsMode_()) {
    const sync = ensureNativeFormLinksFromDashboard_(dashboard);
    return { updatedForms: sync.createdRows + sync.updatedRows, mode: "NATIVE_HTML" };
  }

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  const formSheet = dashboard.getSheetByName("FORM_LINKS");
  if (!sourceSheet || !formSheet || formSheet.getLastRow() < 2) return { updatedForms: 0 };

  const sourceRows = sourceSheet.getLastRow() > 1
    ? sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues()
    : [];
  const sourceMap = {};
  sourceRows.forEach((row) => {
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    if (!module || !siteKey || !workbookUrl || !isTruthy_(row[8])) return;
    sourceMap[`${module}|${siteKey}`] = workbookUrl;
  });

  const formRows = formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 6).getValues();
  const updatedUrls = formRows.map((row) => [row[2]]);
  let updated = 0;

  formRows.forEach((row, index) => {
    const formKey = String(row[0] || "").trim();
    const formLabel = String(row[1] || "").trim();
    const formUrl = String(row[2] || "").trim();
    const module = String(row[3] || "").trim().toLowerCase();
    const siteKey = String(row[4] || "").trim();
    const isActive = row[5];
    if (!formKey || !module || !siteKey || !isTruthy_(isActive)) return;

    const workbookUrl = sourceMap[`${module}|${siteKey}`];
    if (!workbookUrl) return;

    try {
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      const form = openManagedForm_(dashboard, formKey, formUrl, siteKey, formLabel);
      const changed = refreshLiveChoicesOnForm_(form, workbook, module);
      if (changed) {
        const prefilled = buildSitePrefilledUrl_(form, siteKey) || form.getPublishedUrl() || form.getEditUrl();
        updatedUrls[index][0] = prefilled;
        updated += 1;
      }
    } catch (error) {
      Logger.log(`Live form refresh ignored (${formKey}): ${String(error.message || error)}`);
    }
  });

  if (updated > 0) {
    formSheet.getRange(2, 3, updatedUrls.length, 1).setValues(updatedUrls);
  }
  return { updatedForms: updated };
}

function refreshLinkedFormsChoicesForWorkbook_(workbook) {
  if (!workbook) return 0;
  if (isNativeFormsMode_()) return 0;
  const dashboard = openDashboardSpreadsheet_();
  if (!dashboard) return 0;

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  const formSheet = dashboard.getSheetByName("FORM_LINKS");
  if (!sourceSheet || !formSheet || sourceSheet.getLastRow() < 2 || formSheet.getLastRow() < 2) return 0;

  const workbookId = workbook.getId();
  const sourceRows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  const matches = sourceRows.filter((row) => {
    const url = String(row[3] || "").trim();
    const enabled = row[8];
    return isTruthy_(enabled) && extractIdFromDriveUrl_(url) === workbookId;
  });
  if (!matches.length) return 0;

  const allowed = {};
  matches.forEach((row) => {
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    if (module && siteKey) allowed[`${module}|${siteKey}`] = true;
  });

  const formRows = formSheet.getRange(2, 1, formSheet.getLastRow() - 1, 6).getValues();
  const updatedUrls = formRows.map((row) => [row[2]]);
  let updated = 0;

  formRows.forEach((row, index) => {
    const formKey = String(row[0] || "").trim();
    const formLabel = String(row[1] || "").trim();
    const formUrl = String(row[2] || "").trim();
    const module = String(row[3] || "").trim().toLowerCase();
    const siteKey = String(row[4] || "").trim();
    const isActive = row[5];
    if (!formKey || !module || !siteKey || !isTruthy_(isActive)) return;
    if (!allowed[`${module}|${siteKey}`]) return;

    try {
      const form = openManagedForm_(dashboard, formKey, formUrl, siteKey, formLabel);
      const changed = refreshLiveChoicesOnForm_(form, workbook, module);
      if (changed) {
        const prefilled = buildSitePrefilledUrl_(form, siteKey) || form.getPublishedUrl() || form.getEditUrl();
        updatedUrls[index][0] = prefilled;
        updated += 1;
      }
    } catch (error) {
      Logger.log(`Live workbook refresh ignored (${formKey}): ${String(error.message || error)}`);
    }
  });

  if (updated > 0) {
    formSheet.getRange(2, 3, updatedUrls.length, 1).setValues(updatedUrls);
  }
  return updated;
}

function addSiteQuestion_(form, siteKey, availableSites) {
  form.addListItem()
    .setTitle("Site concerne")
    .setChoiceValues(getSiteChoiceValues_(siteKey, availableSites))
    .setRequired(true);
}

function addItemQuestion_(form, spreadsheet, itemsTab, stockTab) {
  if (!spreadsheet) {
    form.addTextItem()
      .setTitle("Article concerne")
      .setHelpText("Saisir ItemID existant ou nom d'un nouvel article")
      .setRequired(true);
    return;
  }

  const choices = getItemChoiceValues_(spreadsheet, itemsTab, stockTab);
  if (!choices.length) {
    form.addTextItem()
      .setTitle("Article concerne")
      .setHelpText("Saisir ItemID existant ou nom d'un nouvel article")
      .setRequired(true);
    return;
  }

  form.addMultipleChoiceItem()
    .setTitle("Article concerne")
    .setHelpText("Choisir un article existant ou saisir un nouvel article (Autre)")
    .setChoiceValues(choices)
    .showOtherOption(true)
    .setRequired(true);
}

function getReplenishmentItemChoiceValues_(spreadsheet, stockTab) {
  const stockSheet = spreadsheet && stockTab ? spreadsheet.getSheetByName(stockTab) : null;
  if (!stockSheet || stockSheet.getLastRow() < 2) return [];
  const rows = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 8).getValues();
  return rows
    .map((row) => {
      const itemId = String(row[0] || "").trim();
      const itemName = String(row[1] || "").trim();
      const threshold = row[4];
      const stock = row[5];
      const status = normalizeTextKey_(row[7]);
      if (!itemId) return "";
      if (status !== "rupture" && status !== "sous_seuil") return "";
      return `${itemId} | ${itemName || itemId} | stock:${stock} | seuil:${threshold}`;
    })
    .filter((value) => value !== "");
}

function addReplenishmentItemQuestion_(form, spreadsheet, itemsTab, stockTab) {
  if (!spreadsheet) {
    addItemQuestion_(form, spreadsheet, itemsTab, stockTab);
    return;
  }

  const choices = getReplenishmentItemChoiceValues_(spreadsheet, stockTab);
  if (!choices.length) {
    addItemQuestion_(form, spreadsheet, itemsTab, stockTab);
    return;
  }

  form.addMultipleChoiceItem()
    .setTitle("Article a reapprovisionner")
    .setHelpText("Choisir un article en rupture/sous-seuil, ou saisir un autre article (Autre).")
    .setChoiceValues(choices)
    .showOtherOption(true)
    .setRequired(true);
}

function addDeleteItemsQuestion_(form, spreadsheet, itemsTab, stockTab) {
  if (!spreadsheet) {
    form.addTextItem()
      .setTitle("Supprimer des articles (optionnel)")
      .setHelpText("Saisir un ou plusieurs ItemID separes par virgule. Action reservee ADMIN / CONTROLEUR.")
      .setRequired(false);
    return;
  }

  const choices = getItemChoiceValues_(spreadsheet, itemsTab, stockTab).slice(0, LIVE_FORM_MAX_CHOICES);
  if (!choices.length) {
    form.addTextItem()
      .setTitle("Supprimer des articles (optionnel)")
      .setHelpText("Saisir un ou plusieurs ItemID separes par virgule. Action reservee ADMIN / CONTROLEUR.")
      .setRequired(false);
    return;
  }

  form.addCheckboxItem()
    .setTitle("Supprimer des articles (optionnel)")
    .setHelpText("Selection multiple possible. Action reservee ADMIN / CONTROLEUR (controle serveur).")
    .setChoiceValues(choices)
    .setRequired(false);
}

function addNewItemCreationQuestions_(form, includeLotFields) {
  form.addTextItem()
    .setTitle("ItemID")
    .setHelpText("Code article unique (ex: INC-0100, PHA-0100).")
    .setRequired(true);
  form.addTextItem()
    .setTitle("Nom article")
    .setRequired(true);
  form.addTextItem().setTitle("Categorie article (optionnel)").setRequired(false);
  form.addTextItem().setTitle("Unite article (optionnel)").setRequired(false);
  form.addTextItem().setTitle("Seuil mini article (optionnel)").setRequired(false);
  form.addTextItem()
    .setTitle("Quantite initiale")
    .setHelpText("Quantite initiale entree en stock (0 autorise).")
    .setRequired(true);
  form.addMultipleChoiceItem()
    .setTitle("Type de mouvement")
    .setChoiceValues(["IN"])
    .setRequired(true);
  form.addTextItem().setTitle("Motif").setRequired(true);
  form.addTextItem()
    .setTitle("Email operateur")
    .setHelpText("Laisser vide pour utiliser automatiquement l'email du compte connecte.")
    .setRequired(false);
  if (includeLotFields) {
    form.addTextItem().setTitle("Numero de lot").setRequired(false);
    form.addTextItem().setTitle("Date de peremption").setRequired(false);
  }
  form.addTextItem().setTitle("Reference document").setRequired(false);
  form.addParagraphTextItem().setTitle("Commentaire").setRequired(false);
}

function refreshLiveChoicesOnForm_(form, workbook, module) {
  const context = getModuleItemContext_(module);
  if (!context || !form || !workbook) return false;

  const choices = getItemChoiceValues_(workbook, context.itemsTab, context.stockTab).slice(0, LIVE_FORM_MAX_CHOICES);
  const replenishmentChoices = getReplenishmentItemChoiceValues_(workbook, context.stockTab).slice(0, LIVE_FORM_MAX_CHOICES);
  if (!choices.length) return false;

  let changed = false;
  form.getItems().forEach((item) => {
    const title = normalizeTextKey_(item.getTitle());
    if (
      item.getType() === FormApp.ItemType.MULTIPLE_CHOICE
      && (title.indexOf("article concerne") !== -1 || title.indexOf("article a reapprovisionner") !== -1)
    ) {
      const targetChoices = title.indexOf("article a reapprovisionner") !== -1 && replenishmentChoices.length
        ? replenishmentChoices
        : choices;
      item.asMultipleChoiceItem().setChoiceValues(targetChoices).showOtherOption(true);
      changed = true;
      return;
    }
    if (
      item.getType() === FormApp.ItemType.CHECKBOX
      && (title.indexOf("supprimer des articles") !== -1 || title.indexOf("articles a supprimer") !== -1)
    ) {
      item.asCheckboxItem().setChoiceValues(choices);
      changed = true;
    }
  });
  return changed;
}

function buildFireMovementForm_(folder, siteKey, spreadsheet) {
  const form = FormApp.create(`Mouvement de stock incendie - ${siteKey}`);
  form.setDescription("Declaration de mouvement (entree, sortie, correction) pour le stock incendie.");
  addSiteQuestion_(form, siteKey, [siteKey]);
  addItemQuestion_(form, spreadsheet, "ITEMS", "STOCK_VIEW");
  form.addMultipleChoiceItem()
    .setTitle("Type de mouvement")
    .setHelpText("DELETE_ITEM reserve aux profils ADMIN / CONTROLEUR (controle serveur).")
    .setChoiceValues(["IN", "OUT", "CORRECTION", "DELETE_ITEM"])
    .setRequired(true);
  form.addTextItem().setTitle("Quantite").setHelpText("Mettre 0 si l'action est une suppression d'article.").setRequired(true);
  form.addTextItem().setTitle("Cout unitaire").setRequired(false);
  form.addTextItem().setTitle("Motif").setRequired(true);
  form.addTextItem().setTitle("Email operateur").setHelpText("Laisser vide pour utiliser automatiquement l'email du compte connecte.").setRequired(false);
  form.addTextItem().setTitle("Reference document").setRequired(false);
  form.addParagraphTextItem().setTitle("Commentaire").setRequired(false);
  addDeleteItemsQuestion_(form, spreadsheet, "ITEMS", "STOCK_VIEW");
  form.addTextItem().setTitle("Categorie article (optionnel)").setRequired(false);
  form.addTextItem().setTitle("Unite article (optionnel)").setRequired(false);
  form.addTextItem().setTitle("Seuil mini article (optionnel)").setRequired(false);
  return finalizeForm_(form, folder, siteKey);
}

function buildFireInventoryForm_(folder, siteKey, spreadsheet) {
  const form = FormApp.create(`Demande de reapprovisionnement incendie - ${siteKey}`);
  form.setDescription("Formulaire de demande de reapprovisionnement en cas de rupture/sous-seuil. La demande sera validee par un controleur ou un admin avant increment du stock.");
  addSiteQuestion_(form, siteKey, [siteKey]);
  addReplenishmentItemQuestion_(form, spreadsheet, "ITEMS", "STOCK_VIEW");
  form.addTextItem().setTitle("Quantite demandee").setHelpText("Quantite a ajouter au stock apres validation.").setRequired(true);
  form.addMultipleChoiceItem().setTitle("Priorite").setChoiceValues(["HIGH", "MEDIUM", "LOW"]).setRequired(true);
  form.addTextItem().setTitle("Motif de la demande").setRequired(true);
  form.addTextItem().setTitle("Email demandeur").setHelpText("Laisser vide pour utiliser automatiquement l'email du compte connecte.").setRequired(false);
  form.addParagraphTextItem().setTitle("Commentaire").setRequired(false);
  return finalizeForm_(form, folder, siteKey);
}

function buildFireItemCreationForm_(folder, siteKey, spreadsheet) {
  const form = FormApp.create(`Creation nouvel article incendie - ${siteKey}`);
  form.setDescription("Creation d'un nouvel article incendie avec stock initial.");
  addSiteQuestion_(form, siteKey, [siteKey]);
  addNewItemCreationQuestions_(form, false);
  return finalizeForm_(form, folder, siteKey);
}

function buildPharmaMovementForm_(folder, siteKey, spreadsheet) {
  const form = FormApp.create(`Mouvement stock pharmacie - ${siteKey}`);
  form.setDescription("Declaration de mouvement (entree, sortie, correction) pour la pharmacie.");
  addSiteQuestion_(form, siteKey, [siteKey]);
  addItemQuestion_(form, spreadsheet, "ITEMS_PHARMACY", "STOCK_VIEW_PHARMACY");
  form.addMultipleChoiceItem()
    .setTitle("Type de mouvement")
    .setHelpText("DELETE_ITEM reserve aux profils ADMIN / CONTROLEUR (controle serveur).")
    .setChoiceValues(["IN", "OUT", "CORRECTION", "DELETE_ITEM"])
    .setRequired(true);
  form.addTextItem().setTitle("Quantite").setHelpText("Mettre 0 si l'action est une suppression d'article.").setRequired(true);
  form.addTextItem().setTitle("Cout unitaire").setRequired(false);
  form.addTextItem().setTitle("Motif").setRequired(true);
  form.addTextItem().setTitle("Email operateur").setHelpText("Laisser vide pour utiliser automatiquement l'email du compte connecte.").setRequired(false);
  form.addTextItem().setTitle("Numero de lot").setRequired(false);
  form.addTextItem().setTitle("Date de peremption").setRequired(false);
  form.addTextItem().setTitle("Reference document").setRequired(false);
  form.addParagraphTextItem().setTitle("Commentaire").setRequired(false);
  addDeleteItemsQuestion_(form, spreadsheet, "ITEMS_PHARMACY", "STOCK_VIEW_PHARMACY");
  form.addTextItem().setTitle("Categorie article (optionnel)").setRequired(false);
  form.addTextItem().setTitle("Unite article (optionnel)").setRequired(false);
  form.addTextItem().setTitle("Seuil mini article (optionnel)").setRequired(false);
  return finalizeForm_(form, folder, siteKey);
}

function buildPharmaItemCreationForm_(folder, siteKey, spreadsheet) {
  const form = FormApp.create(`Creation nouvel article pharmacie - ${siteKey}`);
  form.setDescription("Creation d'un nouvel article pharmacie avec stock initial.");
  addSiteQuestion_(form, siteKey, [siteKey]);
  addNewItemCreationQuestions_(form, true);
  return finalizeForm_(form, folder, siteKey);
}

function buildPharmaInventoryForm_(folder, siteKey, spreadsheet) {
  const form = FormApp.create(`Inventaire stock pharmacie - ${siteKey}`);
  form.setDescription("Comptage physique pour mise a jour des quantites pharmacie.");
  addSiteQuestion_(form, siteKey, [siteKey]);
  addItemQuestion_(form, spreadsheet, "ITEMS_PHARMACY", "STOCK_VIEW_PHARMACY");
  form.addTextItem().setTitle("Quantite comptee").setHelpText("Mettre 0 si l'action est une suppression d'article.").setRequired(true);
  form.addTextItem().setTitle("Email controleur").setHelpText("Laisser vide pour utiliser automatiquement l'email du compte connecte.").setRequired(false);
  form.addParagraphTextItem().setTitle("Commentaire").setRequired(false);
  addDeleteItemsQuestion_(form, spreadsheet, "ITEMS_PHARMACY", "STOCK_VIEW_PHARMACY");
  form.addTextItem().setTitle("Categorie article (optionnel)").setRequired(false);
  form.addTextItem().setTitle("Unite article (optionnel)").setRequired(false);
  form.addTextItem().setTitle("Seuil mini article (optionnel)").setRequired(false);
  return finalizeForm_(form, folder, siteKey);
}

function finalizeForm_(form, folder, siteKey) {
  ensureFormEmailCollection_(form);
  form.setAcceptingResponses(true);
  moveFileToFolder_(getCachedFileById_(form.getId()), folder);
  const prefilledUrl = buildSitePrefilledUrl_(form, siteKey);
  const publishedUrl = form.getPublishedUrl();
  return { form, id: form.getId(), url: prefilledUrl || publishedUrl || form.getEditUrl() };
}

function openFormByUrlOrId_(formRef) {
  const ref = String(formRef || "").trim();
  if (!ref) throw new Error("Reference formulaire vide.");
  if (/^https?:\/\//i.test(ref)) {
    try {
      return getCachedFormByUrl_(ref);
    } catch (openByUrlError) {
      const extractedFromUrl = extractFormIdFromUrl_(ref);
      if (extractedFromUrl) {
        return getCachedFormById_(extractedFromUrl);
      }
      throw openByUrlError;
    }
  }
  try {
    return getCachedFormById_(ref);
  } catch (openByIdError) {
    const extractedId = extractFormIdFromUrl_(ref);
    if (extractedId && extractedId !== ref) {
      return getCachedFormById_(extractedId);
    }
    throw openByIdError;
  }
}

function openManagedForm_(dashboard, formKey, formRef, siteKey, formLabel) {
  const attempts = [];
  const seen = {};
  const candidates = [];
  const pushCandidate = (label, value) => {
    const candidate = String(value || "").trim();
    if (!candidate || seen[candidate]) return;
    seen[candidate] = true;
    candidates.push({ label, value: candidate });
  };

  pushCandidate("url", formRef);
  pushCandidate("storedId", getStoredFormId_(formKey));
  pushCandidate("logId", findFormIdInDeploymentLog_(dashboard, formKey, siteKey, formLabel));
  pushCandidate("urlExtractedId", extractFormIdFromUrl_(formRef));

  for (let i = 0; i < candidates.length; i += 1) {
    const candidate = candidates[i];
    try {
      const form = openFormByUrlOrId_(candidate.value);
      setStoredFormId_(formKey, form.getId());
      return form;
    } catch (error) {
      attempts.push(`${candidate.label}: ${String(error.message || error)}`);
    }
  }

  throw new Error(`Formulaire inaccessible pour ${formKey}: ${attempts.join(" | ")}`);
}

function formLabelFromKey_(formKey, siteKey) {
  const key = String(formKey || "").toUpperCase();
  const site = String(siteKey || "").trim();
  if (key.indexOf("FORM_FIRE_MOVEMENT_") === 0) return `Form mouvement incendie ${site}`;
  if (isLegacyFireInventoryFormKey_(key)) return `Form demande reappro incendie ${site}`;
  if (key.indexOf("FORM_FIRE_ITEM_CREATE_") === 0) return `Form creation article incendie ${site}`;
  if (key.indexOf("FORM_PHARMA_MOVEMENT_") === 0) return `Form mouvement pharmacie ${site}`;
  if (key.indexOf("FORM_PHARMA_INVENTORY_") === 0) return `Form inventaire pharmacie ${site}`;
  if (key.indexOf("FORM_PHARMA_ITEM_CREATE_") === 0) return `Form creation article pharmacie ${site}`;
  return "";
}

function findFormIdInDeploymentLog_(dashboard, formKey, siteKey, formLabel) {
  if (!dashboard) return "";
  const logSheet = dashboard.getSheetByName("DEPLOYMENT_LOG");
  if (!logSheet || logSheet.getLastRow() < 2) return "";

  const expectedLabel = String(formLabel || "").trim() || formLabelFromKey_(formKey, siteKey);
  const expectedSite = String(siteKey || "").trim();
  const hints = formLogHintFromKey_(formKey);
  const rows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).getValues();

  for (let i = rows.length - 1; i >= 0; i -= 1) {
    const type = String(rows[i][0] || "").trim().toUpperCase();
    const rowSite = String(rows[i][1] || "").trim();
    const name = String(rows[i][2] || "").trim();
    const id = String(rows[i][4] || "").trim();
    if (type !== "FORM" || !id) continue;

    if (expectedLabel && name === expectedLabel && (!expectedSite || rowSite === expectedSite)) {
      return id;
    }

    const normalizedName = normalizeTextKey_(name);
    const typeMatch = hints.typeHint && normalizedName.indexOf(hints.typeHint) >= 0;
    const altTypeMatch = hints.altTypeHint && normalizedName.indexOf(hints.altTypeHint) >= 0;
    if (
      (!expectedSite || rowSite === expectedSite)
      && hints.moduleHint
      && (typeMatch || altTypeMatch)
      && normalizedName.indexOf(hints.moduleHint) >= 0
    ) {
      return id;
    }
  }

  return "";
}

function formLogHintFromKey_(formKey) {
  const key = String(formKey || "").toUpperCase();
  const moduleHint = key.indexOf("_PHARMA_") >= 0 ? "pharmacie" : key.indexOf("_FIRE_") >= 0 ? "incendie" : "";
  const typeHint = key.indexOf("_MOVEMENT_") >= 0
    ? "mouvement"
    : key.indexOf("_INVENTORY_") >= 0
      ? "inventaire"
      : key.indexOf("_ITEM_CREATE_") >= 0
        ? "creation"
        : "";
  const altTypeHint = key.indexOf("_INVENTORY_") >= 0
    ? "reappro"
    : key.indexOf("_ITEM_CREATE_") >= 0
      ? "article"
      : "";
  return { moduleHint, typeHint, altTypeHint };
}

function extractFormIdFromUrl_(url) {
  const value = String(url || "").trim();
  if (!value) return "";

  const publicMatch = value.match(/\/forms\/d\/e\/([a-zA-Z0-9_-]+)/i);
  if (publicMatch && publicMatch[1]) return publicMatch[1];

  const editMatch = value.match(/\/forms\/d\/([a-zA-Z0-9_-]+)/i);
  if (editMatch && editMatch[1]) return editMatch[1];

  const genericMatch = value.match(/\/d\/([a-zA-Z0-9_-]+)/i);
  return genericMatch && genericMatch[1] ? genericMatch[1] : "";
}
