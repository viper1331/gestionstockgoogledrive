/*
 * Bloc 3 - Calcul metier / modules
 * Moteur incendie/pharmacie, flux metier et traitements post-submission.
 */

function syncAllModuleItemsFromDashboard() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard) {
    throw new Error("Aucun classeur actif. Ouvrez le dashboard puis relancez syncAllModuleItemsFromDashboard.");
  }

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    throw new Error("CONFIG_SOURCES vide ou introuvable.");
  }

  const sourceRows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  const audit = [["SourceKey", "Module", "SiteKey", "AddedItems", "WorkbookUrl"]];
  let totalAdded = 0;

  sourceRows.forEach((row) => {
    const sourceKey = String(row[0] || "").trim();
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];
    if (!sourceKey || !module || !siteKey || !workbookUrl || !isTruthy_(enabled)) return;

    const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
    let added = 0;
    if (module === "incendie") {
      added = syncFireItems_(workbook);
    } else if (module === "pharmacie") {
      added = syncPharmaItems_(workbook);
    }
    totalAdded += added;
    audit.push([sourceKey, module, siteKey, added, workbookUrl]);
  });

  const auditSheet = dashboard.getSheetByName("ITEM_SYNC_AUDIT") || dashboard.insertSheet("ITEM_SYNC_AUDIT");
  auditSheet.clearContents();
  auditSheet.clearFormats();
  auditSheet.getRange(1, 1, audit.length, audit[0].length).setValues(audit);
  auditSheet.getRange(1, 1, 1, audit[0].length).setFontWeight("bold");
  auditSheet.autoResizeColumns(1, audit[0].length);
  SpreadsheetApp.flush();

  return { totalAdded, auditedSources: audit.length - 1 };
}

function onModuleFormSubmit_(e) {
  const spreadsheet = e && e.source ? e.source : SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) return;

  const emailAutofill = autofillEmailColumnsFromSubmission_(e);
  const rowDeletions = applyItemDeletionsFromSubmission_(e);

  if (spreadsheet.getSheetByName("ITEMS")) {
    syncFireItems_(spreadsheet);
    applyFireFormulas_(spreadsheet);
    protectSheets_(spreadsheet, ["MOVEMENTS", "INVENTORY_COUNT", "STOCK_VIEW", "ALERTS", "PURCHASE_REQUESTS"]);
  }
  if (spreadsheet.getSheetByName("ITEMS_PHARMACY")) {
    syncPharmaItems_(spreadsheet);
    applyPharmaFormulas_(spreadsheet);
    protectSheets_(spreadsheet, ["MOVEMENTS_PHARMACY", "INVENTORY_COUNT_PHARMACY", "STOCK_VIEW_PHARMACY", "ALERTS_PHARMACY", "PURCHASE_REQUESTS_PHARMACY"]);
  }

  const refreshedForms = refreshLinkedFormsChoicesForWorkbook_(spreadsheet);
  Logger.log(`Form submit processed for ${spreadsheet.getName()} | emailAutofill:${emailAutofill} | deletedItems:${rowDeletions} | formsRefreshed:${refreshedForms}`);
  SpreadsheetApp.flush();
}

function reapplyModuleFormulasFromDashboard() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet();
  if (!dashboard) {
    throw new Error("Aucun classeur actif. Ouvrez le dashboard puis relancez reapplyModuleFormulasFromDashboard.");
  }

  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet) {
    throw new Error("CONFIG_SOURCES introuvable dans le dashboard actif.");
  }

  const sourceRows = sourceSheet.getLastRow() > 1
    ? sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues()
    : [];
  const audit = [["SourceKey", "Module", "SiteKey", "Status", "Detail", "WorkbookUrl"]];

  sourceRows.forEach((row) => {
    const sourceKey = String(row[0] || "").trim();
    const module = String(row[1] || "").trim().toLowerCase();
    const siteKey = String(row[2] || "").trim();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];

    if (!sourceKey || !module || !siteKey || !workbookUrl || !isTruthy_(enabled)) {
      return;
    }

    try {
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      let addedItems = 0;
      if (module === "incendie") {
        addedItems = syncFireItems_(workbook);
        applyFireFormulas_(workbook);
        protectSheets_(workbook, ["MOVEMENTS", "INVENTORY_COUNT", "STOCK_VIEW", "ALERTS", "PURCHASE_REQUESTS"]);
      } else if (module === "pharmacie") {
        addedItems = syncPharmaItems_(workbook);
        applyPharmaFormulas_(workbook);
        protectSheets_(workbook, ["MOVEMENTS_PHARMACY", "INVENTORY_COUNT_PHARMACY", "STOCK_VIEW_PHARMACY", "ALERTS_PHARMACY", "PURCHASE_REQUESTS_PHARMACY"]);
      } else {
        audit.push([sourceKey, module, siteKey, "SKIPPED", "Module non reconnu", workbookUrl]);
        return;
      }

      audit.push([sourceKey, module, siteKey, "OK", `Formules reappliquees | items ajoutes: ${addedItems}`, workbookUrl]);
    } catch (error) {
      audit.push([sourceKey, module, siteKey, "ERROR", String(error.message || error), workbookUrl]);
    }
  });

  const auditSheet = dashboard.getSheetByName("MODULE_FORMULA_AUDIT") || dashboard.insertSheet("MODULE_FORMULA_AUDIT");
  auditSheet.clearContents();
  auditSheet.clearFormats();
  auditSheet.getRange(1, 1, audit.length, audit[0].length).setValues(audit);
  auditSheet.getRange(1, 1, 1, audit[0].length).setFontWeight("bold");
  auditSheet.autoResizeColumns(1, audit[0].length);
  SpreadsheetApp.flush();

  Logger.log(`Module formulas audit generated (${audit.length - 1} rows).`);
  return audit;
}

function countActiveItems_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const values = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues();
  let count = 0;
  values.forEach((row) => {
    if (isTruthy_(row[0])) count += 1;
  });
  return count;
}

function countMatchesInColumn_(sheet, colIndex, matchValue) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const values = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).getDisplayValues();
  const target = String(matchValue || "").trim().toUpperCase();
  let count = 0;
  values.forEach((row) => {
    if (String(row[0] || "").trim().toUpperCase() === target) count += 1;
  });
  return count;
}

function countAnyMatchesInColumn_(sheet, colIndex, acceptedValues) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const accepted = {};
  (acceptedValues || []).forEach((value) => {
    accepted[String(value || "").trim().toUpperCase()] = true;
  });
  const values = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).getDisplayValues();
  let count = 0;
  values.forEach((row) => {
    const key = String(row[0] || "").trim().toUpperCase();
    if (accepted[key]) count += 1;
  });
  return count;
}

function getMaxDateInColumn_(sheet, colIndex) {
  if (!sheet || sheet.getLastRow() < 2) return "";
  const values = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).getValues();
  let maxValue = null;
  values.forEach((row) => {
    const value = row[0];
    if (!(value instanceof Date)) return;
    if (!maxValue || value.getTime() > maxValue.getTime()) maxValue = value;
  });
  return maxValue;
}

function buildItemActiveMap_(itemsSheet) {
  const map = {};
  if (!itemsSheet || itemsSheet.getLastRow() < 2) return map;
  const rows = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 8).getValues();
  rows.forEach((row) => {
    const key = normalizeTextKey_(row[0]);
    if (!key) return;
    map[key] = row[7];
  });
  return map;
}

function buildStockPresenceMap_(stockSheet) {
  const map = {};
  if (!stockSheet || stockSheet.getLastRow() < 2) return map;
  const values = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 1).getValues();
  values.forEach((row) => {
    const key = normalizeTextKey_(row[0]);
    if (key) map[key] = true;
  });
  return map;
}

function createFireModule_(folder, siteKey, now) {
  const workbook = createWorkbook_(folder, `OPS_INCENDIE_${siteKey}`, FIRE_TABS);
  seedFireData_(workbook.spreadsheet, siteKey, now);

  const forms = [];
  if (isNativeFormsMode_()) {
    buildFormLinkDefinitionsForModuleSite_("incendie", siteKey).forEach((definition) => {
      forms.push({ key: definition.key, label: definition.label, url: definition.url, id: "" });
    });
  } else {
    const movementForm = buildFireMovementForm_(folder, siteKey, workbook.spreadsheet);
    // Compatibilite legacy: la cle FORM_FIRE_INVENTORY_* est conservee,
    // mais ce formulaire correspond a une demande de reappro incendie.
    const replenishmentForm = buildFireInventoryForm_(folder, siteKey, workbook.spreadsheet);
    const newItemForm = buildFireItemCreationForm_(folder, siteKey, workbook.spreadsheet);

    attachFormToSpreadsheet_(movementForm.form, workbook.spreadsheet, "FIRE_FORM_MOVEMENTS_RAW");
    attachFormToSpreadsheet_(replenishmentForm.form, workbook.spreadsheet, rawSheetNameForFireReplenishment_());
    attachFormToSpreadsheet_(newItemForm.form, workbook.spreadsheet, "FIRE_FORM_MOVEMENTS_RAW");

    forms.push(
      { key: `FORM_FIRE_MOVEMENT_${siteKey}`, label: `Form mouvement incendie ${siteKey}`, url: movementForm.url, id: movementForm.id },
      { key: `FORM_FIRE_INVENTORY_${siteKey}`, label: `Form demande reappro incendie ${siteKey}`, url: replenishmentForm.url, id: replenishmentForm.id },
      { key: `FORM_FIRE_ITEM_CREATE_${siteKey}`, label: `Form creation article incendie ${siteKey}`, url: newItemForm.url, id: newItemForm.id },
    );
  }

  applyFireFormulas_(workbook.spreadsheet);
  protectSheets_(workbook.spreadsheet, ["MOVEMENTS", "INVENTORY_COUNT", "STOCK_VIEW", "ALERTS", "PURCHASE_REQUESTS"]);
  applyFileAccess_(workbook.file, DEPLOYMENT_CONFIG.adminEditors, []);

  return {
    workbookName: workbook.spreadsheet.getName(),
    workbookUrl: workbook.url,
    workbookId: workbook.id,
    forms,
  };
}

function createPharmaModule_(folder, siteKey, now) {
  const workbook = createWorkbook_(folder, `OPS_PHARMACIE_${siteKey}`, PHARMA_TABS);
  seedPharmaData_(workbook.spreadsheet, siteKey, now);

  const forms = [];
  if (isNativeFormsMode_()) {
    buildFormLinkDefinitionsForModuleSite_("pharmacie", siteKey).forEach((definition) => {
      forms.push({ key: definition.key, label: definition.label, url: definition.url, id: "" });
    });
  } else {
    const movementForm = buildPharmaMovementForm_(folder, siteKey, workbook.spreadsheet);
    const inventoryForm = buildPharmaInventoryForm_(folder, siteKey, workbook.spreadsheet);
    const newItemForm = buildPharmaItemCreationForm_(folder, siteKey, workbook.spreadsheet);

    attachFormToSpreadsheet_(movementForm.form, workbook.spreadsheet, "PHARMA_FORM_MOVEMENTS_RAW");
    attachFormToSpreadsheet_(inventoryForm.form, workbook.spreadsheet, "PHARMA_FORM_INVENTORY_RAW");
    attachFormToSpreadsheet_(newItemForm.form, workbook.spreadsheet, "PHARMA_FORM_MOVEMENTS_RAW");

    forms.push(
      { key: `FORM_PHARMA_MOVEMENT_${siteKey}`, label: `Form mouvement pharmacie ${siteKey}`, url: movementForm.url, id: movementForm.id },
      { key: `FORM_PHARMA_INVENTORY_${siteKey}`, label: `Form inventaire pharmacie ${siteKey}`, url: inventoryForm.url, id: inventoryForm.id },
      { key: `FORM_PHARMA_ITEM_CREATE_${siteKey}`, label: `Form creation article pharmacie ${siteKey}`, url: newItemForm.url, id: newItemForm.id },
    );
  }

  applyPharmaFormulas_(workbook.spreadsheet);
  protectSheets_(workbook.spreadsheet, [
    "MOVEMENTS_PHARMACY",
    "INVENTORY_COUNT_PHARMACY",
    "STOCK_VIEW_PHARMACY",
    "ALERTS_PHARMACY",
    "PURCHASE_REQUESTS_PHARMACY",
  ]);
  applyFileAccess_(workbook.file, DEPLOYMENT_CONFIG.adminEditors, []);

  return {
    workbookName: workbook.spreadsheet.getName(),
    workbookUrl: workbook.url,
    workbookId: workbook.id,
    forms,
  };
}

function createWorkbook_(folder, name, tabsMap) {
  const spreadsheet = SpreadsheetApp.create(name);
  spreadsheet.setSpreadsheetLocale(DEPLOYMENT_CONFIG.locale);
  spreadsheet.setSpreadsheetTimeZone(DEPLOYMENT_CONFIG.timezone);

  const file = getCachedFileById_(spreadsheet.getId());
  moveFileToFolder_(file, folder);

  const tabNames = Object.keys(tabsMap);
  const defaultSheet = spreadsheet.getSheets()[0];
  defaultSheet.setName(tabNames[0]);

  tabNames.forEach((tabName, index) => {
    const sheet = index === 0 ? defaultSheet : spreadsheet.insertSheet(tabName);
    initSheet_(sheet, tabsMap[tabName]);
  });

  return {
    spreadsheet,
    file,
    id: spreadsheet.getId(),
    url: spreadsheet.getUrl(),
  };
}

function initSheet_(sheet, headers) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function seedFireData_(spreadsheet, siteKey, now) {
  spreadsheet.getSheetByName("ITEMS").getRange(2, 1, 1, 11).setValues([
    ["INC-0001", "incendie", siteKey, "Extinguisher 6kg", "Extinguisher", "unit", 5, true, "SUP-001", "Zone A", now],
  ]);

  spreadsheet.getSheetByName("SUPPLIERS").getRange(2, 1, 1, 8).setValues([
    ["SUP-001", "Falcon Fire Supply", "Ops Contact", "contact@falconfire.example", "+33100000000", "Paris", true, now],
  ]);
}

function seedPharmaData_(spreadsheet, siteKey, now) {
  spreadsheet.getSheetByName("ITEMS_PHARMACY").getRange(2, 1, 1, 11).setValues([
    ["PHA-0001", "pharmacie", siteKey, "Paracetamol 500mg", "Analgesic", "box", 20, true, "SUP-P-001", "Shelf A", now],
  ]);

  spreadsheet.getSheetByName("SUPPLIERS_PHARMACY").getRange(2, 1, 1, 8).setValues([
    ["SUP-P-001", "Pharma Logistic", "Pharmacy Contact", "contact@pharmalog.example", "+33100000001", "Lyon", true, now],
  ]);
}

function getModuleItemContext_(module) {
  const normalized = String(module || "").trim().toLowerCase();
  if (normalized === "pharmacie") return { itemsTab: "ITEMS_PHARMACY", stockTab: "STOCK_VIEW_PHARMACY" };
  if (normalized === "incendie") return { itemsTab: "ITEMS", stockTab: "STOCK_VIEW" };
  return null;
}

function applyFireFormulas_(spreadsheet) {
  syncFireItems_(spreadsheet);

  const m = spreadsheet.getSheetByName("MOVEMENTS");
  const ic = spreadsheet.getSheetByName("INVENTORY_COUNT");
  const fireMovRaw = spreadsheet.getSheetByName("FIRE_FORM_MOVEMENTS_RAW");
  const fireInvRaw = spreadsheet.getSheetByName("FIRE_FORM_INVENTORY_RAW");
  if (!m || !ic || !fireMovRaw || !fireInvRaw) {
    throw new Error("MOVEMENTS / INVENTORY_COUNT / FIRE_FORM_*_RAW introuvable(s).");
  }

  const move = buildRawColumnResolver_(fireMovRaw);
  const moveTimestamp = move.byTokens(["horodateur", "timestamp"], 1);
  const moveSite = move.byTokens(["site concerne", "sitekey", "site"], 2);
  const moveItem = move.item(3);
  const moveType = move.byTokens(["type de mouvement", "movementtype", "movement type"], 4);
  const moveQty = move.byTokens(["quantite", "quantity"], 5);
  const moveUnitCost = move.byTokens(["cout unitaire", "unit cost", "unitcost"], 6);
  const moveReason = move.byTokens(["motif", "reason"], 7);
  const moveActorEmail = move.byTokens(["email operateur", "actoremail", "operator email"], 8);
  const moveDocRef = move.byTokens(["reference document", "document ref", "documentref"], 9);
  const moveComment = move.byTokens(["commentaire", "comment"], 10);

  m.getRange("A2").setFormula(`=ARRAYFORMULA(IF(${moveTimestamp}="";"";"MOV-INC-"&TEXT(ROW(${moveTimestamp})-1;"000000")))`);
  m.getRange("B2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveTimestamp}))`);
  m.getRange("C2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IFERROR(IF(COUNTIF(ITEMS!A:A;IFERROR(TRIM(REGEXEXTRACT(${moveItem};"^[^|]+"));${moveItem}))>0;IFERROR(TRIM(REGEXEXTRACT(${moveItem};"^[^|]+"));${moveItem});INDEX(ITEMS!A:A;MATCH(IFERROR(TRIM(REGEXEXTRACT(${moveItem};"^[^|]+"));${moveItem});ITEMS!D:D;0)));IFERROR(TRIM(REGEXEXTRACT(${moveItem};"^[^|]+"));${moveItem}))))`);
  m.getRange("D2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";"incendie"))');
  m.getRange("E2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveSite}))`);
  m.getRange("F2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveType}))`);
  m.getRange("G2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IF(${moveType}="OUT";-ABS(IFERROR(VALUE(${moveQty});0));IF(${moveType}="IN";ABS(IFERROR(VALUE(${moveQty});0));IFERROR(VALUE(${moveQty});0)))))`);
  m.getRange("H2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IFERROR(VALUE(${moveUnitCost});0)))`);
  m.getRange("I2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveReason}))`);
  m.getRange("J2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveActorEmail}))`);
  m.getRange("K2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveDocRef}))`);
  m.getRange("L2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveComment}))`);

  const inv = buildRawColumnResolver_(fireInvRaw);
  const invTimestamp = inv.byTokens(["horodateur", "timestamp"], 1);
  const invSite = inv.byTokens(["site concerne", "sitekey", "site"], 2);
  const invItem = inv.item(3);
  const invRequestedQty = inv.byTokens(["quantite demandee", "quantite souhaitee", "quantite comptee", "countedqty", "counted qty", "quantite"], 4);
  const invPriority = inv.optionalByTokens(["priorite", "priority"]);
  const invReason = inv.optionalByTokens(["motif de la demande", "motif", "raison", "reason"]);
  const invComment = inv.optionalByTokens(["commentaire", "comment"]);
  const invRequesterEmail = inv.optionalByTokens(["email demandeur", "email controleur", "counteremail", "requester email"]);
  const invSubmitterEmail = inv.optionalByTokens(["adresse e-mail", "email address", "e-mail", "email"]);
  const invPriorityExpr = invPriority ? `UPPER(TRIM(${invPriority}))` : '""';
  const invRequesterExpr = invRequesterEmail ? `TRIM(${invRequesterEmail})` : '""';
  const invSubmitterExpr = invSubmitterEmail ? `TRIM(${invSubmitterEmail})` : '""';
  const invReasonExpr = invReason ? `TRIM(${invReason})` : '""';
  const invCommentExpr = invComment ? `TRIM(${invComment})` : '""';

  // Journal de demandes de reappro (audit interne module incendie).
  ic.getRange("A2").setFormula(`=ARRAYFORMULA(IF(${invTimestamp}="";"";IF(ABS(IFERROR(VALUE(${invRequestedQty});0))<=0;"";"DEM-INC-"&TEXT(ROW(${invTimestamp})-1;"000000"))))`);
  ic.getRange("B2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${invTimestamp}))`);
  ic.getRange("C2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IFERROR(IF(COUNTIF(ITEMS!A:A;IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem}))>0;IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem});INDEX(ITEMS!A:A;MATCH(IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem});ITEMS!D:D;0)));IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem}))))`);
  ic.getRange("D2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";"incendie"))');
  ic.getRange("E2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${invSite}))`);
  ic.getRange("F2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";IFERROR(VLOOKUP(C2:C;STOCK_VIEW!A:F;6;FALSE);0)))');
  ic.getRange("G2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";ABS(IFERROR(VALUE(${invRequestedQty});0))))`);
  ic.getRange("H2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";G2:G))');
  ic.getRange("I2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IF(LEN(${invRequesterExpr})>0;${invRequesterExpr};${invSubmitterExpr})))`);
  ic.getRange("J2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";""))');
  ic.getRange("K2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";"A_VALIDER"))');
  ic.getRange("L2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IF(LEN(${invCommentExpr})>0;${invCommentExpr};${invReasonExpr})))`);

  applyStockViewFormulas_(spreadsheet.getSheetByName("STOCK_VIEW"), "ITEMS", "MOVEMENTS", "incendie", null);
  applyAlertAndPurchaseFormulas_(spreadsheet.getSheetByName("ALERTS"), spreadsheet.getSheetByName("PURCHASE_REQUESTS"), "STOCK_VIEW");

  const purchase = spreadsheet.getSheetByName("PURCHASE_REQUESTS");
  if (purchase) {
    purchase.getRange("A2").setFormula(`=ARRAYFORMULA(IF(${invTimestamp}="";"";IF(ABS(IFERROR(VALUE(${invRequestedQty});0))<=0;"";"REQ-INC-FRM-"&TEXT(ROW(${invTimestamp})-1;"000000"))))`);
    purchase.getRange("B2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";"incendie"))');
    purchase.getRange("C2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${invSite}))`);
    purchase.getRange("D2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IFERROR(IF(COUNTIF(ITEMS!A:A;IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem}))>0;IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem});INDEX(ITEMS!A:A;MATCH(IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem});ITEMS!D:D;0)));IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem}))))`);
    purchase.getRange("E2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IFERROR(VLOOKUP(D2:D;ITEMS!A:D;4;FALSE);IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+\\|\\s*([^|]+)"));D2:D))))`);
    purchase.getRange("F2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";IFERROR(VLOOKUP(D2:D;ITEMS!A:I;9;FALSE);"")))');
    purchase.getRange("G2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";ABS(IFERROR(VALUE(${invRequestedQty});0))))`);
    purchase.getRange("H2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IF(REGEXMATCH(${invPriorityExpr};"^(HIGH|MEDIUM|LOW)$");${invPriorityExpr};"MEDIUM")))`);
    purchase.getRange("I2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";"A_VALIDER"))');
    purchase.getRange("J2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IF(LEN(${invRequesterExpr})>0;${invRequesterExpr};${invSubmitterExpr})))`);
    purchase.getRange("K2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${invTimestamp}))`);
  }

  const items = spreadsheet.getSheetByName("ITEMS");
  if (items) {
    items.getRange("L1").setValue("CurrentQty").setFontWeight("bold");
    items.getRange("L2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";IFERROR(VLOOKUP(A2:A;STOCK_VIEW!A:F;6;FALSE);0)))');
  }
}

function applyPharmaFormulas_(spreadsheet) {
  syncPharmaItems_(spreadsheet);

  const m = spreadsheet.getSheetByName("MOVEMENTS_PHARMACY");
  const ic = spreadsheet.getSheetByName("INVENTORY_COUNT_PHARMACY");
  const pharmaMovRaw = spreadsheet.getSheetByName("PHARMA_FORM_MOVEMENTS_RAW");
  const pharmaInvRaw = spreadsheet.getSheetByName("PHARMA_FORM_INVENTORY_RAW");
  if (!m || !ic || !pharmaMovRaw || !pharmaInvRaw) {
    throw new Error("MOVEMENTS_PHARMACY / INVENTORY_COUNT_PHARMACY / PHARMA_FORM_*_RAW introuvable(s).");
  }

  const move = buildRawColumnResolver_(pharmaMovRaw);
  const moveTimestamp = move.byTokens(["horodateur", "timestamp"], 1);
  const moveSite = move.byTokens(["site concerne", "sitekey", "site"], 2);
  const moveItem = move.item(3);
  const moveType = move.byTokens(["type de mouvement", "movementtype", "movement type"], 4);
  const moveQty = move.byTokens(["quantite", "quantity"], 5);
  const moveUnitCost = move.byTokens(["cout unitaire", "unit cost", "unitcost"], 6);
  const moveReason = move.byTokens(["motif", "reason"], 7);
  const moveActorEmail = move.byTokens(["email operateur", "actoremail", "operator email"], 8);
  const moveLot = move.byTokens(["numero de lot", "lotnumber", "lot number"], 9);
  const moveExpiry = move.byTokens(["date de peremption", "expirydate", "expiry date", "peremption"], 10);
  const moveDocRef = move.byTokens(["reference document", "document ref", "documentref"], 11);
  const moveComment = move.byTokens(["commentaire", "comment"], 12);

  m.getRange("A2").setFormula(`=ARRAYFORMULA(IF(${moveTimestamp}="";"";"MOV-PHA-"&TEXT(ROW(${moveTimestamp})-1;"000000")))`);
  m.getRange("B2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveTimestamp}))`);
  m.getRange("C2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IFERROR(IF(COUNTIF(ITEMS_PHARMACY!A:A;IFERROR(TRIM(REGEXEXTRACT(${moveItem};"^[^|]+"));${moveItem}))>0;IFERROR(TRIM(REGEXEXTRACT(${moveItem};"^[^|]+"));${moveItem});INDEX(ITEMS_PHARMACY!A:A;MATCH(IFERROR(TRIM(REGEXEXTRACT(${moveItem};"^[^|]+"));${moveItem});ITEMS_PHARMACY!D:D;0)));IFERROR(TRIM(REGEXEXTRACT(${moveItem};"^[^|]+"));${moveItem}))))`);
  m.getRange("D2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";"pharmacie"))');
  m.getRange("E2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveSite}))`);
  m.getRange("F2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveType}))`);
  m.getRange("G2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IF(${moveType}="OUT";-ABS(IFERROR(VALUE(${moveQty});0));IF(${moveType}="IN";ABS(IFERROR(VALUE(${moveQty});0));IFERROR(VALUE(${moveQty});0)))))`);
  m.getRange("H2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IFERROR(VALUE(${moveUnitCost});0)))`);
  m.getRange("I2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveReason}))`);
  m.getRange("J2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveActorEmail}))`);
  m.getRange("K2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveDocRef}))`);
  m.getRange("L2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveComment}))`);
  m.getRange("M2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveLot}))`);
  m.getRange("N2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${moveExpiry}))`);

  const inv = buildRawColumnResolver_(pharmaInvRaw);
  const invTimestamp = inv.byTokens(["horodateur", "timestamp"], 1);
  const invSite = inv.byTokens(["site concerne", "sitekey", "site"], 2);
  const invItem = inv.item(3);
  const invCountedQty = inv.byTokens(["quantite comptee", "countedqty", "counted qty", "quantite"], 4);
  const invCounterEmail = inv.byTokens(["email controleur", "counteremail", "controller email"], 5);
  const invComment = inv.byTokens(["commentaire", "comment"], 6);

  ic.getRange("A2").setFormula(`=ARRAYFORMULA(IF(${invTimestamp}="";"";"CNT-PHA-"&TEXT(ROW(${invTimestamp})-1;"000000")))`);
  ic.getRange("B2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${invTimestamp}))`);
  ic.getRange("C2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IFERROR(IF(COUNTIF(ITEMS_PHARMACY!A:A;IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem}))>0;IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem});INDEX(ITEMS_PHARMACY!A:A;MATCH(IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem});ITEMS_PHARMACY!D:D;0)));IFERROR(TRIM(REGEXEXTRACT(${invItem};"^[^|]+"));${invItem}))))`);
  ic.getRange("D2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";"pharmacie"))');
  ic.getRange("E2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${invSite}))`);
  ic.getRange("F2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";IFERROR(SUMIFS(MOVEMENTS_PHARMACY!$G:$G;MOVEMENTS_PHARMACY!$C:$C;C2:C);0)))');
  ic.getRange("G2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";IFERROR(VALUE(${invCountedQty});0)))`);
  ic.getRange("H2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";G2:G-F2:F))');
  ic.getRange("I2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${invCounterEmail}))`);
  ic.getRange("J2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";""))');
  ic.getRange("K2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";"A_VALIDER"))');
  ic.getRange("L2").setFormula(`=ARRAYFORMULA(IF(A2:A="";"";${invComment}))`);

  applyStockViewFormulas_(spreadsheet.getSheetByName("STOCK_VIEW_PHARMACY"), "ITEMS_PHARMACY", "MOVEMENTS_PHARMACY", "pharmacie", "INVENTORY_COUNT_PHARMACY");

  const lots = spreadsheet.getSheetByName("LOTS_PHARMACY");
  lots.getRange("H2").setFormula('=IF(F2<=0;"CLOSED";IF(E2<TODAY();"EXPIRED";IF(E2<=TODAY()+30;"ALERT_30D";IF(E2<=TODAY()+90;"ALERT_90D";"OK"))))');
  lots.getRange("H2").copyTo(lots.getRange("H2:H5000"));

  const alerts = spreadsheet.getSheetByName("ALERTS_PHARMACY");
  alerts.getRange("A2").setFormula('=ARRAYFORMULA(IFERROR(FILTER(HSTACK("ALT-PHA-"&ROW(LOTS_PHARMACY!A2:A);IF(LOTS_PHARMACY!A2:A<>"";"pharmacie";"");LOTS_PHARMACY!C2:C;LOTS_PHARMACY!H2:H;IF(LOTS_PHARMACY!H2:H="EXPIRED";"HIGH";"MEDIUM");LOTS_PHARMACY!B2:B;IFERROR(VLOOKUP(LOTS_PHARMACY!B2:B;ITEMS_PHARMACY!A:D;4;FALSE);"");"Lot "&LOTS_PHARMACY!D2:D&" -> "&LOTS_PHARMACY!H2:H;IF(LOTS_PHARMACY!A2:A<>"";TODAY();"");IF(LOTS_PHARMACY!A2:A<>"";"OPEN";"");IF(LOTS_PHARMACY!A2:A<>"";"";""));(LOTS_PHARMACY!H2:H="ALERT_90D") + (LOTS_PHARMACY!H2:H="ALERT_30D") + (LOTS_PHARMACY!H2:H="EXPIRED"));""))');

  const purchase = spreadsheet.getSheetByName("PURCHASE_REQUESTS_PHARMACY");
  purchase.getRange("A2").setFormula('=ARRAYFORMULA(IFERROR(FILTER(HSTACK("REQ-"&ROW(STOCK_VIEW_PHARMACY!A2:A);STOCK_VIEW_PHARMACY!K2:K;STOCK_VIEW_PHARMACY!L2:L;STOCK_VIEW_PHARMACY!A2:A;STOCK_VIEW_PHARMACY!B2:B;STOCK_VIEW_PHARMACY!J2:J;ABS(STOCK_VIEW_PHARMACY!G2:G);IF(STOCK_VIEW_PHARMACY!H2:H="RUPTURE";"HIGH";"MEDIUM");IF(STOCK_VIEW_PHARMACY!A2:A<>"";"A_VALIDER";"");IF(STOCK_VIEW_PHARMACY!A2:A<>"";"";"");IF(STOCK_VIEW_PHARMACY!A2:A<>"";TODAY();""));(STOCK_VIEW_PHARMACY!H2:H="SOUS_SEUIL") + (STOCK_VIEW_PHARMACY!H2:H="RUPTURE"));""))');

  const items = spreadsheet.getSheetByName("ITEMS_PHARMACY");
  if (items) {
    items.getRange("L1").setValue("CurrentQty").setFontWeight("bold");
    items.getRange("L2").setFormula('=ARRAYFORMULA(IF(A2:A="";"";IFERROR(VLOOKUP(A2:A;STOCK_VIEW_PHARMACY!A:F;6;FALSE);0)))');
  }
}

function applyStockViewFormulas_(sheet, itemsTab, movementsTab, moduleLabel, inventoryTab) {
  sheet.getRange("A2").setFormula(`=FILTER(${itemsTab}!A2:A;${itemsTab}!A2:A<>"";${itemsTab}!H2:H=TRUE)`);
  sheet.getRange("B2").setFormula(`=IF(A2="";"";INDEX(${itemsTab}!D:D;MATCH(A2;${itemsTab}!A:A;0)))`);
  sheet.getRange("C2").setFormula(`=IF(A2="";"";INDEX(${itemsTab}!E:E;MATCH(A2;${itemsTab}!A:A;0)))`);
  sheet.getRange("D2").setFormula(`=IF(A2="";"";INDEX(${itemsTab}!F:F;MATCH(A2;${itemsTab}!A:A;0)))`);
  sheet.getRange("E2").setFormula(`=IF(A2="";"";INDEX(${itemsTab}!G:G;MATCH(A2;${itemsTab}!A:A;0)))`);
  if (inventoryTab) {
    sheet.getRange("F2").setFormula(`=IF(A2="";"";IFERROR(LOOKUP(2;1/(${inventoryTab}!$C:$C=$A2);${inventoryTab}!$G:$G)+IFERROR(SUMIFS(${movementsTab}!$G:$G;${movementsTab}!$C:$C;$A2;${movementsTab}!$B:$B;">"&LOOKUP(2;1/(${inventoryTab}!$C:$C=$A2);${inventoryTab}!$B:$B));0);IFERROR(SUMIFS(${movementsTab}!$G:$G;${movementsTab}!$C:$C;$A2);0)))`);
  } else {
    sheet.getRange("F2").setFormula(`=IF(A2="";"";IFERROR(SUMIFS(${movementsTab}!$G:$G;${movementsTab}!$C:$C;$A2);0))`);
  }
  sheet.getRange("G2").setFormula('=IF(A2="";"";E2-F2)');
  sheet.getRange("H2").setFormula('=IF(A2="";"";IF(F2<=0;"RUPTURE";IF(F2<E2;"SOUS_SEUIL";"OK")))');
  sheet.getRange("I2").setFormula(`=IF(A2="";"";IFERROR(MAX(FILTER(${movementsTab}!$B:$B;${movementsTab}!$C:$C=$A2));""))`);
  sheet.getRange("J2").setFormula(`=IF(A2="";"";INDEX(${itemsTab}!I:I;MATCH(A2;${itemsTab}!A:A;0)))`);
  sheet.getRange("K2").setFormula(`=IF(A2="";"";"${moduleLabel}")`);
  sheet.getRange("L2").setFormula(`=IF(A2="";"";INDEX(${itemsTab}!C:C;MATCH(A2;${itemsTab}!A:A;0)))`);
  sheet.getRange("B2:L2").copyTo(sheet.getRange("B2:L5000"));
}

function applyAlertAndPurchaseFormulas_(alertsSheet, purchaseSheet, stockViewTab) {
  if (!alertsSheet || !purchaseSheet) {
    throw new Error("ALERTS et/ou PURCHASE_REQUESTS introuvable(s). Verifiez la creation des onglets du module.");
  }

  alertsSheet.getRange("A2").setFormula(`=ARRAYFORMULA(IFERROR(FILTER(HSTACK("ALT-"&ROW(${stockViewTab}!A2:A);${stockViewTab}!K2:K;${stockViewTab}!L2:L;${stockViewTab}!H2:H;IF(${stockViewTab}!H2:H="RUPTURE";"HIGH";"MEDIUM");${stockViewTab}!A2:A;${stockViewTab}!B2:B;"Stock status: "&${stockViewTab}!H2:H;${stockViewTab}!I2:I;IF(${stockViewTab}!A2:A<>"";"OPEN";"");IF(${stockViewTab}!A2:A<>"";"";""));(${stockViewTab}!H2:H="SOUS_SEUIL") + (${stockViewTab}!H2:H="RUPTURE"));""))`);
  purchaseSheet.getRange("A2").setFormula(`=ARRAYFORMULA(IFERROR(FILTER(HSTACK("REQ-"&ROW(${stockViewTab}!A2:A);${stockViewTab}!K2:K;${stockViewTab}!L2:L;${stockViewTab}!A2:A;${stockViewTab}!B2:B;${stockViewTab}!J2:J;ABS(${stockViewTab}!G2:G);IF(${stockViewTab}!H2:H="RUPTURE";"HIGH";"MEDIUM");IF(${stockViewTab}!A2:A<>"";"A_VALIDER";"");IF(${stockViewTab}!A2:A<>"";"";"");IF(${stockViewTab}!A2:A<>"";TODAY();""));(${stockViewTab}!H2:H="SOUS_SEUIL") + (${stockViewTab}!H2:H="RUPTURE"));""))`);
}
