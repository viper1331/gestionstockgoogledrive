/*
 * Bloc 2 - Normalisation RAW / ingestion
 * Validation technique, mapping headers/valeurs, append RAW et normalisation.
 */

function buildNativeSubmissionPayload_(meta, payload, actorEmail, access) {
  const now = new Date();
  const type = meta.type;
  const siteKey = String(payload.siteKey || "").trim();
  const itemValue = String(payload.itemValue || "").trim();
  const quantity = parseNumberOrDefault_(payload.quantity, 0);
  const reason = String(payload.reason || "").trim();
  const documentRef = String(payload.documentRef || "").trim();
  const comment = String(payload.comment || "").trim();
  const category = String(payload.category || "").trim();
  const unit = String(payload.unit || "").trim();
  const minThreshold = parseNumberOrDefault_(payload.minThreshold, 0);
  const deleteItemsValues = normalizeStringArray_(payload.deleteItems);
  const deleteItems = deleteItemsValues.join(", ");
  const lotNumber = String(payload.lotNumber || "").trim();
  const expiryDate = String(payload.expiryDate || "").trim();

  if (hasNativeDeleteIntentPayload_(payload) && !canUseNativeDestructiveActions_(access)) {
    throw new Error("Action reservee aux profils ADMIN / CONTROLEUR.");
  }

  if (!siteKey) {
    throw new Error("Le site est obligatoire.");
  }

  if (type === NATIVE_FORM_TYPES.FIRE_MOVEMENT || type === NATIVE_FORM_TYPES.PHARMA_MOVEMENT) {
    if (!itemValue) throw new Error("Article concerne obligatoire.");
    const movementType = String(payload.movementType || "IN").trim().toUpperCase();
    if (["IN", "OUT", "CORRECTION", "DELETE_ITEM"].indexOf(movementType) === -1) {
      throw new Error("Type de mouvement invalide.");
    }
    if (!reason) throw new Error("Motif obligatoire.");
    const operatorEmail = normalizeEmail_(payload.operatorEmail) || actorEmail;
    const requiredHeaders = [
      "Horodateur",
      "Site concerne",
      "Article concerne",
      "Type de mouvement",
      "Quantite",
      "Cout unitaire",
      "Motif",
      "Email operateur",
      "Reference document",
      "Commentaire",
      "Supprimer des articles (optionnel)",
      "Categorie article (optionnel)",
      "Unite article (optionnel)",
      "Seuil mini article (optionnel)",
    ];
    if (type === NATIVE_FORM_TYPES.PHARMA_MOVEMENT) {
      requiredHeaders.push("Numero de lot");
      requiredHeaders.push("Date de peremption");
    }
    const valuesByHeader = {
      Horodateur: now,
      "Site concerne": siteKey,
      "Article concerne": itemValue,
      "Type de mouvement": movementType,
      Quantite: quantity,
      "Cout unitaire": parseNumberOrDefault_(payload.unitCost, 0),
      Motif: reason,
      "Email operateur": operatorEmail,
      "Reference document": documentRef,
      Commentaire: comment,
      "Supprimer des articles (optionnel)": deleteItems,
      "Categorie article (optionnel)": category,
      "Unite article (optionnel)": unit,
      "Seuil mini article (optionnel)": minThreshold,
    };
    if (type === NATIVE_FORM_TYPES.PHARMA_MOVEMENT) {
      valuesByHeader["Numero de lot"] = lotNumber;
      valuesByHeader["Date de peremption"] = expiryDate;
    }
    return {
      rawSheetName: type === NATIVE_FORM_TYPES.FIRE_MOVEMENT ? "FIRE_FORM_MOVEMENTS_RAW" : "PHARMA_FORM_MOVEMENTS_RAW",
      requiredHeaders,
      valuesByHeader,
    };
  }

  if (type === NATIVE_FORM_TYPES.FIRE_REPLENISHMENT) {
    if (!itemValue) throw new Error("Article a reapprovisionner obligatoire.");
    if (quantity <= 0) throw new Error("Quantite demandee doit etre > 0.");
    const priorityRaw = String(payload.priority || "MEDIUM").trim().toUpperCase();
    const priority = ["HIGH", "MEDIUM", "LOW"].indexOf(priorityRaw) >= 0 ? priorityRaw : "MEDIUM";
    const requesterEmail = normalizeEmail_(payload.requesterEmail) || actorEmail;
    if (!reason) throw new Error("Motif de la demande obligatoire.");
    return {
      rawSheetName: "FIRE_FORM_INVENTORY_RAW",
      requiredHeaders: [
        "Horodateur",
        "Site concerne",
        "Article a reapprovisionner",
        "Quantite demandee",
        "Priorite",
        "Motif de la demande",
        "Email demandeur",
        "Commentaire",
      ],
      valuesByHeader: {
        Horodateur: now,
        "Site concerne": siteKey,
        "Article a reapprovisionner": itemValue,
        "Quantite demandee": quantity,
        Priorite: priority,
        "Motif de la demande": reason,
        "Email demandeur": requesterEmail,
        Commentaire: comment,
      },
    };
  }

  if (type === NATIVE_FORM_TYPES.PHARMA_INVENTORY) {
    if (!itemValue) throw new Error("Article concerne obligatoire.");
    const counterEmail = normalizeEmail_(payload.counterEmail) || actorEmail;
    return {
      rawSheetName: "PHARMA_FORM_INVENTORY_RAW",
      requiredHeaders: [
        "Horodateur",
        "Site concerne",
        "Article concerne",
        "Quantite comptee",
        "Email controleur",
        "Commentaire",
        "Supprimer des articles (optionnel)",
        "Categorie article (optionnel)",
        "Unite article (optionnel)",
        "Seuil mini article (optionnel)",
      ],
      valuesByHeader: {
        Horodateur: now,
        "Site concerne": siteKey,
        "Article concerne": itemValue,
        "Quantite comptee": quantity,
        "Email controleur": counterEmail,
        Commentaire: comment,
        "Supprimer des articles (optionnel)": deleteItems,
        "Categorie article (optionnel)": category,
        "Unite article (optionnel)": unit,
        "Seuil mini article (optionnel)": minThreshold,
      },
    };
  }

  if (type === NATIVE_FORM_TYPES.FIRE_ITEM_CREATE || type === NATIVE_FORM_TYPES.PHARMA_ITEM_CREATE) {
    const autoItemId = buildNextAutoItemId_(payload && payload.sourceWorkbook ? payload.sourceWorkbook : null, meta.module);
    const itemId = autoItemId;
    const itemName = String(payload.itemName || "").trim();
    if (!itemId) throw new Error("Impossible de generer automatiquement un ItemID.");
    if (!itemName) throw new Error("Nom article obligatoire.");
    const operatorEmail = normalizeEmail_(payload.operatorEmail) || actorEmail;
    const effectiveReason = reason || "Creation article";
    const articleLabel = `${itemId} | ${itemName}`;
    const requiredHeaders = [
      "Horodateur",
      "Site concerne",
      "ItemID",
      "Article concerne",
      "Nom article",
      "Categorie article (optionnel)",
      "Unite article (optionnel)",
      "Seuil mini article (optionnel)",
      "Quantite",
      "Type de mouvement",
      "Motif",
      "Email operateur",
      "Reference document",
      "Commentaire",
    ];
    if (type === NATIVE_FORM_TYPES.PHARMA_ITEM_CREATE) {
      requiredHeaders.push("Numero de lot");
      requiredHeaders.push("Date de peremption");
    }
    const valuesByHeader = {
      Horodateur: now,
      "Site concerne": siteKey,
      ItemID: itemId,
      "Article concerne": articleLabel,
      "Nom article": itemName,
      "Categorie article (optionnel)": category,
      "Unite article (optionnel)": unit,
      "Seuil mini article (optionnel)": minThreshold,
      Quantite: quantity,
      "Type de mouvement": "IN",
      Motif: effectiveReason,
      "Email operateur": operatorEmail,
      "Reference document": documentRef,
      Commentaire: comment,
    };
    if (type === NATIVE_FORM_TYPES.PHARMA_ITEM_CREATE) {
      valuesByHeader["Numero de lot"] = lotNumber;
      valuesByHeader["Date de peremption"] = expiryDate;
    }
    return {
      rawSheetName: type === NATIVE_FORM_TYPES.FIRE_ITEM_CREATE ? "FIRE_FORM_MOVEMENTS_RAW" : "PHARMA_FORM_MOVEMENTS_RAW",
      requiredHeaders,
      valuesByHeader,
      generatedItemId: itemId,
    };
  }

  throw new Error(`Soumission native non supportee: ${type}`);
}

function itemIdPrefixForModule_(module) {
  const normalizedModule = normalizeModuleKey_(module);
  if (normalizedModule === "pharmacie") return "PHA-";
  return "INC-";
}

function extractNumericPartFromItemId_(itemId, expectedPrefix) {
  const value = String(itemId || "").trim().toUpperCase();
  if (!value) return null;
  const prefix = String(expectedPrefix || "").trim().toUpperCase().replace(/[-_\s]+$/, "");
  const prefixPattern = prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = value.match(new RegExp(`^${prefixPattern}[-_\\s]?(\\d+)$`, "i"));
  if (!match || !match[1]) return null;
  const parsed = Number(match[1]);
  return Number.isFinite(parsed) ? parsed : null;
}

function buildNextAutoItemId_(workbook, module) {
  const context = getModuleItemContext_(module);
  if (!workbook || !context || !context.itemsTab) return "";
  const itemsSheet = workbook.getSheetByName(context.itemsTab);
  if (!itemsSheet) return "";

  const prefix = itemIdPrefixForModule_(module);
  let maxId = 0;
  if (itemsSheet.getLastRow() >= 2) {
    const values = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 1).getValues();
    values.forEach((row) => {
      const numeric = extractNumericPartFromItemId_(row[0], prefix);
      if (numeric !== null && numeric > maxId) maxId = numeric;
    });
  }
  const next = maxId + 1;
  return `${prefix}${String(next).padStart(4, "0")}`;
}

function ensureRawHeaders_(sheet, requiredHeaders) {
  if (!sheet) throw new Error("Onglet RAW introuvable.");
  const normalizedRequired = (requiredHeaders || [])
    .map((header) => String(header || "").trim())
    .filter((header) => header !== "");
  if (!normalizedRequired.length) return;

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const rawHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let headerCount = 0;
  for (let i = rawHeaders.length - 1; i >= 0; i -= 1) {
    if (String(rawHeaders[i] || "").trim() !== "") {
      headerCount = i + 1;
      break;
    }
  }
  if (headerCount === 0) {
    sheet.getRange(1, 1, 1, normalizedRequired.length).setValues([normalizedRequired]).setFontWeight("bold");
    return;
  }

  const existingHeaders = rawHeaders.slice(0, headerCount);
  const existingMap = {};
  existingHeaders.forEach((header) => {
    existingMap[normalizeTextKey_(header)] = true;
  });
  const missing = normalizedRequired.filter((header) => !existingMap[normalizeTextKey_(header)]);
  if (!missing.length) return;

  const startCol = headerCount + 1;
  sheet.getRange(1, startCol, 1, missing.length).setValues([missing]).setFontWeight("bold");
}

function appendNativeRow_(workbook, rawSheetName, requiredHeaders, valuesByHeader, actorEmail) {
  const rawSheet = workbook ? workbook.getSheetByName(rawSheetName) : null;
  if (!rawSheet) throw new Error(`Onglet RAW introuvable: ${rawSheetName}`);
  ensureRawHeaders_(rawSheet, requiredHeaders);

  const lastColumn = Math.max(rawSheet.getLastColumn(), 1);
  const headers = rawSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    const key = normalizeTextKey_(header);
    if (!key) return;
    headerMap[key] = index;
  });

  const row = new Array(headers.length).fill("");
  Object.keys(valuesByHeader || {}).forEach((header) => {
    const index = headerMap[normalizeTextKey_(header)];
    if (index === undefined) return;
    row[index] = valuesByHeader[header];
  });

  const actor = normalizeEmail_(actorEmail);
  if (actor) {
    const emailIdx = findHeaderIndex_(headers.map((header) => normalizeTextKey_(header)), ["adresse e-mail", "email address", "email"]);
    if (emailIdx >= 0 && String(row[emailIdx] || "").trim() === "") {
      row[emailIdx] = actor;
    }
  }

  const rowIndex = rawSheet.getLastRow() + 1;
  rawSheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  return rowIndex;
}

function getRawSheetNamesForModule_(module) {
  const normalized = String(module || "").trim().toLowerCase();
  if (normalized === "incendie") {
    return ["FIRE_FORM_MOVEMENTS_RAW", "FIRE_FORM_INVENTORY_RAW"];
  }
  if (normalized === "pharmacie") {
    return ["PHARMA_FORM_MOVEMENTS_RAW", "PHARMA_FORM_INVENTORY_RAW"];
  }
  return [];
}

function recoverRawSheetFromBackupIfNeeded_(spreadsheet, rawSheetName) {
  if (!spreadsheet || !rawSheetName) return 0;
  const rawSheet = spreadsheet.getSheetByName(rawSheetName);
  if (!rawSheet) return 0;
  if (rawSheet.getLastRow() > 1) return 0;

  const prefix = `${rawSheetName}_BACKUP_`;
  const backups = spreadsheet.getSheets()
    .filter((sheet) => {
      const name = String(sheet.getName() || "");
      return name.indexOf(prefix) === 0 && sheet.getLastRow() > 1;
    })
    .sort((a, b) => String(b.getName() || "").localeCompare(String(a.getName() || "")));
  if (!backups.length) return 0;

  appendSheetData_(backups[0], rawSheet);
  return Math.max(rawSheet.getLastRow() - 1, 0);
}

function recoverRawDataFromBackupsFromDashboard_() {
  const dashboard = resolveAdminDashboard_("Recuperation RAW depuis backup");
  const sourceSheet = dashboard.getSheetByName("CONFIG_SOURCES");
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    return { checkedSources: 0, recoveredSheets: 0, recoveredRows: 0 };
  }

  const sourceRows = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 9).getValues();
  let checkedSources = 0;
  let recoveredSheets = 0;
  let recoveredRows = 0;

  sourceRows.forEach((row) => {
    const module = String(row[1] || "").trim().toLowerCase();
    const workbookUrl = String(row[3] || "").trim();
    const enabled = row[8];
    if (!module || !workbookUrl || !isTruthy_(enabled)) return;

    const rawNames = getRawSheetNamesForModule_(module);
    if (!rawNames.length) return;
    checkedSources += 1;

    try {
      const workbook = getCachedSpreadsheetByUrl_(workbookUrl);
      rawNames.forEach((rawSheetName) => {
        const restored = recoverRawSheetFromBackupIfNeeded_(workbook, rawSheetName);
        if (restored <= 0) return;
        recoveredSheets += 1;
        recoveredRows += restored;
      });
    } catch (error) {
      Logger.log(`RAW recovery warning (${workbookUrl}): ${String(error.message || error)}`);
    }
  });

  return { checkedSources, recoveredSheets, recoveredRows };
}

function autofillEmailColumnsFromSubmission_(e) {
  if (!e || !e.range || !e.source) return 0;
  const rawSheet = e.range.getSheet();
  const rowIndex = e.range.getRow();
  if (!rawSheet || rowIndex < 2) return 0;

  const headers = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0]
    .map((value) => normalizeTextKey_(value));
  const row = rawSheet.getRange(rowIndex, 1, 1, rawSheet.getLastColumn()).getValues()[0];
  const respondentEmail = getRespondentEmailFromSubmission_(e) || getCurrentUserEmail_();
  if (!respondentEmail) return 0;

  const emailTargets = [
    findHeaderIndex_(headers, ["email controleur", "counteremail"]),
    findHeaderIndex_(headers, ["email operateur", "actoremail"]),
    findHeaderIndex_(headers, ["email demandeur", "requester email", "demandeur"]),
  ];

  let changed = 0;
  emailTargets.forEach((index) => {
    if (index < 0) return;
    if (String(row[index] || "").trim() !== "") return;
    row[index] = respondentEmail;
    changed += 1;
  });

  if (changed > 0) {
    rawSheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  }
  return changed;
}

function getRespondentEmailFromSubmission_(e) {
  if (!e || !e.namedValues) return "";
  const keys = Object.keys(e.namedValues);
  for (let i = 0; i < keys.length; i += 1) {
    const key = normalizeTextKey_(keys[i]);
    if (key.indexOf("adresse e-mail") === -1 && key.indexOf("email address") === -1 && key !== "email") {
      continue;
    }
    const value = e.namedValues[keys[i]];
    if (value && value.length) {
      const email = normalizeEmail_(value[0]);
      if (email) return email;
    }
  }
  return "";
}

function itemsSheetNameForRawSheet_(rawSheetName) {
  const name = String(rawSheetName || "").trim();
  if (name === "FIRE_FORM_MOVEMENTS_RAW" || name === "FIRE_FORM_INVENTORY_RAW") return "ITEMS";
  if (name === "PHARMA_FORM_MOVEMENTS_RAW" || name === "PHARMA_FORM_INVENTORY_RAW") return "ITEMS_PHARMACY";
  return "";
}

function applyItemDeletionsFromSubmission_(e) {
  if (!e || !e.range || !e.source) return 0;

  const rawSheet = e.range.getSheet();
  const rowIndex = e.range.getRow();
  if (!rawSheet || rowIndex < 2) return 0;

  const itemsSheetName = itemsSheetNameForRawSheet_(rawSheet.getName());
  if (!itemsSheetName) return 0;
  const itemsSheet = e.source.getSheetByName(itemsSheetName);
  if (!itemsSheet) return 0;

  const headers = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0]
    .map((value) => normalizeTextKey_(value));
  const row = rawSheet.getRange(rowIndex, 1, 1, rawSheet.getLastColumn()).getValues()[0];
  const deleteIds = extractDeleteTargetIdsFromRawRow_(row, headers);
  if (!deleteIds.length) return 0;
  if (!canApplyDeletionForSubmission_(e, row, headers)) {
    Logger.log(`Suppression ignoree: droits insuffisants pour la ligne ${rowIndex} (${rawSheet.getName()}).`);
    return 0;
  }
  return deactivateItemsByIds_(itemsSheet, deleteIds);
}

function getSubmissionActorEmail_(e, row, normalizedHeaders) {
  const fromSubmission = getRespondentEmailFromSubmission_(e);
  if (fromSubmission) return fromSubmission;
  const emailIdx = findHeaderIndex_(normalizedHeaders, [
    "email operateur",
    "email demandeur",
    "email controleur",
    "adresse e-mail",
    "email address",
    "email",
  ]);
  return normalizeEmail_(readByIndex_(row, emailIdx, -1));
}

function canApplyDeletionForSubmission_(e, row, normalizedHeaders) {
  const actorEmail = getSubmissionActorEmail_(e, row, normalizedHeaders);
  if (!actorEmail) return false;
  try {
    const adminDashboard = resolveAdminDashboard_("Controle suppression article");
    const actorAccess = getAccessProfileForEmail_(adminDashboard, actorEmail);
    return canUseNativeDestructiveActions_(actorAccess);
  } catch (error) {
    Logger.log(`Controle suppression warning: ${String(error.message || error)}`);
    return false;
  }
}

function extractDeleteTargetIdsFromRawRow_(row, normalizedHeaders) {
  const ids = [];
  const deleteIdx = findHeaderIndex_(normalizedHeaders, ["supprimer des articles", "articles a supprimer", "delete items", "suppression articles"]);
  const itemIdx = findItemFieldIndex_(normalizedHeaders);
  const movementTypeIdx = findHeaderIndex_(normalizedHeaders, ["type de mouvement", "movementtype", "movement type"]);

  const deleteRaw = readByIndex_(row, deleteIdx, -1);
  extractItemIdsFromResponseCell_(deleteRaw).forEach((id) => ids.push(id));

  const movementType = normalizeTextKey_(readByIndex_(row, movementTypeIdx, -1));
  if (movementType === "delete_item" || movementType === "suppression" || movementType === "delete") {
    const itemRaw = readByIndex_(row, itemIdx, 2);
    extractItemIdsFromResponseCell_(itemRaw).forEach((id) => ids.push(id));
  }

  const dedup = {};
  ids.forEach((id) => {
    const key = normalizeTextKey_(id);
    if (!key || dedup[key]) return;
    dedup[key] = id;
  });
  return Object.keys(dedup).map((key) => dedup[key]);
}

function extractItemIdsFromResponseCell_(value) {
  const raw = String(value || "").trim();
  if (!raw) return [];

  const ids = [];
  raw.split(/[,\n;]/).forEach((token) => {
    const parsed = parseItemInput_(token);
    const id = String(parsed.itemId || "").trim();
    if (id) ids.push(id);
  });
  return ids;
}

function deactivateItemsByIds_(itemsSheet, ids) {
  if (!itemsSheet || !ids || !ids.length || itemsSheet.getLastRow() < 2) return 0;

  const wanted = {};
  ids.forEach((id) => {
    const key = normalizeTextKey_(id);
    if (key) wanted[key] = true;
  });
  if (!Object.keys(wanted).length) return 0;

  const data = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 11).getValues();
  let changed = 0;
  data.forEach((row) => {
    const idKey = normalizeTextKey_(row[0]);
    if (!idKey || !wanted[idKey]) return;
    if (!isTruthy_(row[7])) return;
    row[7] = false;
    row[10] = new Date();
    changed += 1;
  });

  if (changed > 0) {
    itemsSheet.getRange(2, 1, data.length, 11).setValues(data);
  }
  return changed;
}

function extractIdFromDriveUrl_(url) {
  const value = String(url || "").trim();
  const match = value.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match && match[1] ? match[1] : "";
}

function rawSheetNameForFormKey_(formKey) {
  const key = String(formKey || "").toUpperCase();
  if (key.indexOf("FORM_FIRE_MOVEMENT_") === 0) return "FIRE_FORM_MOVEMENTS_RAW";
  if (isLegacyFireInventoryFormKey_(key)) return rawSheetNameForFireReplenishment_();
  if (key.indexOf("FORM_FIRE_ITEM_CREATE_") === 0) return "FIRE_FORM_MOVEMENTS_RAW";
  if (key.indexOf("FORM_PHARMA_MOVEMENT_") === 0) return "PHARMA_FORM_MOVEMENTS_RAW";
  if (key.indexOf("FORM_PHARMA_INVENTORY_") === 0) return "PHARMA_FORM_INVENTORY_RAW";
  if (key.indexOf("FORM_PHARMA_ITEM_CREATE_") === 0) return "PHARMA_FORM_MOVEMENTS_RAW";
  return "";
}

function syncFireItems_(spreadsheet) {
  let added = 0;
  added += syncItemsFromRaw_(spreadsheet, "FIRE_FORM_MOVEMENTS_RAW", "ITEMS", "incendie");
  // Le formulaire FIRE_FORM_INVENTORY_RAW est dedie aux demandes de reapprovisionnement.
  // Il ne doit pas creer automatiquement de nouveaux articles dans ITEMS.
  return added;
}

function syncPharmaItems_(spreadsheet) {
  let added = 0;
  added += syncItemsFromRaw_(spreadsheet, "PHARMA_FORM_MOVEMENTS_RAW", "ITEMS_PHARMACY", "pharmacie");
  added += syncItemsFromRaw_(spreadsheet, "PHARMA_FORM_INVENTORY_RAW", "ITEMS_PHARMACY", "pharmacie");
  return added;
}

function syncItemsFromRaw_(spreadsheet, rawSheetName, itemsSheetName, moduleLabel) {
  const rawSheet = spreadsheet.getSheetByName(rawSheetName);
  const itemsSheet = spreadsheet.getSheetByName(itemsSheetName);
  if (!rawSheet || !itemsSheet) return 0;

  const rawLastRow = rawSheet.getLastRow();
  if (rawLastRow < 2) return 0;

  const existing = itemsSheet.getLastRow() > 1
    ? itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 11).getValues()
    : [];
  const knownIds = {};
  const knownNames = {};
  existing.forEach((row) => {
    const idKey = normalizeTextKey_(row[0]);
    const nameKey = normalizeTextKey_(row[3]);
    if (idKey) knownIds[idKey] = true;
    if (nameKey) knownNames[nameKey] = true;
  });

  const headers = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0]
    .map((value) => normalizeTextKey_(value));
  const siteIdx = findHeaderIndex_(headers, ["site concerne", "sitekey", "site"]);
  const itemIdx = findItemFieldIndex_(headers);
  const itemNameIdx = findHeaderIndex_(headers, ["nom article", "itemname", "item name"]);
  const categoryIdx = findHeaderIndex_(headers, ["categorie article", "categorie"]);
  const unitIdx = findHeaderIndex_(headers, ["unite article", "unite"]);
  const thresholdIdx = findHeaderIndex_(headers, ["seuil mini article", "seuil", "minthreshold"]);

  const rawValues = rawSheet.getRange(2, 1, rawLastRow - 1, rawSheet.getLastColumn()).getValues();
  const rowsToInsert = [];

  rawValues.forEach((row) => {
    if (isDeleteIntentRow_(row, headers)) return;

    const siteKey = String(readByIndex_(row, siteIdx, 1) || "").trim(); // fallback col B
    const itemInput = String(readByIndex_(row, itemIdx, 2) || "").trim(); // fallback col C
    const itemNameInput = String(readByIndex_(row, itemNameIdx, -1) || "").trim();
    const categoryInput = String(readByIndex_(row, categoryIdx, -1) || "").trim();
    const unitInput = String(readByIndex_(row, unitIdx, -1) || "").trim();
    const thresholdInput = readByIndex_(row, thresholdIdx, -1);
    const parsed = parseItemInput_(itemInput);
    const itemId = parsed.itemId;
    const itemName = (
      (itemNameInput && (!parsed.itemName || normalizeTextKey_(parsed.itemName) === normalizeTextKey_(parsed.itemId)))
        ? itemNameInput
        : parsed.itemName
    );
    const idKey = normalizeTextKey_(itemId);
    const nameKey = normalizeTextKey_(itemName);

    if (!idKey) return;
    if (knownIds[idKey] || (nameKey && knownNames[nameKey])) return;

    knownIds[idKey] = true;
    if (nameKey) knownNames[nameKey] = true;
    rowsToInsert.push([
      itemId,           // ItemID
      moduleLabel,      // Module
      siteKey,          // SiteKey
      itemName || itemId, // ItemName
      categoryInput || "AutoImported", // Category
      unitInput || "unit",             // Unit
      parseNumberOrDefault_(thresholdInput, 0), // MinThreshold
      true,             // IsActive
      "",               // SupplierID
      "",               // StorageZone
      new Date(),       // LastUpdatedAt
    ]);
  });
  if (!rowsToInsert.length) return 0;

  const startRow = Math.max(2, itemsSheet.getLastRow() + 1);
  itemsSheet.getRange(startRow, 1, rowsToInsert.length, 11).setValues(rowsToInsert);
  return rowsToInsert.length;
}

function isDeleteIntentRow_(row, normalizedHeaders) {
  return extractDeleteTargetIdsFromRawRow_(row, normalizedHeaders).length > 0;
}

function findItemFieldIndex_(normalizedHeaders) {
  const strictTokens = [
    "article concerne",
    "itemid",
    "item id",
    "id article",
    "article selectionne",
    "article choisi",
    "article a reapprovisionner",
  ];
  const strict = findHeaderIndex_(normalizedHeaders, strictTokens);
  if (strict >= 0) return strict;
  return 2;
}

function columnToLetter_(columnNumber) {
  let column = Math.max(1, Number(columnNumber) || 1);
  let letter = "";
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}

function buildRawColumnResolver_(rawSheet) {
  if (!rawSheet) {
    return {
      byTokens: () => "",
      optionalByTokens: () => "",
      item: () => "",
    };
  }

  const sheetName = rawSheet.getName();
  const headers = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0]
    .map((value) => normalizeTextKey_(value));

  const toRangeRef = (columnNumber) => {
    const col = Math.max(1, Number(columnNumber) || 1);
    const letter = columnToLetter_(col);
    return `${sheetName}!${letter}2:${letter}`;
  };

  return {
    byTokens(tokens, fallbackColumnNumber) {
      const idx = findHeaderIndex_(headers, tokens || []);
      return toRangeRef(idx >= 0 ? idx + 1 : fallbackColumnNumber);
    },
    optionalByTokens(tokens) {
      const idx = findHeaderIndex_(headers, tokens || []);
      return idx >= 0 ? toRangeRef(idx + 1) : "";
    },
    item(fallbackColumnNumber) {
      const idx = findItemFieldIndex_(headers);
      return toRangeRef(idx >= 0 ? idx + 1 : fallbackColumnNumber);
    },
  };
}

function readByIndex_(row, index, fallbackIndex) {
  if (index >= 0 && index < row.length) return row[index];
  if (fallbackIndex >= 0 && fallbackIndex < row.length) return row[fallbackIndex];
  return "";
}

function parseItemInput_(value) {
  const raw = String(value || "").trim();
  if (!raw) return { itemId: "", itemName: "" };

  if (raw.indexOf("|") !== -1) {
    const parts = raw.split("|");
    const itemId = String(parts[0] || "").trim();
    const itemName = String(parts[1] || "").trim();
    return { itemId: itemId || raw, itemName: itemName || itemId || raw };
  }

  return { itemId: raw, itemName: raw };
}

function findResponseSheetCandidate_(spreadsheet, expectedName) {
  const expected = spreadsheet.getSheetByName(expectedName);
  if (expected) return expected;

  const candidates = spreadsheet.getSheets().filter((sheet) => /form responses|reponses|responses/i.test(sheet.getName()));
  if (!candidates.length) return null;

  candidates.sort((a, b) => b.getSheetId() - a.getSheetId());
  return candidates[0];
}

function ensureFormDestination_(form, spreadsheet, rawSheetName, keepExistingData, forceReset) {
  let currentDestinationId = "";
  try {
    currentDestinationId = form.getDestinationId() || "";
  } catch (e) {
    currentDestinationId = "";
  }

  if (forceReset) {
    try {
      if (typeof form.removeDestination === "function") {
        form.removeDestination();
      }
    } catch (e) {
      // No-op.
    }
    currentDestinationId = "";
  }

  const targetSpreadsheetId = spreadsheet.getId();
  const beforeSheetIds = collectSheetIds_(spreadsheet);
  let responseSheet = null;
  const changed = currentDestinationId !== targetSpreadsheetId;

  if (changed) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, targetSpreadsheetId);
    responseSheet = waitForFormResponseSheet_(spreadsheet, beforeSheetIds, 40, 500);
  }

  if (!responseSheet) {
    responseSheet = findResponseSheetCandidate_(spreadsheet, rawSheetName);
  }
  if (!responseSheet) {
    throw new Error(`Impossible de localiser l'onglet reponses pour ${rawSheetName}.`);
  }

  if (responseSheet.getName() !== rawSheetName) {
    const existingRaw = spreadsheet.getSheetByName(rawSheetName);
    if (existingRaw && existingRaw.getSheetId() !== responseSheet.getSheetId()) {
      const backupName = `${rawSheetName}_BACKUP_${Utilities.formatDate(new Date(), DEPLOYMENT_CONFIG.timezone, "yyyyMMdd_HHmmss")}`;
      existingRaw.setName(backupName);
      if (keepExistingData) {
        appendSheetData_(existingRaw, responseSheet);
      }
    }
    responseSheet.setName(rawSheetName);
  }

  return {
    changed,
    spreadsheetId: targetSpreadsheetId,
    rawSheetName,
  };
}

function attachFormToSpreadsheet_(form, spreadsheet, rawSheetName, keepExistingData, forceReset) {
  const preserve = keepExistingData !== false;
  ensureFormDestination_(form, spreadsheet, rawSheetName, preserve, !!forceReset);
}

function waitForFormResponseSheet_(spreadsheet, beforeSheetIds, maxAttempts, sleepMs) {
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    SpreadsheetApp.flush();
    const sheets = spreadsheet.getSheets();

    const newSheet = sheets.find((sheet) => !beforeSheetIds[sheet.getSheetId()]);
    if (newSheet) return newSheet;

    Utilities.sleep(sleepMs);
  }

  return null;
}
