const LX_KEY_IMPORT = PropertiesService.getScriptProperties().getProperty('lxKey');
const GEMINI_KEY_IMPORT = PropertiesService.getScriptProperties().getProperty('Gemini-Key');

/**
 * Zentrale Orchestrator-Funktion für den Lexoffice- und Gemini-Import.
 */
function extractVouchersToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Belege");
  const sheetAmz = ss.getSheetByName("AMZ");
  const sheetContacts = ss.getSheetByName("Kontakte");
  const sheetCategories = ss.getSheetByName("d");

  if (!sheetLog || !sheetContacts || !sheetCategories) {
    Logger.log("FEHLER: Erforderliche Blätter (Belege, Kontakte oder d) nicht gefunden.");
    return;
  }

  // 1. Kategorien laden & für Schema bereinigen
  const categoriesMap = import_getCategoriesFromTabD(sheetCategories);
  const categoryNamesList = Object.keys(categoriesMap).filter(k => k && k.trim() !== "");
  if (categoryNamesList.length === 0) {
    Logger.log("FEHLER: Keine Ausgaben-Kategorien im Tab 'd' gefunden.");
    return;
  }

  // 2. Tabellen-Header & 21 Spalten sicherstellen
  [sheetLog, sheetAmz].forEach(sh => {
    if (sh) import_sicherstelleHeaderUndSpalten(sh);
  });

  // 3. Kontakte & Bestandsdaten einlesen
  const manualContactsMap = import_getManualContactsMap(sheetContacts);
  const existingVoucherIds = import_getExistingVoucherIds([sheetLog, sheetAmz]);
  const existingInvoiceMap = import_getExistingInvoiceMap([sheetLog, sheetAmz]);

  // 4. Ungeprüfte Belege abrufen
  const voucherList = import_holeUngepruefteBelege();
  if (voucherList.length === 0) {
    Logger.log("Keine ungeprüften Belege vorhanden.");
    return;
  }

  // 5. Belege verarbeiten
  const neueZeilenLog = [];
  const neueZeilenAmz = [];

  for (let i = 0; i < voucherList.length; i++) {
    const voucherMeta = voucherList[i];
    const voucherId = voucherMeta.id;

    if (voucherMeta.voucherType === "salesinvoice" || voucherMeta.voucherType === "salescreditnote") continue;
    if (existingVoucherIds.includes(voucherId)) continue;

    const zeilenErgebnis = import_verarbeiteEinzelnenBeleg(
      voucherMeta,
      categoryNamesList,
      categoriesMap,
      manualContactsMap,
      existingInvoiceMap
    );

    // Beleg-ID im Speicher vormerken (auch wenn er durch Filter ignoriert wurde)
    existingVoucherIds.push(voucherId);

    if (!zeilenErgebnis) continue;

    if (zeilenErgebnis.istAmazon && sheetAmz) {
      neueZeilenAmz.push(zeilenErgebnis.zeile);
    } else {
      neueZeilenLog.push(zeilenErgebnis.zeile);
    }

    if (zeilenErgebnis.invoiceKey) {
      existingInvoiceMap[zeilenErgebnis.invoiceKey] = { voucherId: voucherId };
    }
  }

  // 6. Batch-Schreiben & finale Sortierung nach Spalte T
  import_schreibeUndSortiereBatch(sheetLog, neueZeilenLog);
  if (sheetAmz) import_schreibeUndSortiereBatch(sheetAmz, neueZeilenAmz);

  Logger.log(`ABGESCHLOSSEN: ${neueZeilenLog.length} in 'Belege', ${neueZeilenAmz.length} in 'AMZ'.`);
}

/**
 * Holt ungeprüfte Belege aus Lexoffice (neueste zuerst).
 */
function import_holeUngepruefteBelege() {
  const urlList = "https://api.lexware.io/v1/voucherlist?voucherType=purchaseinvoice,purchasecreditnote,invoice,creditnote&voucherStatus=unchecked&size=250&sort=createdDate,DESC";
  const optionsGet = {
    method: "get",
    headers: { "Authorization": "Bearer " + LX_KEY_IMPORT, "Accept": "application/json" },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(urlList, optionsGet);
    if (response.getResponseCode() !== 200) {
      Logger.log(`Lexoffice Status ${response.getResponseCode()}: ${response.getContentText()}`);
      return [];
    }
    const result = JSON.parse(response.getContentText());
    return result.content || [];
  } catch (e) {
    Logger.log("Fehler beim Abruf von Lexoffice: " + e.toString());
    return [];
  }
}

/**
 * Einzelbeleg verarbeiten inklusive Duplikat- und Ignorierprüfung.
 */
function import_verarbeiteEinzelnenBeleg(voucherMeta, categoryNamesList, categoriesMap, manualContactsMap, existingInvoiceMap) {
  const voucherId = voucherMeta.id;
  const voucherDetails = import_getVoucherDetails(voucherId);
  if (!voucherDetails) return null;

  const fileBlob = import_getVoucherFileBlob(voucherDetails);
  let parsedData = {};

  if (fileBlob) {
    Logger.log(`Starte Gemini-Analyse für Beleg: ${voucherId} (${fileBlob.getContentType()}, ${fileBlob.getBytes().length} Bytes)`);
    parsedData = import_analyzeInvoiceWithGeminiInteractions(fileBlob, categoryNamesList);
  } else {
    Logger.log(`Kein PDF für Beleg ${voucherId} erhalten. Überspringe Gemini.`);
  }

  const company = (parsedData.companyName || voucherMeta.contactName || "Unbekannt").trim();
  const contactKey = company.toLowerCase();
  const contactConfig = manualContactsMap[contactKey];

  // Prüfung 1: Kontakt mit 'x' markiert -> überspringen
  if (contactConfig && contactConfig.istIgnoriert) {
    Logger.log(`Kontakt ${company} ignoriert (Spalte K = x).`);
    return null;
  }

  // Prüfung 2: Duplikate anhand Firma & Rechnungsnummer
  const voucherNumber = (parsedData.voucherNumber || voucherMeta.voucherNumber || "").trim();
  let invoiceKey = null;
  if (voucherNumber !== "" && company !== "Unbekannt") {
    invoiceKey = `${contactKey}|${voucherNumber.toLowerCase()}`;
    if (existingInvoiceMap[invoiceKey]) {
      Logger.log(`Duplikat erkannt für ${company} (${voucherNumber}). Überspringe.`);
      return null;
    }
  }

  const zeilenObj = import_baueZeilenObjekt(
    voucherMeta,
    voucherDetails,
    parsedData,
    company,
    contactConfig,
    categoriesMap
  );

  return {
    zeile: zeilenObj.zeile,
    istAmazon: zeilenObj.istAmazon,
    invoiceKey: invoiceKey
  };
}

/**
 * Baut das finale 21-spaltige Zeilen-Array.
 */
function import_baueZeilenObjekt(voucherMeta, voucherDetails, parsedData, company, contactConfig, categoriesMap) {
  const voucherNumber = (parsedData.voucherNumber || voucherMeta.voucherNumber || "").trim();
  const rawDate = parsedData.voucherDate || voucherMeta.voucherDate || new Date();
  const formattedDate = import_formatDateToISO(rawDate);
  const formattedDueDate = import_formatDateToISO(parsedData.dueDate || rawDate);

  const isAmazon = company.toLowerCase().includes("amazon");
  const belegKat = import_ermittleKategorieUndSteuer(company, voucherNumber, parsedData, contactConfig, categoriesMap);

  const gross = Math.abs(parsedData.grossAmount !== undefined ? parsedData.grossAmount : (voucherMeta.totalAmount || 0));
  let net = Math.abs(parsedData.netAmount || 0);
  if (!net || net === 0) {
    net = belegKat.taxPercent > 0 ? Math.round((gross / (1 + belegKat.taxPercent / 100)) * 100) / 100 : gross;
  }

  const ibanCheck = import_evaluiereIban(parsedData.iban, contactConfig);
  const wiseBatchLog = (contactConfig && contactConfig.isSepa && !contactConfig.verifiedIban) ? "Keine Zahlung Notwendig" : "";

  const eingangsZeitstempel = voucherDetails.createdDate 
    ? import_formatDateTimeToISO(voucherDetails.createdDate) 
    : import_formatDateTimeToISO(new Date());

  const metadatenJson = JSON.stringify({
    eingangsdatum: eingangsZeitstempel,
    rechnungsdatum: formattedDate,
    adresse: parsedData.issuerAddress || "",
    steuernummer: parsedData.issuerTaxNumber || "",
    vatId: parsedData.issuerVatId || "",
    telefon: parsedData.issuerPhone || "",
    bic: parsedData.issuerBic || ""
  });
  const positionenJson = JSON.stringify(parsedData.lineItems || []);
  const issuerEmail = parsedData.issuerEmail || (contactConfig ? contactConfig.email : "");

  const zeile = [
    formattedDate,                                      // Spalte A: Rechnungsdatum (YYYY-MM-DD)
    company,                                            // Spalte B: Firma
    voucherNumber,                                      // Spalte C: Rechnungs-Nr.
    issuerEmail,                                        // Spalte D: E-Mail
    import_bestimmeVoucherType(parsedData, voucherMeta), // Spalte E: Belegtyp
    belegKat.taxPercent,                                // Spalte F: MwSt.
    voucherMeta.id,                                     // Spalte G: Lx Voucher ID
    gross,                                              // Spalte H: Brutto
    net,                                                // Spalte I: Netto
    belegKat.categoryName,                              // Spalte J: Lx Kategorie
    formattedDueDate,                                   // Spalte K: Due Date
    belegKat.categoryId,                                // Spalte L: Kategorie ID
    "",                                                 // Spalte M: Verbucht
    "",                                                 // Spalte N: Zahlungsstatus
    ibanCheck.rawGeminiIban,                             // Spalte O: IBAN Gemini
    ibanCheck.ibanMatchStatus,                           // Spalte P: IBAN Match
    wiseBatchLog,                                       // Spalte Q: Wise Batch Log
    "",                                                 // Spalte R: Notiz / Reserve 1
    "",                                                 // Spalte S: Reserve 2
    metadatenJson,                                      // Spalte T: Metadaten JSON
    positionenJson                                      // Spalte U: Positionen JSON
  ];

  return { zeile, istAmazon };
}

/**
 * Schreibt neue Zeilen per Batch und sortiert anschließend das Sheet nach Spalte T.
 */
function import_schreibeUndSortiereBatch(sheet, neueZeilen) {
  if (!sheet) return;

  if (neueZeilen.length > 0) {
    const startZeile = sheet.getLastRow() + 1;
    sheet.getRange(startZeile, 1, neueZeilen.length, neueZeilen[0].length).setValues(neueZeilen);
  }

  import_sortiereSheetNachSpalteT(sheet);
}

/**
 * Sortiert alle Datenzeilen nach dem Eingangszeitstempel in Spalte T (aufsteigend).
 */
function import_sortiereSheetNachSpalteT(sheet) {
  const letzteZeile = sheet.getLastRow();
  if (letzteZeile <= 2) return;

  const anzahlSpalten = Math.max(sheet.getLastColumn(), 21);
  const datenRange = sheet.getRange(2, 1, letzteZeile - 1, anzahlSpalten);
  const werte = datenRange.getValues();

  werte.sort((a, b) => {
    const zeitA = import_parseEingangszeit(a[19], a[0]);
    const zeitB = import_parseEingangszeit(b[19], b[0]);
    return zeitA.localeCompare(zeitB);
  });

  datenRange.setValues(werte);
}

/**
 * Normalisiert den Eingangszeitstempel aus JSON oder Date-Objekt zu ISO (YYYY-MM-DD HH:mm).
 */
function import_parseEingangszeit(jsonFeld, datumFallback) {
  if (jsonFeld) {
    try {
      const parsed = JSON.parse(jsonFeld);
      if (parsed.eingangsdatum) return String(parsed.eingangsdatum);
    } catch (e) {}
  }
  if (datumFallback) {
    if (datumFallback instanceof Date) {
      return Utilities.formatDate(datumFallback, Session.getScriptTimeZone(), "yyyy-MM-dd 00:00");
    }
    const str = String(datumFallback).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
      return str.substring(0, 10) + " 00:00";
    }
  }
  return "1970-01-01 00:00";
}

/**
 * Bestimmt MwSt und Kategorie.
 */
function import_ermittleKategorieUndSteuer(company, voucherNumber, parsedData, contactConfig, categoriesMap) {
  let taxPercent = parsedData.taxRatePercentage !== undefined ? parsedData.taxRatePercentage : 19;
  let categoryName = parsedData.lxCategory || "Sonstige Ausgaben";
  let categoryId = categoriesMap[categoryName] || categoriesMap["Sonstige Ausgaben"] || "";

  if (company.toLowerCase().includes("amazon")) {
    const vNumUpper = (voucherNumber || "").toUpperCase();
    const isEuInvoice = /^(PL|FR|IT|ES|CZ)/.test(vNumUpper) || parsedData.isEuDelivery === true || parsedData.taxRatePercentage === 0;

    if (isEuInvoice) {
      return { taxPercent: 19, categoryName: "Innergemeinschaftlicher Erwerb", categoryId: "dba64b05-85f7-4359-8e83-a2556a62eeac" };
    }
    return { taxPercent: parsedData.taxRatePercentage || 19, categoryName: "Material/Waren", categoryId: "efa82f42-fd85-11e1-a21f-0800200c9a66" };
  }

  if (contactConfig) {
    if (contactConfig.taxRate !== "") taxPercent = Number(contactConfig.taxRate);
    if (contactConfig.categoryName !== "") {
      categoryName = contactConfig.categoryName;
      categoryId = contactConfig.categoryId || categoriesMap[categoryName] || categoryId;
    }
  }

  return { taxPercent, categoryName, categoryId };
}

/**
 * Evaluierung der IBAN.
 */
function import_evaluiereIban(geminiIbanRaw, contactConfig) {
  const rawGeminiIban = geminiIbanRaw ? String(geminiIbanRaw).replace(/\s+/g, '').toUpperCase() : "";
  let ibanMatchStatus = "Keine Kontaktdaten";

  if (contactConfig && contactConfig.verifiedIban) {
    const verifiedIban = contactConfig.verifiedIban;
    if (rawGeminiIban !== "" && rawGeminiIban === verifiedIban) {
      ibanMatchStatus = "Match";
    } else if (rawGeminiIban !== "") {
      ibanMatchStatus = "Kein Match";
    } else {
      ibanMatchStatus = "Keine IBAN im Beleg";
    }
  } else if (rawGeminiIban !== "") {
    ibanMatchStatus = "Ungeprüfter Kontakt";
  }

  return { rawGeminiIban, ibanMatchStatus };
}

function import_bestimmeVoucherType(parsedData, voucherMeta) {
  const isCreditNote = parsedData.isCreditNote === true ||
    voucherMeta.voucherType === "purchasecreditnote" ||
    (parsedData.grossAmount !== undefined && parsedData.grossAmount < 0) ||
    (voucherMeta.totalAmount !== undefined && voucherMeta.totalAmount < 0);
  return isCreditNote ? "purchasecreditnote" : "purchaseinvoice";
}

/**
 * Gemini Interactions API-Aufruf mit bereinigtem Enum und vollständigen Pflichtfeldern.
 */
function import_analyzeInvoiceWithGeminiInteractions(blob, allowedCategories) {
  const base64Data = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();
  const inputType = (mimeType === "application/pdf") ? "document" : "image";
  const url = `https://generativelanguage.googleapis.com/v1beta/interactions?key=${GEMINI_KEY_IMPORT}`;

  const cleanCategories = Array.from(new Set(allowedCategories.filter(c => Boolean(c && c.trim()))));
  if (!cleanCategories.includes("Sonstige Ausgaben")) {
    cleanCategories.push("Sonstige Ausgaben");
  }

  const schema = {
    type: "object",
    properties: {
      companyName: { type: "string", description: "Exakter Name der ausstellenden Firma laut Beleg" },
      issuerEmail: { type: "string", description: "Offizielle E-Mail-Adresse des Ausstellers" },
      issuerAddress: { type: "string", description: "Vollständige Anschrift des Ausstellers (Straße, PLZ, Ort)" },
      issuerTaxNumber: { type: "string", description: "Steuernummer des Ausstellers" },
      issuerVatId: { type: "string", description: "Umsatzsteuer-Identifikationsnummer (z. B. DE123456789)" },
      issuerPhone: { type: "string", description: "Telefonnummer des Ausstellers" },
      issuerBic: { type: "string", description: "BIC der Bankverbindung" },
      voucherNumber: { type: "string", description: "Rechnungs- oder Gutschriftsnummer" },
      voucherDate: { type: "string", description: "Rechnungsdatum im Format YYYY-MM-DD" },
      dueDate: { type: "string", description: "Fälligkeitsdatum im Format YYYY-MM-DD" },
      netAmount: { type: "number", description: "Nettobetrag als reine Zahl" },
      grossAmount: { type: "number", description: "Bruttobetrag als reine Zahl" },
      taxRatePercentage: { type: "number", description: "MwSt-Satz in Prozent (z. B. 0 oder 19)" },
      iban: { type: "string", description: "IBAN des Zahlungsempfängers ohne Leerzeichen" },
      isEuDelivery: { type: "boolean", description: "TRUE bei Reverse Charge / EU-Lieferung" },
      isCreditNote: { type: "boolean", description: "TRUE bei Gutschrift oder Stornorechnung" },
      lxCategory: { type: "string", enum: cleanCategories, description: "Passende Ausgaben-Kategorie" },
      lineItems: {
        type: "array",
        description: "Alle einzelnen Rechnungspositionen",
        items: {
          type: "object",
          properties: {
            menge: { type: "number", description: "Menge / Anzahl" },
            beschreibung: { type: "string", description: "Leistungs- oder Artikelbeschreibung" },
            einzelpreis: { type: "number", description: "Einzelpreis als Zahl" },
            gesamtpreis: { type: "number", description: "Gesamtpreis als Zahl" }
          },
          required: ["beschreibung", "gesamtpreis"]
        }
      }
    },
    required: ["companyName", "voucherNumber", "voucherDate", "grossAmount", "isCreditNote", "isEuDelivery", "lxCategory"]
  };

  const payload = {
    model: "gemini-3.8-flash",
    input: [
      { type: inputType, data: base64Data, mime_type: mimeType },
      { type: "text", text: "Analysiere diese Eingangsrechnung oder Gutschrift exakt. Extrahiere Kopf- und Fußzeilendaten (E-Mail, Anschrift, Steuernummer), Zahlungsdaten sowie alle Positionen." }
    ],
    response_format: { type: "text", mime_type: "application/json", schema: schema }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    const rawText = response.getContentText();
    if (statusCode !== 200) {
      Logger.log(`GEMINI API FEHLER (Status ${statusCode}): ${rawText}`);
      return {};
    }

    const resJson = JSON.parse(rawText);
    let jsonString = resJson.output_text;

    if (!jsonString && resJson.steps) {
      for (let s = 0; s < resJson.steps.length; s++) {
        const step = resJson.steps[s];
        if (step.type === "model_output" && step.content) {
          for (let c = 0; c < step.content.length; c++) {
            if (step.content[c].type === "text" && step.content[c].text) {
              jsonString = step.content[c].text;
              break;
            }
          }
        }
      }
    }

    if (jsonString) {
      const parsed = JSON.parse(jsonString);
      Logger.log(`GEMINI ERFOLG! Extrahierte IBAN: '${parsed.iban || "KEINE IBAN GEFUNDEN"}' für ${parsed.companyName}`);
      return parsed;
    } else {
      Logger.log("GEMINI ANTWORT OHNE INHALT: " + rawText);
    }
  } catch (e) {
    Logger.log("FEHLER im Gemini Request: " + e.toString());
  }

  return {};
}

/**
 * Stellt sicher, dass das Blatt mindestens 21 Spalten besitzt und alle Header vorhanden sind.
 */
function import_sicherstelleHeaderUndSpalten(sheet) {
  if (!sheet) return;

  const benoetigteSpalten = 21;
  const aktuelleSpalten = sheet.getMaxColumns();
  if (aktuelleSpalten < benoetigteSpalten) {
    sheet.insertColumnsAfter(aktuelleSpalten, benoetigteSpalten - aktuelleSpalten);
  }

  const standardHeaders = [
    "Rechnungsdatum", "Firma", "Rechnungs-Nr.", "E-Mail", "Typ",
    "MwSt.", "Lx Voucher ID", "Brutto", "Netto", "Lx Kategorie",
    "Due Date", "Kategorie ID", "Verbucht", "Zahlungsstatus",
    "IBAN Gemini", "IBAN Match", "Wise Batch Log", "Notiz",
    "Reserve", "Metadaten JSON", "Positionen JSON"
  ];

  for (let i = 0; i < standardHeaders.length; i++) {
    const colIdx = i + 1;
    const currentVal = sheet.getRange(1, colIdx).getValue();
    if (currentVal === "") {
      sheet.getRange(1, colIdx).setValue(standardHeaders[i]).setFontWeight("bold");
    }
  }
}

function import_getCategoriesFromTabD(sheetCategories) {
  const map = {};
  if (!sheetCategories) return map;
  const lastRow = sheetCategories.getLastRow();
  if (lastRow > 1) {
    const data = sheetCategories.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      const name = String(data[i][0] || "").trim();
      const uuid = String(data[i][1] || "").trim();
      const type = String(data[i][2] || "").trim().toLowerCase();
      if (name && uuid && (type === "outgo" || type === "")) {
        map[name] = uuid;
      }
    }
  }
  return map;
}

function import_getVoucherDetails(voucherId) {
  try {
    const url = `https://api.lexware.io/v1/vouchers/${voucherId}`;
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": "Bearer " + LX_KEY_IMPORT, "Accept": "application/json" },
      muteHttpExceptions: true
    });
    return JSON.parse(response.getContentText());
  } catch (e) {
    return null;
  }
}

function import_getVoucherFileBlob(voucherDetails) {
  try {
    if (voucherDetails && voucherDetails.files && voucherDetails.files.length > 0) {
      const fileId = voucherDetails.files[0];
      const fileUrl = `https://api.lexware.io/v1/files/${fileId}`;
      return UrlFetchApp.fetch(fileUrl, {
        method: "get",
        headers: { "Authorization": "Bearer " + LX_KEY_IMPORT, "Accept": "*/*" },
        muteHttpExceptions: true
      }).getBlob();
    }
  } catch (e) {
    return null;
  }
}

function import_getManualContactsMap(sheetContacts) {
  const map = {};
  if (!sheetContacts) return map;
  const lastRow = sheetContacts.getLastRow();
  if (lastRow > 1) {
    const data = sheetContacts.getRange(2, 1, lastRow - 1, 11).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const name = String(row[0] || "").trim().toLowerCase();
      if (!name) continue;

      map[name] = {
        taxRate: row[2],
        categoryName: String(row[3] || "").trim(),
        categoryId: String(row[4] || "").trim(),
        isSepa: String(row[5] || "").trim().toUpperCase() === "SEPA",
        verifiedIban: String(row[6] || "").replace(/\s+/g, '').toUpperCase(),
        email: String(row[9] || "").trim().toLowerCase(),
        istIgnoriert: String(row[10] || "").trim().toLowerCase() === "x"
      };
    }
  }
  return map;
}

function import_formatDateToISO(dateInput) {
  if (!dateInput) return "";
  if (dateInput instanceof Date) {
    return Utilities.formatDate(dateInput, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const str = String(dateInput).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.substring(0, 10);
  }
  try {
    const d = new Date(dateInput);
    if (isNaN(d.getTime())) return str.substring(0, 10);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } catch (e) {
    return "";
  }
}

function import_formatDateTimeToISO(dateInput) {
  if (!dateInput) return "";
  try {
    const d = (dateInput instanceof Date) ? dateInput : new Date(dateInput);
    if (isNaN(d.getTime())) return String(dateInput);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  } catch (e) {
    return "";
  }
}

function import_getExistingVoucherIds(sheets) {
  const ids = [];
  sheets.forEach(sheet => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const colG = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
      for (let i = 0; i < colG.length; i++) {
        if (colG[i][0]) ids.push(String(colG[i][0]).trim());
      }
    }
  });
  return ids;
}

function import_getExistingInvoiceMap(sheets) {
  const map = {};
  sheets.forEach(sheet => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      for (let i = 0; i < data.length; i++) {
        const comp = String(data[i][1] || "").trim().toLowerCase();
        const num = String(data[i][2] || "").trim().toLowerCase();
        const vId = String(data[i][6] || "").trim();
        if (comp && num && vId) map[`${comp}|${num}`] = { voucherId: vId };
      }
    }
  });
  return map;
}