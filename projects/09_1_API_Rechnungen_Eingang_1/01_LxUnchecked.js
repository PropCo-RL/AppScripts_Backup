const LX_KEY_IMPORT = PropertiesService.getScriptProperties().getProperty('lxKey');
const GEMINI_KEY_IMPORT = PropertiesService.getScriptProperties().getProperty('Gemini-Key');

function extractVouchersToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Belege");
  const sheetAmz = ss.getSheetByName("AMZ");
  const sheetContacts = ss.getSheetByName("Kontakte");
  const sheetCategories = ss.getSheetByName("d");

  // 1. Kategorien dynamisch aus Tab 'd' auslesen (nur Typ 'outgo')
  let categoriesMap = import_getCategoriesFromTabD(sheetCategories);
  let categoryNamesList = Object.keys(categoriesMap);

  if (categoryNamesList.length === 0) {
    Logger.log("FEHLER: Keine Ausgaben-Kategorien im Tab 'd' gefunden.");
    return;
  }

  // 2. Kontakte & bereits vorhandene Belege/Rechnungsnummern aus dem Sheet einlesen
  let manualContactsMap = import_getManualContactsMap(sheetContacts);
  let existingVoucherIds = import_getExistingVoucherIds([sheetLog, sheetAmz]);
  let existingInvoiceMap = import_getExistingInvoiceMap([sheetLog, sheetAmz]);

  // Tabellen-Header sicherstellen (Spalte O: IBAN Gemini, Spalte P: IBAN Match)
  [sheetLog, sheetAmz].forEach(sh => {
    if (sh) {
      if (sh.getRange(1, 15).getValue() === "") sh.getRange(1, 15).setValue("IBAN Gemini").setFontWeight("bold");
      if (sh.getRange(1, 16).getValue() === "") sh.getRange(1, 16).setValue("IBAN Match").setFontWeight("bold");
    }
  });

  // 3. Ungeprüfte Belege bei Lexoffice abfragen
  const urlList = "https://api.lexware.io/v1/voucherlist?voucherType=purchaseinvoice,purchasecreditnote,invoice,creditnote&voucherStatus=unchecked&size=250&sort=createdDate,DESC";
  const optionsGet = {
    method: "get",
    headers: { "Authorization": "Bearer " + LX_KEY_IMPORT, "Accept": "application/json" },
    muteHttpExceptions: true
  };

  const responseList = UrlFetchApp.fetch(urlList, optionsGet);
  const resultList = JSON.parse(responseList.getContentText());

  if (!resultList.content || resultList.content.length === 0) {
    Logger.log("Keine ungeprüften Eingangsbelege in Lexware gefunden.");
    return;
  }

  // 4. Jeden Beleg verarbeiten
  for (const voucherMeta of resultList.content) {
    const voucherId = voucherMeta.id;

    if (voucherMeta.voucherType === "salesinvoice" || voucherMeta.voucherType === "salescreditnote" || existingVoucherIds.includes(voucherId)) {
      continue;
    }

    const voucherDetails = import_getVoucherDetails(voucherId);
    if (!voucherDetails) continue;

    const fileBlob = import_getVoucherFileBlob(voucherDetails);
    let parsedData = {};

    if (fileBlob) {
      Logger.log(`Starte Gemini-Analyse für Beleg-ID: ${voucherId} (${fileBlob.getContentType()}, Size: ${fileBlob.getBytes().length} bytes)`);
      parsedData = import_analyzeInvoiceWithGeminiInteractions(fileBlob, categoryNamesList);
    } else {
      Logger.log(`WARNUNG: Kein PDF/Datei für Beleg ${voucherId} von Lexoffice erhalten. Überspringe Gemini.`);
    }

    const company = (parsedData.companyName || voucherMeta.contactName || "Unbekannt").trim();
    const voucherNumber = (parsedData.voucherNumber || voucherMeta.voucherNumber || "").trim();

    if (voucherNumber !== "" && company !== "Unbekannt") {
      const invoiceKey = `${company.toLowerCase()}|${voucherNumber.toLowerCase()}`;
      if (existingInvoiceMap[invoiceKey]) {
        Logger.log(`Duplikat erkannt für ${company} (${voucherNumber}). Überspringe Eintrag im Sheet.`);
        continue;
      }
      existingInvoiceMap[invoiceKey] = { voucherId: voucherId };
    }

    const isCreditNote = parsedData.isCreditNote === true || 
                         voucherMeta.voucherType === "purchasecreditnote" ||
                         (parsedData.grossAmount !== undefined && parsedData.grossAmount < 0) ||
                         (voucherMeta.totalAmount !== undefined && voucherMeta.totalAmount < 0);

    const voucherType = isCreditNote ? "purchasecreditnote" : "purchaseinvoice";

    let rawDate = parsedData.voucherDate || voucherMeta.voucherDate || new Date();
    let formattedDate = import_formatDateToISO(rawDate);
    let formattedDueDate = import_formatDateToISO(parsedData.dueDate || rawDate);

    const isAmazon = company.toLowerCase().includes("amazon");
    let taxPercent = parsedData.taxRatePercentage !== undefined ? parsedData.taxRatePercentage : 19;
    let categoryName = parsedData.lxCategory || "Sonstige Ausgaben";
    let categoryId = categoriesMap[categoryName] || categoriesMap["Sonstige Ausgaben"] || "";

    if (isAmazon) {
      const vNumUpper = voucherNumber.toUpperCase();
      const isEuInvoice = vNumUpper.startsWith("PL") || 
                          vNumUpper.startsWith("FR") || 
                          vNumUpper.startsWith("IT") || 
                          vNumUpper.startsWith("ES") || 
                          vNumUpper.startsWith("CZ") || 
                          parsedData.isEuDelivery === true || 
                          parsedData.taxRatePercentage === 0;

      if (isEuInvoice) {
        categoryName = "Innergemeinschaftlicher Erwerb";
        categoryId = "dba64b05-85f7-4359-8e83-a2556a62eeac";
        taxPercent = 19;
      } else {
        categoryName = "Material/Waren";
        categoryId = "efa82f42-fd85-11e1-a21f-0800200c9a66";
        taxPercent = parsedData.taxRatePercentage || 19;
      }
    } else {
      const contactKey = company.toLowerCase();
      if (manualContactsMap[contactKey]) {
        if (manualContactsMap[contactKey].taxRate !== "") {
          taxPercent = Number(manualContactsMap[contactKey].taxRate);
        }
        if (manualContactsMap[contactKey].categoryName !== "") {
          categoryName = manualContactsMap[contactKey].categoryName;
          categoryId = manualContactsMap[contactKey].categoryId || categoriesMap[categoryName] || categoryId;
        }
      }
    }

    // IBAN-EVALUIERUNG
    const rawGeminiIban = parsedData.iban ? String(parsedData.iban).replace(/\s+/g, '').toUpperCase() : "";
    let ibanMatchStatus = "Keine Kontaktdaten";

    const contactKey = company.toLowerCase();
    if (manualContactsMap[contactKey] && manualContactsMap[contactKey].verifiedIban) {
      const verifiedIban = manualContactsMap[contactKey].verifiedIban;
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

    const gross = Math.abs(parsedData.grossAmount || voucherMeta.totalAmount || 0);
    let net = Math.abs(parsedData.netAmount || 0);
    if (!net || net === 0) {
      net = taxPercent > 0 ? Math.round((gross / (1 + taxPercent / 100)) * 100) / 100 : gross;
    }

    const targetSheet = (sheetAmz && isAmazon) ? sheetAmz : sheetLog;

    targetSheet.appendRow([
      formattedDate,      // Spalte A
      company,            // Spalte B
      voucherNumber,      // Spalte C
      "",                 // Spalte D
      voucherType,        // Spalte E
      taxPercent,         // Spalte F
      voucherId,          // Spalte G
      gross,              // Spalte H
      net,                // Spalte I
      categoryName,       // Spalte J
      formattedDueDate,   // Spalte K
      categoryId,         // Spalte L
      "",                 // Spalte M: Verbucht
      "",                 // Spalte N: Zahlungsstatus
      rawGeminiIban,      // Spalte O: IBAN Gemini
      ibanMatchStatus     // Spalte P: IBAN Match
    ]);

    existingVoucherIds.push(voucherId);
  }
}

// --- Hilfsfunktionen ---

function import_getCategoriesFromTabD(sheetCategories) {
  let map = {};
  if (!sheetCategories) return map;
  const lastRow = sheetCategories.getLastRow();
  if (lastRow > 1) {
    const data = sheetCategories.getRange(2, 1, lastRow - 1, 3).getValues();
    data.forEach(row => {
      const name = String(row[0]).trim();
      const uuid = String(row[1]).trim();
      const type = String(row[2]).trim().toLowerCase();
      if (name && uuid && (type === "outgo" || type === "")) {
        map[name] = uuid;
      }
    });
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
  } catch (e) { return null; }
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
  } catch (e) { return null; }
}

function import_getManualContactsMap(sheetContacts) {
  let map = {};
  if (!sheetContacts) return map;
  const lastRow = sheetContacts.getLastRow();
  if (lastRow > 1) {
    const data = sheetContacts.getRange(2, 1, lastRow - 1, 7).getValues();
    data.forEach(row => {
      const name = String(row[0]).trim().toLowerCase();
      const taxRate = row[2];
      const categoryName = String(row[3]).trim();
      const categoryId = String(row[4]).trim();
      const verifiedIban = String(row[6]).replace(/\s+/g, '').toUpperCase();
      
      if (name) map[name] = { taxRate, categoryName, categoryId, verifiedIban };
    });
  }
  return map;
}

function import_formatDateToISO(dateInput) {
  if (!dateInput) return "";
  try {
    const d = new Date(dateInput);
    if (isNaN(d.getTime())) return String(dateInput).substring(0, 10);
    return Utilities.formatDate(d, "GMT", "yyyy-MM-dd");
  } catch (e) { return ""; }
}

function import_getExistingVoucherIds(sheets) {
  let ids = [];
  sheets.forEach(sheet => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const colG = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
      colG.forEach(row => { if (row[0]) ids.push(String(row[0]).trim()); });
    }
  });
  return ids;
}

function import_getExistingInvoiceMap(sheets) {
  let map = {};
  sheets.forEach(sheet => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      data.forEach(row => {
        const comp = String(row[1]).trim().toLowerCase();
        const num = String(row[2]).trim().toLowerCase();
        const vId = String(row[6]).trim();
        if (comp && num && vId) map[`${comp}|${num}`] = { voucherId: vId };
      });
    }
  });
  return map;
}

// GEMINI ANALYSE FÜR INTERACTIONS API (Inklusive Steps-Parsing für Gemini 3.8)
function import_analyzeInvoiceWithGeminiInteractions(blob, allowedCategories) {
  const base64Data = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();
  const isPdf = mimeType === "application/pdf";
  const inputType = isPdf ? "document" : "image";
  
  const url = `https://generativelanguage.googleapis.com/v1beta/interactions?key=${GEMINI_KEY_IMPORT}`;

  const schema = {
    type: "object",
    properties: {
      companyName: { type: "string", description: "Exakter Name der ausstellenden Firma laut Beleg" },
      voucherNumber: { type: "string", description: "Rechnungsnummer oder Gutschriftsnummer" },
      voucherDate: { type: "string", description: "Rechnungsdatum/Gutschriftsdatum im Format YYYY-MM-DD" },
      dueDate: { type: "string", description: "Zahlungsziel/Fälligkeitsdatum im Format YYYY-MM-DD" },
      netAmount: { type: "number", description: "Nettobetrag als Zahl" },
      grossAmount: { type: "number", description: "Bruttobetrag/Gesamtbetrag als Zahl" },
      taxRatePercentage: { type: "number", description: "MwSt-Satz in Prozent laut Beleg (z.B. 0 oder 19)." },
      iban: { type: "string", description: "Die auf dem Beleg stehende IBAN des Zahlungsempfängers/Lieferanten (ohne Leerzeichen)." },
      isEuDelivery: { type: "boolean", description: "TRUE wenn auf dem Beleg EU-Lieferung/Artikel 138/Reverse Charge steht." },
      isCreditNote: { type: "boolean", description: "TRUE bei Gutschrift oder Stornorechnung." },
      lxCategory: { type: "string", enum: allowedCategories, description: "Passende Ausgaben-Kategorie." }
    },
    required: ["companyName", "voucherNumber", "voucherDate", "grossAmount", "taxRatePercentage", "isEuDelivery", "isCreditNote", "lxCategory"]
  };

  const payload = {
    model: "gemini-3.8-flash",
    input: [
      { type: inputType, data: base64Data, mime_type: mimeType },
      { type: "text", text: "Analysiere diese Eingangsrechnung oder Gutschrift exakt. Extrahiere alle geforderten Felder inklusive der IBAN des Zahlungsempfängers." }
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

    // Falls output_text nicht direkt gefüllt ist, aus den Steps der Interactions API extrahieren
    if (!jsonString && resJson.steps) {
      for (const step of resJson.steps) {
        if (step.type === "model_output" && step.content) {
          for (const item of step.content) {
            if (item.type === "text" && item.text) {
              jsonString = item.text;
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
      Logger.log("GEMINI ANTWORTE OHNE INHALT: " + rawText);
    }
  } catch (e) {
    Logger.log("FEHLER im Gemini Request: " + e.toString());
  }
  return {};
}