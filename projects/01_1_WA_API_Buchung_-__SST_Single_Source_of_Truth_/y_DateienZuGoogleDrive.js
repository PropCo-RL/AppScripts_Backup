// ============================================================================
// SEPARATES ERGÄNZUNGSSCRIPT: PDF-SPEICHERUNG IN GOOGLE DRIVE & LINK-GENERIERUNG
// ============================================================================

const TARGET_FOLDER_ID = "13CMmLQzGkfh46SSGHoW7qh79ItDaOZLf";

/**
 * HAUPTFUNKTION: Verarbeitet alle 2 Bereiche (RE, VL)
 * Kann manuell ausgeführt oder als eigener Zeit-Trigger (z.B. stündlich) eingerichtet werden.
 */
function BBB_syncAllPdfsToDrive() {
  syncLexofficeInvoicesToDrive("RE");
  syncLexofficeInvoicesToDrive("VL");
}

/**
 * Lädt Lexoffice Rechnungs-PDFs für den übergebenen Tab ('RE' oder 'VL')
 * liest Lexoffice ID aus Spalte T (20) und schreibt den Download-Link in Spalte AK (37).
 */
function syncLexofficeInvoicesToDrive(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  
  // Liest alle Daten von Zeile 2 bis Spalte AK (37)
  const values = sheet.getRange(2, 1, lastRow - 1, 37).getValues();

  values.forEach((row, index) => {
    const rowIndex = index + 2;
    const lxInvoiceId = row[19] ? row[19].toString().trim() : ""; // Spalte T (20)
    const currentAkValue = row[36] ? row[36].toString().trim() : ""; // Spalte AK (37)

    // Nur ausführen, wenn eine Lexoffice-ID da ist und Spalte AK noch LEER ist (oder nicht mit http startet)
    if (!lxInvoiceId || lxInvoiceId.startsWith("ERROR") || currentAkValue.startsWith("http")) {
      return;
    }

    try {
      // 1. PDF-Document ID von Lexoffice abfragen
      const docResponse = sendLexofficeGetRequest(`/v1/invoices/${lxInvoiceId}/document`);
      if (docResponse.statusCode !== 200 || !docResponse.body || !docResponse.body.documentFileId) {
        sheet.getRange(rowIndex, 37).setValue(`ERROR Document-ID: ${docResponse.statusCode}`);
        return;
      }

      const documentFileId = docResponse.body.documentFileId;

      // 2. Das Rechnungs-PDF (Binary File) herunterladen
      const pdfBlob = downloadLexofficePdfFile(`/v1/files/${documentFileId}`);
      
      // Dateiname formatieren (z.B. Rechnung_RE_Zeile2_ID.pdf)
      const guestOrCompany = row[1] || row[2] || "Gast";
      const sanitizedName = guestOrCompany.toString().replace(/[^a-zA-Z0-9_-]/g, "_");
      pdfBlob.setName(`Rechnung_${sheetName}_${sanitizedName}_${lxInvoiceId.substring(0, 8)}.pdf`);

      // 3. In Google Drive speichern
      const file = folder.createFile(pdfBlob);
      const fileId = file.getId();

      // Direct-Download-Link für den Server / PHP / GitHub (Ohne Freigabe des Ordners)
      const directDownloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

      // 4. In Spalte AK (37) eintragen
      sheet.getRange(rowIndex, 37).setValue(directDownloadUrl);
      Logger.log(`[${sheetName}] PDF erfolgreich gespeichert für Zeile ${rowIndex}: ${directDownloadUrl}`);

    } catch (err) {
      Logger.log(`[${sheetName}] Fehler Zeile ${rowIndex}: ${err.message}`);
      sheet.getRange(rowIndex, 37).setValue(`ERROR: ${err.message}`);
    }
  });
}

// ============================================================================
// HILFSFUNKTIONEN FÜR DIE LEXOFFICE API
// ============================================================================

function sendLexofficeGetRequest(endpoint) {
  const props = PropertiesService.getScriptProperties();
  const lxKey = props.getProperty("lxKey");
  if (!lxKey) throw new Error("Script Property 'lxKey' fehlt!");

  const targetUrl = "https://api.lexoffice.io" + endpoint;
  const options = {
    "method": "get",
    "headers": {
      "Authorization": "Bearer " + lxKey,
      "Accept": "application/json"
    },
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(targetUrl, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  let responseBody;
  try {
    responseBody = responseText ? JSON.parse(responseText) : null;
  } catch (e) {
    responseBody = responseText;
  }

  return { statusCode: responseCode, body: responseBody };
}

function downloadLexofficePdfFile(endpoint) {
  const props = PropertiesService.getScriptProperties();
  const lxKey = props.getProperty("lxKey");
  if (!lxKey) throw new Error("Script Property 'lxKey' fehlt!");

  const targetUrl = "https://api.lexoffice.io" + endpoint;
  const options = {
    "method": "get",
    "headers": {
      "Authorization": "Bearer " + lxKey,
      "Accept": "application/pdf"
    },
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(targetUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(`PDF Download fehlgeschlagen with Status: ${response.getResponseCode()}`);
  }

  return response.getBlob();
}