// ID des WhatsApp-Webhook Spreadsheets (01.2_WA D360 WebHook)
const WA_WEBHOOK_SPREADSHEET_ID = "1bMPVHjGNnwWDH0KXUyB47_VbLcq_lsFK1jSbdj_GmH0";

function scanWhatsAppForPaymentProofs() {
  Logger.log("=== Starte Suche nach Zahlungsnachweisen via 360dialog ===");
  const d360ApiKey = PropertiesService.getScriptProperties().getProperty('D360-Key');
  if (!d360ApiKey) {
    Logger.log("FEHLER: 'D360-Key' nicht in ScriptProperties gefunden!");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName("Output");
  if (!outputSheet) {
    Logger.log("FEHLER: Tab 'Output' nicht im aktiven Spreadsheet gefunden.");
    return;
  }

  const outputData = outputSheet.getDataRange().getValues();
  if (outputData.length <= 1) return;
  const headers = outputData[0];

  // Index-Ermittlung im Output-Sheet
  const phoneColIdx = headers.indexOf("Telefon") !== -1 ? headers.indexOf("Telefon") : 0; // Spalte A
  const rechnungStatusColIdx = headers.indexOf("Rechnung Status") !== -1 ? headers.indexOf("Rechnung Status") : 11; // Spalte L (Index 11)
  const emailStatusColIdx = headers.indexOf("Email Status") !== -1 ? headers.indexOf("Email Status") : headers.indexOf("Email status"); // Spalte I
  let driveBelegColIdx = headers.indexOf("Drive Zahlungsbeleg");
  if (driveBelegColIdx === -1) {
    driveBelegColIdx = 17; // Spalte R (Index 17)
    outputSheet.getRange(1, 18).setValue("Drive Zahlungsbeleg");
  }

  // Webhook-Sheet öffnen
  let waSs;
  try {
    waSs = SpreadsheetApp.openById(WA_WEBHOOK_SPREADSHEET_ID);
  } catch (e) {
    Logger.log("FEHLER beim Öffnen des Webhook-Sheets: " + e.toString());
    return;
  }

  const waSheet = waSs.getSheets()[0];
  const waData = waSheet.getDataRange().getValues();
  if (waData.length <= 1) return;

  const waHeaders = waData[0];
  let waCheckColIdx = waHeaders.indexOf("Zahlungsbeleg Check");
  if (waCheckColIdx === -1) waCheckColIdx = 19; // Spalte T (Index 19)

  const targetFolder = DriveApp.getFolderById(PAYMENT_PROOFS_FOLDER_ID);

  // 1. Webhook JSONs parsen (nur Zeilen, die noch kein 'x' in Spalte T haben)
  const waEvents = [];
  const processedWaRows = [];

  for (let i = 1; i < waData.length; i++) {
    const checkVal = String(waData[i][waCheckColIdx] || "").trim().toLowerCase();
    if (checkVal === "x") continue; // Bereits geprüfte Nachricht überspringen

    processedWaRows.push(i + 1); // Zeilennummer für spätere 'x'-Markierung merken

    const jsonStr = String(waData[i][6] || "").trim(); // Spalte G (Index 6)
    if (!jsonStr.startsWith("{")) continue;

    try {
      const jsonObj = JSON.parse(jsonStr);
      const extractedMedia = parseWhatsAppJsonMedia(jsonObj);
      if (extractedMedia) {
        waEvents.push(extractedMedia);
      }
    } catch (err) {
      // Ungültiges JSON ignorieren
    }
  }

  Logger.log(`Neue ungeprüfte WhatsApp-Zeilen: ${processedWaRows.length} | Davon Medienereignisse: ${waEvents.length}`);

  // 2. Zeilen im Output-Sheet durchgehen
  for (let r = 1; r < outputData.length; r++) {
    const row = outputData[r];
    const rawPhone = cleanPhoneNumber(row[phoneColIdx]);
    const rechnungStatus = String(row[rechnungStatusColIdx] || "").trim().toLowerCase();
    const emailStatusVal = row[emailStatusColIdx];
    const existingDriveUrlStr = String(row[driveBelegColIdx] || "").trim();

    // 🎯 FILTER 1: Nur ausführen, wenn Spalte L "offen" oder leer ist (nicht bei "bezahlt" oder "storniert")
    if (rechnungStatus !== "offen" && rechnungStatus !== "") {
      continue;
    }

    if (!rawPhone) continue;

    // Versanddatum der Rechnung / E-Mail ermitteln
    let sentDate = parseAnyTimestamp(emailStatusVal);
    if (!sentDate) {
      sentDate = new Date();
      sentDate.setHours(0, 0, 0, 0); // Fallback: Ab heute 00:00 Uhr
    }

    // Bereits gespeicherte Links in Array laden
    const savedUrls = existingDriveUrlStr ? existingDriveUrlStr.split(",").map(u => u.trim()).filter(u => u.length > 0) : [];

    for (let e = 0; e < waEvents.length; e++) {
      const event = waEvents[e];

      // Nummernabgleich & Zeitstempel-Prüfung
      if (checkPhoneMatch(rawPhone, event.senderPhone) && event.timestamp >= sentDate) {
        Logger.log(`Match für ${rawPhone}! Lade Datei via 360dialog [Media-ID: ${event.mediaId || 'URL'}]...`);
        try {
          // Download über 360dialog Proxy
          const blob = downloadMediaFrom360Dialog(event.mediaId, event.mediaUrl, d360ApiKey);
          if (blob) {
            const sizeInKb = blob.getBytes().length / 1024;
            if (sizeInKb >= 10) {
              const fileName = `WA_Beleg_${rawPhone}_${Utilities.formatDate(event.timestamp, Session.getScriptTimeZone(), "yyyyMMdd_HHmm")}_${event.filename}`;
              blob.setName(fileName);
              const newFile = targetFolder.createFile(blob);
              newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              const fileSavedUrl = newFile.getUrl();

              // 🎯 SAMMEL-LOGIK: Mehrere Dateien kommagetrennt speichern ohne Duplikate
              if (!savedUrls.includes(fileSavedUrl)) {
                savedUrls.push(fileSavedUrl);
                Logger.log(`SUCCESS: Beleg hinzugefügt: ${fileSavedUrl}`);
              }
            }
          }
        } catch (fetchErr) {
          Logger.log(`Fehler beim 360dialog Fetch für ${rawPhone}: ` + fetchErr.toString());
        }
      }
    }

    // 🎯 Falls neue Belege gefunden wurden, alle Links kommagetrennt in Spalte R schreiben
    if (savedUrls.length > 0) {
      const finalUrlString = savedUrls.join(", ");
      if (finalUrlString !== existingDriveUrlStr) {
        outputSheet.getRange(r + 1, driveBelegColIdx + 1).setValue(finalUrlString);
      }
    }
  }

  // 3. Alle soeben geprüften Webhook-Zeilen in Spalte T mit 'x' markieren
  if (processedWaRows.length > 0) {
    processedWaRows.forEach(rowNum => {
      waSheet.getRange(rowNum, waCheckColIdx + 1).setValue("x");
    });
    SpreadsheetApp.flush();
    Logger.log(`${processedWaRows.length} Webhook-Zeilen mit 'x' markiert.`);
  }

  Logger.log("=== WhatsApp-Scan beendet ===");
}

/**
 * 360dialog Download-Proxy
 */
function downloadMediaFrom360Dialog(mediaId, mediaUrl, d360ApiKey) {
  try {
    if (mediaId) {
      const getUrlApi = `https://waba-v2.360dialog.io/${mediaId}`;
      const response = UrlFetchApp.fetch(getUrlApi, {
        headers: { "D360-API-KEY": d360ApiKey },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() === 200) {
        const json = JSON.parse(response.getContentText());
        if (json.url) {
          const downloadUrl = json.url.replace("https://lookaside.fbsbx.com", "https://waba-v2.360dialog.io").replace(/\\/g, "");
          const fileRes = UrlFetchApp.fetch(downloadUrl, {
            headers: { "D360-API-KEY": d360ApiKey },
            muteHttpExceptions: true
          });
          if (fileRes.getResponseCode() === 200) return fileRes.getBlob();
        }
      }
    }

    if (mediaUrl) {
      const downloadUrl = mediaUrl.replace("https://lookaside.fbsbx.com", "https://waba-v2.360dialog.io").replace(/\\/g, "");
      const fileRes = UrlFetchApp.fetch(downloadUrl, {
        headers: { "D360-API-KEY": d360ApiKey },
        muteHttpExceptions: true
      });
      if (fileRes.getResponseCode() === 200) return fileRes.getBlob();
    }
  } catch (e) {
    Logger.log("Fehler beim 360dialog Medien-Download: " + e.toString());
  }
  return null;
}

/**
 * JSON-Parser für WhatsApp Webhooks
 */
function parseWhatsAppJsonMedia(jsonObj) {
  try {
    const entry = jsonObj.entry && jsonObj.entry[0];
    const change = entry && entry.changes && entry.changes[0];
    const val = change && change.value;
    if (!val) return null;

    const contact = val.contacts && val.contacts[0];
    const waId = contact ? String(contact.wa_id || "").replace(/\D/g, "") : "";

    // A) Inbound Messages
    if (val.messages && val.messages.length > 0) {
      const msg = val.messages[0];
      const sender = cleanPhoneNumber(msg.from || waId);
      const timestamp = new Date(parseInt(msg.timestamp, 10) * 1000);

      if (msg.type === "document" && msg.document) {
        return {
          senderPhone: sender,
          timestamp: timestamp,
          mediaId: msg.document.id,
          mediaUrl: msg.document.url,
          filename: msg.document.filename || "beleg.pdf"
        };
      } else if (msg.type === "image" && msg.image) {
        return {
          senderPhone: sender,
          timestamp: timestamp,
          mediaId: msg.image.id,
          mediaUrl: msg.image.url,
          filename: "beleg.jpg"
        };
      }
    }

    // B) Message Echoes
    if (val.message_echoes && val.message_echoes.length > 0) {
      const echo = val.message_echoes[0];
      const sender = cleanPhoneNumber(echo.to || echo.from || waId);
      const timestamp = new Date(parseInt(echo.timestamp, 10) * 1000);

      if (echo.type === "document" && echo.document) {
        return {
          senderPhone: sender,
          timestamp: timestamp,
          mediaId: echo.document.id,
          mediaUrl: echo.document.url,
          filename: echo.document.filename || "beleg.pdf"
        };
      } else if (echo.type === "image" && echo.image) {
        return {
          senderPhone: sender,
          timestamp: timestamp,
          mediaId: echo.image.id,
          mediaUrl: echo.image.url,
          filename: echo.image.filename || "beleg.jpg"
        };
      }
    }
  } catch (e) {
    return null;
  }
  return null;
}

function cleanPhoneNumber(val) {
  if (!val) return "";
  let str = String(val).replace(/\D/g, "");
  if (str.startsWith("00")) str = str.substring(2);
  if (str.startsWith("0") && str.length > 9) str = "49" + str.substring(1);
  return str;
}

function checkPhoneMatch(phone1, phone2) {
  if (!phone1 || !phone2) return false;
  if (phone1 === phone2) return true;
  if (phone1.length >= 8 && phone2.length >= 8) {
    return phone1.slice(-8) === phone2.slice(-8);
  }
  return false;
}

function parseAnyTimestamp(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const str = String(val).trim();
  const match = str.match(/\((\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})\)/);
  if (match) {
    return new Date(
      parseInt(match[3], 10),
      parseInt(match[2], 10) - 1,
      parseInt(match[1], 10),
      parseInt(match[4], 10),
      parseInt(match[5], 10)
    );
  }
  return null;
}