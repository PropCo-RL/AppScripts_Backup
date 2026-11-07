// ID deines Ziel-Ordners in Google Drive für hochgeladene Belege
const PAYMENT_PROOFS_FOLDER_ID = "1Xh84O1T8fiGCcToOk-zszLzjBGcgfAHv";

function scanGmailForPaymentProofs() {
  Logger.log("Starte Suche nach Zahlungsnachweisen in Gmail...");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName("Output");
  if (!outputSheet) {
    Logger.log("Fehler: Tab 'Output' nicht gefunden.");
    return;
  }

  const data = outputSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];

  // Dynamic Index Matching basierend auf deinen Spaltennamen
  const emailColIdx = headers.indexOf("Email") !== -1 ? headers.indexOf("Email") : 2; // Spalte C
  const emailStatusColIdx = headers.indexOf("Email Status") !== -1 ? headers.indexOf("Email Status") : headers.indexOf("Email status"); // Spalte I

  // Spalte R ("Drive Zahlungsbeleg") - Erstellt die Überschrift falls nicht vorhanden
  let driveBelegColIdx = headers.indexOf("Drive Zahlungsbeleg");
  if (driveBelegColIdx === -1) {
    driveBelegColIdx = 17; // Spalte R (Index 17)
    outputSheet.getRange(1, 18).setValue("Drive Zahlungsbeleg");
  }

  const targetFolder = DriveApp.getFolderById(PAYMENT_PROOFS_FOLDER_ID);

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const email = String(row[emailColIdx] || "").trim().toLowerCase();
    const emailStatusVal = String(row[emailStatusColIdx] || "").trim();
    const existingDriveUrl = String(row[driveBelegColIdx] || "").trim();

    // Nur prüfen, wenn eine E-Mail vorliegt und noch kein Beleg in Spalte R gespeichert ist
    if (!email || existingDriveUrl !== "") continue;

    // Zeitstempel der gesendeten Rechnung aus Spalte I auslesen
    const sentDate = parseEmailStatusTimestamp(emailStatusVal);
    if (!sentDate) continue; // Wurde noch keine Rechnung gesendet, abbrechen

    // Suche in Gmail nach Nachrichten von dieser E-Mail-Adresse
    const query = `from:${email} has:attachment`;
    const threads = GmailApp.search(query, 0, 5); // Max. 5 Threads durchsuchen

    let fileSavedUrl = "";

    for (let t = 0; t < threads.length; t++) {
      const messages = threads[t].getMessages();
      for (let m = 0; m < messages.length; m++) {
        const msg = messages[m];
        const msgDate = msg.getDate();
        const msgFrom = msg.getFrom().toLowerCase();

        // 1. NEU: Eigene E-Mails von l8street.com komplett ignorieren
        if (msgFrom.includes("@l8street.com")) {
          continue;
        }

        // 2. Nur E-Mails verarbeiten, die NACH dem Rechnungs-Versanddatum eingetroffen sind
        if (msgDate > sentDate) {
          const attachments = msg.getAttachments();
          for (let a = 0; a < attachments.length; a++) {
            const att = attachments[a];
            const sizeInKb = att.getBytes().length / 1024;
            const mimeType = att.getContentType().toLowerCase();

            // Kriterium: Mindestens 60 KB & gängige Dateiformate (PDF, JPEG, PNG)
            const isSupportedType = mimeType.includes("pdf") || mimeType.includes("image");
            if (sizeInKb >= 60 && isSupportedType) {
              const fileName = `Mail_Beleg_${email}_${Utilities.formatDate(msgDate, Session.getScriptTimeZone(), "yyyyMMdd_HHmm")}_${att.getName()}`;
              const newFile = targetFolder.createFile(att.copyBlob().setName(fileName));
              newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

              fileSavedUrl = newFile.getUrl();
              Logger.log(`--> Beleg von Absender [${msgFrom}] gespeichert: ${fileSavedUrl}`);
              break;
            }
          }
        }
        if (fileSavedUrl) break;
      }
      if (fileSavedUrl) break;
    }

    // Wenn Beleg gefunden und gespeichert wurde, Link in Spalte R schreiben
    if (fileSavedUrl) {
      outputSheet.getRange(r + 1, driveBelegColIdx + 1).setValue(fileSavedUrl);
    }
  }
  Logger.log("Gmail-Scan abgeschlossen.");
}

/**
 * Hilfsfunktion zum Auslesen des Sendedatums aus dem "Email Status"-String
 * Beispiel: "Gesendet an info@donigmbh.de (21.10.2026 14:51)"
 */
function parseEmailStatusTimestamp(statusText) {
  if (!statusText) return null;
  const match = statusText.match(/\((\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})\)/);
  if (!match) return null;
  return new Date(
    parseInt(match[3], 10),
    parseInt(match[2], 10) - 1,
    parseInt(match[1], 10),
    parseInt(match[4], 10),
    parseInt(match[5], 10)
  );
}