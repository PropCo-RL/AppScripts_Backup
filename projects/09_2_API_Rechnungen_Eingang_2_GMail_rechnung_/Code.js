function processInvoicesAndPreFilter() {
  Logger.log("=== START: GMAIL DOWNLOAD & VORFILTERUNG ===");

  // KONFIGURATION
  const SEARCH_LABEL = "support@-rechnung@";
  const PROCESSED_LABEL = "support@-rechnung@/weitergeleitetAnLx";
  const DRIVE_FOLDER_ID = "1j8xPlrOZTFP205HG9EPpGxiQYV-UI6rm";
  const SHEET_NAME = "API Mail Eingang";
  const MY_DOMAIN = "l8street.com";
  const MIN_FILE_SIZE_KB = 15;

  // 1. Folgelabel abrufen oder erstellen
  let processedLabelObj = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!processedLabelObj) {
    processedLabelObj = GmailApp.createLabel(PROCESSED_LABEL);
    Logger.log("Neues Folgelabel '" + PROCESSED_LABEL + "' wurde erstellt.");
  }

  // 2. Drive & Sheet laden
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.appendRow(["Datum / Zeit", "Absender", "Dateiname", "Link zur Datei", "UniqueKey"]);
  } else {
    if (!sheet.getRange(1, 5).getValue()) {
      sheet.getRange(1, 5).setValue("UniqueKey");
    }
  }

  // Hilfsfunktion: Bereinigt "Name <mail@domain.de>" zu "mail@domain.de"
  const extractCleanEmail = (rawStr) => {
    if (!rawStr) return "";
    const match = String(rawStr).match(/<([^>]+)>/);
    return (match ? match[1] : String(rawStr)).trim().toLowerCase();
  };

  // Hilfsfunktion: Datum formatiert als YYYY-MM-DD
  const formatIsoDate = (rawDate) => {
    try {
      if (rawDate instanceof Date) {
        return rawDate.toISOString().slice(0, 10);
      }
      const parsed = new Date(rawDate);
      if (!isNaN(parsed.getTime())) {
        return parsed.toISOString().slice(0, 10);
      }
    } catch (e) {}
    return String(rawDate).slice(0, 10);
  };

  // Eindeutige Duplikat-Schlüssel laden (inkl. abwärtskompatibler Altdaten-Normalisierung)
  const existingKeys = new Set();
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    dataRange.forEach(row => {
      const explicitKey = row[4]; // Spalte E
      if (explicitKey) {
        existingKeys.add(String(explicitKey).toLowerCase());
      }
      
      // Fallback-Schlüssel für Bestandsdaten und zusätzliche Absicherung (ohne Dateigröße)
      const cleanOldSender = extractCleanEmail(row[1]);
      const oldDateStr = formatIsoDate(row[0]);
      const oldFileName = String(row[2]).trim().toLowerCase();
      
      const fallbackKey = `${cleanOldSender}|${oldDateStr}|${oldFileName}`;
      existingKeys.add(fallbackKey);
    });
  }

  // 3. Gezielte Suche im Rechnungs-Label (limitiert auf 50 Threads gegen Script-Timeouts)
  const query = `label:${SEARCH_LABEL} -label:"${PROCESSED_LABEL}" has:attachment`;
  const threads = GmailApp.search(query, 0, 50);

  let savedCount = 0;
  let skippedCount = 0;

  // 4. Threads verarbeiten
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();

    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      const rawSender = message.getFrom();
      const rawReplyTo = message.getReplyTo();
      const cleanSender = extractCleanEmail(rawSender);
      const cleanReplyTo = extractCleanEmail(rawReplyTo);
      const subject = message.getSubject().toLowerCase();
      const toRecipient = extractCleanEmail(message.getTo());
      const ccRecipient = extractCleanEmail(message.getCc());
      const msgDateStr = formatIsoDate(message.getDate());

      // SCHUTZ 1: Eigene gesendete Mails überspringen
      if (cleanSender.includes(MY_DOMAIN)) {
        skippedCount++;
        continue;
      }

      // SCHUTZ 2: Kalender-Antworten und Termin-Mails ausfiltern
      const isCalendarMail = subject.startsWith("accepted:") || 
                             subject.startsWith("declined:") || 
                             subject.startsWith("tentative:") || 
                             subject.startsWith("invitation:") ||
                             subject.startsWith("aktualisiert:") ||
                             subject.startsWith("abgesagt:");
      if (isCalendarMail) {
        Logger.log(`[AUSSORTIERT] Kalender-Statusnachricht: ${subject} (${cleanSender})`);
        skippedCount++;
        continue;
      }

      // SCHUTZ 3: Info- & DMARC-System-Mails ignorieren (Rechnungsdienste wie Lexware/Sevdesk ausnehmen)
      const isBillingService = cleanSender.includes("lexware.de") || cleanSender.includes("sevdesk") || cleanSender.includes("fastbill");
      const isSystemReport = cleanSender.includes("dmarc") || cleanSender.includes("noreply") || subject.includes("report domain");
      const isInfoOnly = (toRecipient.startsWith("info@") || ccRecipient.startsWith("info@")) && 
                         !toRecipient.includes("rechnung@") && !ccRecipient.includes("rechnung@");

      if (!isBillingService && (isSystemReport || isInfoOnly)) {
        Logger.log(`[AUSSORTIERT] Info-/System-Mail: ${subject} (${cleanSender})`);
        skippedCount++;
        continue;
      }

      const attachments = message.getAttachments();
      for (let k = 0; k < attachments.length; k++) {
        const attachment = attachments[k];
        const fileName = attachment.getName().toLowerCase();
        const mimeType = attachment.getContentType().toLowerCase();
        const fileSizeKB = attachment.getSize() / 1024;

        // SCHUTZ 4: Mindestgröße prüfen (ab 15 KB)
        if (fileSizeKB < MIN_FILE_SIZE_KB) {
          Logger.log(`[AUSSORTIERT] Anhang zu klein (${fileSizeKB.toFixed(2)} KB): '${attachment.getName()}'`);
          skippedCount++;
          continue;
        }

        // SCHUTZ 5: Formatprüfung & Stördateien (inkl. rechnung_re- und mahnung_re-)
        const isCalendar = fileName.endsWith(".ics") || mimeType.includes("calendar");
        const isZip = fileName.endsWith(".zip") || mimeType.includes("zip");
        const isImage = mimeType.startsWith("image/") || 
                        fileName.endsWith(".png") || fileName.endsWith(".jpg") || 
                        fileName.endsWith(".jpeg") || fileName.endsWith(".webp") || 
                        fileName.endsWith(".gif");

        const isLegalOrNotice = fileName.includes("datenschutz") ||
                                fileName.includes("widerruf") ||
                                fileName.includes("wr-belehrung") ||
                                fileName.includes("belehrung") ||
                                fileName.includes("agb") ||
                                fileName.includes("anleitung") ||
                                fileName.includes("merkblatt");

        const isXmlOrOwnPattern = fileName.includes("anfrage.xml") ||
                                  fileName.endsWith(".xml") ||
                                  fileName.includes("rechnung_re-") ||
                                  fileName.includes("mahnung_re-");

        if (isCalendar || isZip || isImage || isLegalOrNotice || isXmlOrOwnPattern) {
          Logger.log(`[AUSSORTIERT] Unerwünschter Anhang/Format: '${attachment.getName()}'`);
          skippedCount++;
          continue;
        }

        // SCHUTZ 6: Erlaubte Formate (PDF + Word + Excel)
        const isPdf = mimeType.includes("pdf") || fileName.endsWith(".pdf");
        const isDoc = fileName.endsWith(".docx") || fileName.endsWith(".doc") || fileName.endsWith(".xlsx");

        if (!isPdf && !isDoc) {
          Logger.log(`[AUSSORTIERT] Kein PDF/Office: '${attachment.getName()}'`);
          skippedCount++;
          continue;
        }

        // SCHUTZ 7: Duplikat-Schutz (Prüft Key mit Dateigröße und Basis-Key)
        const effectiveSender = (cleanReplyTo.includes("@") && isBillingService) ? cleanReplyTo : cleanSender;
        const uniqueKey = `${effectiveSender}|${msgDateStr}|${fileName}|${Math.round(fileSizeKB)}`;
        const fallbackKey = `${effectiveSender}|${msgDateStr}|${fileName}`;

        if (existingKeys.has(uniqueKey) || existingKeys.has(fallbackKey)) {
          Logger.log(`[DUPLIKAT] Datei bereits importiert: '${attachment.getName()}' von ${effectiveSender}`);
          skippedCount++;
          continue;
        }

        // Speichern in Google Drive
        const file = folder.createFile(attachment);

        // Anzeige-Absender für das Sheet (dokumentiert Reply-To bei Lexware etc.)
        const displaySender = (rawReplyTo && cleanReplyTo !== cleanSender) 
          ? `${rawSender} (Antwort an: ${rawReplyTo})` 
          : rawSender;

        // In Google Sheet eintragen
        sheet.appendRow([
          message.getDate(),
          displaySender,
          file.getName(),
          file.getUrl(),
          uniqueKey
        ]);

        existingKeys.add(uniqueKey);
        existingKeys.add(fallbackKey);
        savedCount++;
      }
    }

    // Thread nach Durchlauf aller Nachrichten und Anhänge immer als verarbeitet kennzeichnen
    threads[i].addLabel(processedLabelObj);
  }

  Logger.log(`=== ENDE: ${savedCount} Dateien gespeichert, ${skippedCount} Mails/Anhänge aussortiert ===`);
}