// =================================================================
// 1. TEIL: SELBSTHEILUNG (Nachträglicher Mail-Import)
// Gehört an den ANFANG der Pipeline (Schritt 1.1)
// =================================================================
function recoverMissingMMZEmails() {
  const SOURCE_SHEET_NAME = 'MMZ';
  const LABEL_NAME = '02_MMZ Archive';
  const FORWARD_TO = 'cc@l8street.com';
  const TIMEZONE = 'Europe/Zurich';
  const LOOKBACK_HOURS = 3;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) return;

  let label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) label = GmailApp.createLabel(LABEL_NAME);

  const now = new Date();
  const cutoffTime = new Date(now.getTime() - LOOKBACK_HOURS * 60 * 60 * 1000);
  const lastRowSource = sourceSheet.getLastRow();

  // 1. Bereits im Sheet vorhandene Links sammeln
  const registeredLinks = new Set();
  if (lastRowSource > 1) {
    const colAData = sourceSheet.getRange(2, 1, lastRowSource - 1, 1).getValues();
    for (let r = 0; r < colAData.length; r++) {
      const cell = String(colAData[r][0] || "").trim();
      if (cell.startsWith('{')) {
        try {
          const parsed = JSON.parse(cell);
          if (parsed.link) registeredLinks.add(parsed.link);
        } catch (e) {}
      }
    }
  }

  // 2. Gmail tiefenscannen
  const query = `from:buchungsanfrage@mein-monteurzimmer.de in:anywhere newer_than:${LOOKBACK_HOURS}h`;
  const threads = GmailApp.search(query, 0, 50);
  const missingMessages = [];
  const rowsToInsertDirectly = [];
  const threadsToArchive = new Set();
  let totalGmailMailsInWindow = 0;

  for (let t = 0; t < threads.length; t++) {
    const thread = threads[t];
    const messages = thread.getMessages();
    for (let m = 0; m < messages.length; m++) {
      const msg = messages[m];
      const mailDate = msg.getDate();
      if (mailDate.getTime() < cutoffTime.getTime()) continue;

      const rawSubject = msg.getSubject();
      const subjectLower = rawSubject.toLowerCase();
      if (subjectLower.startsWith('re:') || subjectLower.startsWith('aw:') || subjectLower.startsWith('fwd:') || subjectLower.startsWith('[mmz #')) continue;

      totalGmailMailsInWindow++;
      const messageId = msg.getId();
      const mailLink = `https://mail.google.com/mail/u/0/#all/${messageId}`;

      if (!registeredLinks.has(mailLink)) {
        missingMessages.push({
          messageId: messageId,
          date: mailDate.toISOString(),
          subject: rawSubject
        });

        const body = msg.getPlainBody();
        const htmlBody = msg.getBody();
        let inseratId = "";
        const inseratMatch = htmlBody.match(/mein-monteurzimmer\.de(?:%2F|\/)(\d+)(?:%2F|\/)monteurzimmer/i);
        if (inseratMatch && inseratMatch[1]) inseratId = inseratMatch[1].trim();

        const gastNameMatch = body.match(/Name\s*([^\r\n]+)/i);
        const firmaMatch = body.match(/Firma\s*([^\r\n]+)/i);
        let gastName = gastNameMatch ? gastNameMatch[1].trim() : "";
        let firma = firmaMatch ? firmaMatch[1].trim() : "";
        if (["keine angabe", "-", "null", "none"].includes(firma.toLowerCase())) firma = "";

        const telefonMatch = body.match(/Telefon\s*([^\r\n]+)/i);
        const emailMatch = body.match(/E-Mail\s*([^\r\n]+)/i);
        const gastEmail = emailMatch ? emailMatch[1].trim() : "";
        const zeitraumMatch = body.match(/Zeitraum\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4}\s*-\s*[0-9]{2}\.[0-9]{2}\.[0-9]{4}|[^\r\n]+?)(?=\s*(?:Personen|Zimmerart|Telefon|E-Mail|$|\r|\n))/i);

        let personen = "";
        const personenMatch = body.match(/Personen\s*([0-9]+)/i);
        if (personenMatch) {
          personen = personenMatch[1].trim();
        } else {
          const personenAlt = body.match(/Personen\s*([^\r\n]+?)(?=\s*(?:Zimmerart|Bitte|$|\r|\n))/i);
          if (personenAlt) personen = personenAlt[1].replace(/[^0-9]/g, '').trim();
        }

        let unterkunftTitel = "";
        const unterkunftMatch = body.match(/Unterkunft\s*[*„"“']?([^"“”'\r\n*]+)[*“"']?/i);
        if (unterkunftMatch) unterkunftTitel = unterkunftMatch[1].replace(/[*„"“']/g, '').trim();

        let nachricht = "Keine Nachricht";
        const nachrichtMatch = body.match(/NACHRICHT DES\/DER ZIMMERSUCHENDEN[\s\S]*?[„"“']([\s\S]*?)[“"']\s*(?:KONTAKTDATEN|Bitte|Name)/i);
        if (nachrichtMatch && nachrichtMatch[1].trim().length > 0) nachricht = nachrichtMatch[1].trim();
        nachricht = nachricht.replace(/[\r\n\t]+/g, ' ').replace(/\s{2,}/g, ' ');

        const anfrageUrlMatch = body.match(/https:\/\/mein-monteurzimmer\.de\/anfrage\/[a-zA-Z0-9_\-\/]+/);
        const zeitraumStr = zeitraumMatch ? zeitraumMatch[1].trim() : "";

        let uniqueToken = "";
        if (anfrageUrlMatch && anfrageUrlMatch[0]) {
          const urlParts = anfrageUrlMatch[0].split("/").filter(Boolean);
          uniqueToken = urlParts[urlParts.length - 1];
        }
        if (!uniqueToken) uniqueToken = Utilities.getUuid().slice(0, 8);

        const uniqueSubject = `${rawSubject} - [MMZ #${inseratId || 'ID'}-${uniqueToken}] [RECOVERED]`;
        const mailOptions = { htmlBody: htmlBody };
        if (gastEmail && !gastEmail.includes("••")) mailOptions.replyTo = gastEmail;

        let forwardMessageId = "";
        let forwardThreadId = "";
        try {
          GmailApp.sendEmail(FORWARD_TO, uniqueSubject, body, mailOptions);
          for (let attempt = 1; attempt <= 4; attempt++) {
            Utilities.sleep(1000);
            const createdThreads = GmailApp.search(`subject:"${uniqueToken}" to:${FORWARD_TO}`, 0, 1);
            if (createdThreads.length > 0) {
              forwardThreadId = createdThreads[0].getId();
              const msgs = createdThreads[0].getMessages();
              forwardMessageId = msgs[msgs.length - 1].getId();
              Logger.log(`[RECOVER-FORWARD-OK] Mail nacherfasst: Thread-ID ${forwardThreadId}`);
              break;
            }
          }
        } catch (e) {
          Logger.log("[RECOVER-FORWARD-FEHLER] " + e.toString());
        }

        const record = {
          portal: "mein-monteurzimmer.de",
          kundenNummer: inseratId,
          gastName: gastName,
          firma: firma,
          telefonnummer: telefonMatch ? telefonMatch[1].trim() : "",
          zeitraum: zeitraumStr,
          personen: personen,
          unterkunftTitel: unterkunftTitel,
          link: mailLink,
          forwardThreadId: forwardThreadId,
          forwardMessageId: forwardMessageId,
          nachricht: nachricht,
          unterkunftAdresse: "",
          email: gastEmail,
          erfasstAm: new Date().toISOString(),
          anfrageUrl: anfrageUrlMatch ? anfrageUrlMatch[0].trim() : "",
          anfrageErstelltAm: Utilities.formatDate(mailDate, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX")
        };

        rowsToInsertDirectly.push([JSON.stringify(record)]);
        registeredLinks.add(mailLink);
        threadsToArchive.add(thread);
      }
    }
  }

  // 3. Fehlende Mails direkt unten an Spalte A anhängen
  let recoveredCount = 0; // Hier VORHER deklariert, damit es immer existiert
  if (rowsToInsertDirectly.length > 0) {
    const startRow = sourceSheet.getLastRow() + 1;
    sourceSheet.getRange(startRow, 1, rowsToInsertDirectly.length, 1).setValues(rowsToInsertDirectly);
    recoveredCount = rowsToInsertDirectly.length;
    Logger.log(`[SELBSTHEILUNG] ${recoveredCount} fehlende Mail(s) in Spalte A eingetragen.`);
    for (const t of threadsToArchive) {
      t.addLabel(label);
      t.moveToArchive();
      t.markUnread();
    }
  }

  // Metriken für logMMZPipelineAudit() zwischenspeichern
  const recoveryStats = {
    totalGmailMailsInWindow: totalGmailMailsInWindow,
    recoveredCount: recoveredCount,
    missingMessages: missingMessages
  };
  PropertiesService.getScriptProperties().setProperty('MMZ_LAST_RECOVERY', JSON.stringify(recoveryStats));
}

// =================================================================
// 2. TEIL: AUDIT & HÄNGER-LOGGING (Tab "Kontrolle")
// Gehört an das ENDE der Pipeline (Schritt 10)
// =================================================================
function logMMZPipelineAudit() {
  const SOURCE_SHEET_NAME = 'MMZ';
  const AUDIT_SHEET_NAME = 'Kontrolle';
  const LOOKBACK_HOURS = 3;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) return;

  let auditSheet = ss.getSheetByName(AUDIT_SHEET_NAME);
  if (!auditSheet) auditSheet = ss.insertSheet(AUDIT_SHEET_NAME);

  if (!auditSheet.getRange(1, 1).getValue()) {
    auditSheet.getRange(1, 1).setValue("Audit Log (JSON)").setFontWeight("bold");
    auditSheet.getRange("A:A").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  const now = new Date();
  const cutoffTime = new Date(now.getTime() - LOOKBACK_HOURS * 60 * 60 * 1000);
  const currentLastRow = sourceSheet.getLastRow();
  const stuckRows = [];
  let rowsChecked = 0;

  if (currentLastRow > 1) {
    const rows = sourceSheet.getRange(2, 1, currentLastRow - 1, 8).getValues();
    for (let i = 0; i < rows.length; i++) {
      const rowNum = i + 2;
      const cellA = String(rows[i][0] || "").trim();
      const cellB = String(rows[i][1] || "").trim();
      const cellC = String(rows[i][2] || "").trim();
      const cellD = String(rows[i][3] || "").trim();
      const cellH = String(rows[i][7] || "").trim();

      if (!cellA || !cellA.startsWith('{')) continue;

      let rowTime = null;
      try {
        const obj = JSON.parse(cellA);
        if (obj.erfasstAm) rowTime = new Date(obj.erfasstAm);
      } catch (e) {}

      if (rowTime && rowTime.getTime() < cutoffTime.getTime()) continue;
      rowsChecked++;

      if (cellB === "-") continue; // Duplikate ignorieren

      const isStuckAtB = (cellB === "");
      const isStuckAtC = (cellB !== "" && cellC === "");
      const isStuckAtD = (cellC !== "" && !cellC.startsWith("Fehler") && cellD === "");
      const isStuckAtH = (cellD !== "" && cellH === "");

      if (isStuckAtB || isStuckAtC || isStuckAtD || isStuckAtH) {
        stuckRows.push({
          row: rowNum,
          stuckAt: isStuckAtB ? "B_DUPLIKAT" : isStuckAtC ? "C_GEMINI" : isStuckAtD ? "D_VALIDATOR" : "H_DRAFT"
        });
      }
    }
  }

  // Selbstheilungs-Statistiken aus Schritt 1 abrufen
  const props = PropertiesService.getScriptProperties();
  const storedStatsRaw = props.getProperty('MMZ_LAST_RECOVERY');
  let recoveryStats = { totalGmailMailsInWindow: 0, recoveredCount: 0, missingMessages: [] };
  if (storedStatsRaw) {
    try {
      recoveryStats = JSON.parse(storedStatsRaw);
    } catch (e) {}
    // Nach dem Auslesen löschen, damit keine Altlasten bleiben
    props.deleteProperty('MMZ_LAST_RECOVERY');
  }

  // Audit Log mit vollständigen Metriken erstellen
  const auditLog = {
    timestamp: now.toISOString(),
    status: (recoveryStats.missingMessages.length === 0 && stuckRows.length === 0) ? "OK" : "WARNUNG",
    zeitfensterStunden: LOOKBACK_HOURS,
    metrics: {
      gmailMailsImFenster: recoveryStats.totalGmailMailsInWindow,
      gepruefteZeilenImSheet: rowsChecked,
      gefundeneFehlendeMails: recoveryStats.missingMessages.length,
      erfolgreichNachtraeglichEingefuegt: recoveryStats.recoveredCount,
      haengendeZeilen: stuckRows.length
    },
    details: {
      fehlendeMails: recoveryStats.missingMessages,
      haengendeZeilen: stuckRows
    }
  };

  const nextAuditRow = auditSheet.getLastRow() + 1;
  auditSheet.getRange(nextAuditRow, 1).setValue(JSON.stringify(auditLog));
  auditSheet.setRowHeight(nextAuditRow, 21);
  Logger.log(`[AUDIT] Kontrolle abgeschlossen. Geprüft: ${rowsChecked}, Hänger: ${stuckRows.length}, Nachimportiert: ${recoveryStats.recoveredCount}`);
}