function verarbeiteMMZEmails() {
  const SHEET_NAME = 'MMZ';
  const LABEL_NAME = '02_MMZ Archive';
  const FORWARD_TO = 'cc@l8street.com'; // Zieladresse für getrennte Threads
  const TIMEZONE = 'Europe/Zurich';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Das Tabellenblatt "${SHEET_NAME}" wurde nicht gefunden!`);
    return;
  }

  let label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) {
    label = GmailApp.createLabel(LABEL_NAME);
  }

  // 1. Bereits importierte Links auslesen (Duplikatschutz)
  const existingLinks = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    const existingData = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (let r = 0; r < existingData.length; r++) {
      const cell = existingData[r][0];
      if (typeof cell === 'string' && cell.startsWith('{')) {
        try {
          const parsed = JSON.parse(cell);
          if (parsed.link) existingLinks.add(parsed.link);
        } catch(e) {}
      }
    }
  }

  // 2. Zeitfenster: Gestern ab 23:00:00 Uhr bis jetzt
  const now = new Date();
  const yesterday = new Date(now.getTime() - (24 * 60 * 60 * 1000));
  const yesterdaySearchStr = Utilities.formatDate(yesterday, TIMEZONE, "yyyy/MM/dd");
  
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 1);
  cutoffDate.setHours(23, 0, 0, 0);
  
  // Suche OHNE in:inbox, um auch weitergeführte Sammel-Threads im Archiv zu erfassen
  const searchQuery = `from:buchungsanfrage@mein-monteurzimmer.de after:${yesterdaySearchStr}`;
  const threads = GmailApp.search(searchQuery, 0, 100);

  const rowsToInsert = [];
  const processedThreads = new Set();

  for (let t = 0; t < threads.length; t++) {
    const thread = threads[t];
    const messages = thread.getMessages();
    let threadHadNewMessage = false;

    for (let m = 0; m < messages.length; m++) {
      const msg = messages[m];
      const sender = msg.getFrom().toLowerCase();
      if (!sender.includes('buchungsanfrage@mein-monteurzimmer.de')) continue;

      const rawSubject = msg.getSubject();
      const subjectLower = rawSubject.toLowerCase();
      
      // Eigene Antworten, Weiterleitungen und bereits separierte MMZ-Mails ignorieren
      if (subjectLower.startsWith('re:') || subjectLower.startsWith('aw:') || subjectLower.startsWith('fwd:') || subjectLower.startsWith('[mmz #')) continue;

      const mailDate = msg.getDate();
      if (mailDate.getTime() < cutoffDate.getTime()) continue;

      const messageId = msg.getId();
      const mailLink = `https://mail.google.com/mail/u/0/#all/${messageId}`;
      if (existingLinks.has(mailLink)) continue;

      const body = msg.getPlainBody();
      const htmlBody = msg.getBody();

      // Original-Postfach aus To-Header erfassen
      const originalTo = msg.getTo();

      // MMZ Postmark-Message-ID aus den Raw-Headern holen (sofern vorhanden)
      let mmzMessageId = "";
      try {
        const rawContent = msg.getRawContent();
        const pmMatch = rawContent.match(/X-PM-Message-Id:\s*([^\r\n]+)/i);
        if (pmMatch && pmMatch[1]) {
          mmzMessageId = pmMatch[1].trim();
        }
      } catch (e) {}

      // Inserats-ID aus dem HTML-Link extrahieren
      let inseratId = "";
      const inseratMatch = htmlBody.match(/mein-monteurzimmer\.de(?:%2F|\/)(\d+)(?:%2F|\/)monteurzimmer/i);
      if (inseratMatch && inseratMatch[1]) {
        inseratId = inseratMatch[1].trim();
      }

      // Titel der Unterkunft aus Plaintext holen
      let unterkunftTitel = "";
      const unterkunftMatch = body.match(/Unterkunft\s*[*„"“']?([^"\r\n*]+)[*“"']?/i);
      if (unterkunftMatch) {
        unterkunftTitel = unterkunftMatch[1].replace(/[*„"“']/g, '').trim();
      }

      // --- STANDORT-EXTRAKTION (Beide Methoden) ---
      // Methode 1: Aus dem HTML-Slug
      let unterkunftStadtHtml = "";
      const slugMatch = htmlBody.match(/monteurzimmer(?:%2F|\/)([a-z0-9\-]+?)-(?:monteur|zimmer|wohnung|praktische|einfache|[a-z0-9\-]+)/i);
      if (slugMatch && slugMatch[1]) {
        const cleanSlug = slugMatch[1].replace(/-/g, ' ').trim();
        unterkunftStadtHtml = cleanSlug.charAt(0).toUpperCase() + cleanSlug.slice(1);
      }

      // Methode 2: Aus dem Unterkunftstitel
      let unterkunftStadtTitel = "";
      if (unterkunftTitel) {
        const inMatch = unterkunftTitel.match(/(?:in|bei|\-)\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)?)/i);
        if (inMatch) {
          unterkunftStadtTitel = inMatch[1].replace(/\s*-\s*L8.*$/i, '').trim();
        } else {
          // Prüfen ob Titel direkt ein bekanntes Muster wie "Monteurwohnung Reutlingen" hat
          const directMatch = unterkunftTitel.match(/Monteurwohnung(?:en)?\s+([A-ZÄÖÜ][a-zäöüß]+)/i);
          if (directMatch) {
            unterkunftStadtTitel = directMatch[1].trim();
          }
        }
      }

      // Konsolidierte Adresse
      const unterkunftAdresse = unterkunftStadtHtml || unterkunftStadtTitel || "";

      // --- KONTAKTDATEN-ABSCHNITT ISOLIEREN ---
      // Verhindert, dass Text aus der Gästenachricht Datums- oder Namens-Felder überschreibt
      let kontaktSektion = body;
      const kontaktSplit = body.split(/KONTAKTDATEN DES\/DER ZIMMERSUCHENDEN/i);
      if (kontaktSplit.length > 1) {
        kontaktSektion = kontaktSplit[1].split(/Bitte beantworten Sie|Buchungsanfrage ansehen/i)[0];
      }

      // 1. Name & Firma
      const gastNameMatch = kontaktSektion.match(/Name\s*([^\r\n]+)/i);
      const firmaMatch = kontaktSektion.match(/Firma\s*([^\r\n]+)/i);
      let gastName = gastNameMatch ? gastNameMatch[1].trim() : "";
      let firma = firmaMatch ? firmaMatch[1].trim() : "";
      if (["keine angabe", "-", "null", "none"].includes(firma.toLowerCase())) {
        firma = "";
      }
      const saubererGastName = gastName;

      // 2. Telefon & E-Mail mit Maskierungs-Schutz (••• wird ignoriert)
      const telefonMatch = kontaktSektion.match(/Telefon\s*([^\r\n]+)/i);
      let gastTelefon = telefonMatch ? telefonMatch[1].trim() : "";
      if (gastTelefon.includes("••") || gastTelefon.includes("nach Klick") || gastTelefon.includes("&bull;")) {
        gastTelefon = "";
      }

      const emailMatch = kontaktSektion.match(/E-Mail\s*([^\r\n]+)/i);
      let gastEmail = emailMatch ? emailMatch[1].trim() : "";
      if (gastEmail.includes("••") || gastEmail.includes("nach Klick") || gastEmail.includes("&bull;")) {
        gastEmail = "";
      }

      // 3. Zeitraum exakt aus Kontakttabelle
      const zeitraumMatch = kontaktSektion.match(/Zeitraum\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4}\s*-\s*[0-9]{2}\.[0-9]{2}\.[0-9]{4})/i);
      let zeitraumStr = "";
      if (zeitraumMatch) {
        zeitraumStr = zeitraumMatch[1].trim();
      } else {
        const zeitraumAlt = kontaktSektion.match(/Zeitraum\s*([^\r\n]+?)(?=\s*(?:Personen|Zimmerart|$|\r|\n))/i);
        if (zeitraumAlt) zeitraumStr = zeitraumAlt[1].trim();
      }

      // 4. Personen
      let personen = "";
      const personenMatch = kontaktSektion.match(/Personen\s*([0-9]+)/i);
      if (personenMatch) {
        personen = personenMatch[1].trim();
      }

      // 5. Zimmerart
      let zimmerart = "";
      const zimmerartMatch = kontaktSektion.match(/Zimmerart\s*([^\r\n]+)/i);
      if (zimmerartMatch) {
        zimmerart = zimmerartMatch[1].trim();
        if (zimmerart.toLowerCase() === "keine angabe") zimmerart = "";
      }

      // 6. Nachricht des Gastes (aus dem Freitextbereich)
      let nachricht = "Keine Nachricht";
      const nachrichtMatch = body.match(/NACHRICHT DES\/DER ZIMMERSUCHENDEN[\s\S]*?[„"“']([\s\S]*?)[“"']\s*(?:KONTAKTDATEN|Bitte|Name)/i);
      if (nachrichtMatch && nachrichtMatch[1].trim().length > 0) {
        nachricht = nachrichtMatch[1].trim();
      }
      nachricht = nachricht.replace(/[\r\n\t]+/g, ' ').replace(/\s{2,}/g, ' ');

      // 7. Anfrage-URL & Eindeutiges Token
      const anfrageUrlMatch = body.match(/https:\/\/mein-monteurzimmer\.de\/anfrage\/[a-zA-Z0-9_\-\/]+/);
      let uniqueToken = "";
      if (anfrageUrlMatch && anfrageUrlMatch[0]) {
        const urlParts = anfrageUrlMatch[0].split("/").filter(Boolean);
        uniqueToken = urlParts[urlParts.length - 1];
      }
      if (!uniqueToken) {
        uniqueToken = Utilities.getUuid().slice(0, 8);
      }
      const forwardSubject = `${rawSubject} - [MMZ #${inseratId || 'ID'}-${uniqueToken}]`;

      // 8. Weiterleiten & Thread erfassen
      let forwardMessageId = "";
      let forwardThreadId = "";
      try {
        const mailOptions = {
          htmlBody: htmlBody
        };
        if (gastEmail) {
          mailOptions.replyTo = gastEmail;
        }

        GmailApp.sendEmail(FORWARD_TO, forwardSubject, body, mailOptions);

        // Retry-Polling zur Erfassung der Thread-ID
        for (let attempt = 1; attempt <= 4; attempt++) {
          Utilities.sleep(1000);
          const createdThreads = GmailApp.search(`subject:"${uniqueToken}" to:${FORWARD_TO}`, 0, 1);
          if (createdThreads.length > 0) {
            forwardThreadId = createdThreads[0].getId();
            const msgs = createdThreads[0].getMessages();
            forwardMessageId = msgs[msgs.length - 1].getId();
            Logger.log(`[FORWARD-OK] Mail weitergeleitet an ${FORWARD_TO} (ID: ${forwardMessageId} | Thread-ID: ${forwardThreadId})`);
            break;
          }
        }

        if (!forwardThreadId) {
          Logger.log(`[WARNUNG] Mail gesendet, aber Thread für "${uniqueToken}" noch nicht indexiert.`);
        }
      } catch (err) {
        Logger.log(`[FORWARD-FEHLER] ❌ Konnte E-Mail nicht senden: ${err.toString()}`);
      }

      // Datensatz vorbereiten
      const record = {
        portal: "mein-monteurzimmer.de",
        kundenNummer: inseratId,
        gastName: saubererGastName,
        firma: firma,
        telefonnummer: gastTelefon,
        email: gastEmail,
        zeitraum: zeitraumStr,
        personen: personen,
        zimmerart: zimmerart,
        unterkunftTitel: unterkunftTitel,
        unterkunftStadtHtml: unterkunftStadtHtml,
        unterkunftStadtTitel: unterkunftStadtTitel,
        unterkunftAdresse: unterkunftAdresse,
        nachricht: nachricht,
        link: mailLink,
        forwardThreadId: forwardThreadId,
        forwardMessageId: forwardMessageId,
        anfrageUrl: anfrageUrlMatch ? anfrageUrlMatch[0].trim() : "",
        originalTo: originalTo,
        mmzMessageId: mmzMessageId,
        anfrageErstelltAm: Utilities.formatDate(mailDate, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX"),
        erfasstAm: new Date().toISOString()
      };

      rowsToInsert.push([JSON.stringify(record)]);
      existingLinks.add(mailLink);
      threadHadNewMessage = true;
    }

    if (threadHadNewMessage) {
      processedThreads.add(thread);
    }
  }

  // 3. Neue Datensätze eintragen, Threads archivieren und labeln
  if (rowsToInsert.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToInsert.length, 1).setValues(rowsToInsert);

    for (const t of processedThreads) {
      t.addLabel(label);
      t.moveToArchive();
      t.markUnread();
    }
    Logger.log(`${rowsToInsert.length} neue Anfrage(n) erfasst, getrennt weitergeleitet und archiviert.`);
  }
}