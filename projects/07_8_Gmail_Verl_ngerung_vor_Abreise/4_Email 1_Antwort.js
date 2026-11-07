/**
 * SCRIPT 4: Prüft stündlich, ob der Kunde auf die gesendete E-Mail antwortet
 * und trägt den E-Mail-Text als JSON-String in Spalte S ein.
 */
function checkCustomerReplies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName('output');
  if (!outputSheet) {
    Logger.log('Fehler: Tabellenblatt "output" nicht gefunden.');
    return;
  }

  const lastRow = outputSheet.getLastRow();
  if (lastRow < 3) return;

  const dataRange = outputSheet.getRange(3, 1, lastRow - 2, 19); // Spalte A bis S
  const data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 3;

    const customerEmail = String(row[3]).trim().toLowerCase(); // Spalte D
    const statusQ = String(row[16]).trim(); // Spalte Q
    const emailOrThreadIdR = String(row[17]).trim(); // Spalte R (Email- oder Thread-ID)
    const existingReplyS = String(row[18]).trim(); // Spalte S

    // Abbrechen, wenn kein Status, bereits eine Antwort erfasst oder Mail ungültig ist
    if (!statusQ || existingReplyS !== '' || !customerEmail.includes('@')) {
      continue;
    }

    const sendDate = extractDateFromStatus_(statusQ);
    if (!sendDate) continue;

    let latestReplyMessage = null;

    // STRATEGIE 1: Direkt über die Thread-ID aus Spalte R suchen (falls vorhanden)
    if (emailOrThreadIdR) {
      try {
        let thread = null;
        // Versuche zuerst den Thread direkt zu laden
        try {
          thread = GmailApp.getThreadById(emailOrThreadIdR);
        } catch (e) {
          // Falls in R eine Message-ID statt Thread-ID steht:
          const msg = GmailApp.getMessageById(emailOrThreadIdR);
          if (msg) thread = msg.getThread();
        }

        if (thread) {
          latestReplyMessage = getLatestCustomerReplyFromThread_(thread, customerEmail, sendDate);
        }
      } catch (e) {
        Logger.log(`Zeile ${rowIndex}: Fehler beim Abruf per ID (${emailOrThreadIdR}): ${e.message}`);
      }
    }

    // STRATEGIE 2: Fallback über Gmail-Suche, falls per ID nichts gefunden wurde
    if (!latestReplyMessage) {
      const searchDateStr = Utilities.formatDate(sendDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      const query = `from:${customerEmail} after:${searchDateStr}`;
      const threads = GmailApp.search(query);

      for (const thread of threads) {
        const reply = getLatestCustomerReplyFromThread_(thread, customerEmail, sendDate);
        if (reply) {
          latestReplyMessage = reply;
          break; // Nimmt den neuesten passenden Thread
        }
      }
    }

    // Wenn eine Antwort gefunden wurde, als JSON-String in Spalte S eintragen
    if (latestReplyMessage) {
      const replyDateFormatted = Utilities.formatDate(
        latestReplyMessage.getDate(),
        Session.getScriptTimeZone(),
        'dd.MM.yyyy HH:mm:ss'
      );

      const payload = {
        date: replyDateFormatted,
        messageId: latestReplyMessage.getId(),
        body: latestReplyMessage.getPlainBody()
      };

      // Als JSON-String speichern (verhindert #ERROR! in Google Sheets)
      const cellContent = JSON.stringify(payload);

      outputSheet.getRange(rowIndex, 19).setValue(cellContent); // Spalte S
      Logger.log(`Antwort für ${customerEmail} in Zeile ${rowIndex} als JSON erfasst.`);
    }
  }
}

/**
 * Hilfsfunktion: Sucht die NEUESTE Kundenantwort innerhalb eines Threads
 */
function getLatestCustomerReplyFromThread_(thread, customerEmail, sendDate) {
  const messages = thread.getMessages();
  let newestReply = null;

  // Rückwärts iterieren, um die *neueste* Antwort zuerst zu prüfen
  for (let j = messages.length - 1; j >= 0; j--) {
    const msg = messages[j];
    const msgDate = msg.getDate();
    const msgFrom = msg.getFrom().toLowerCase();

    if (msgFrom.includes(customerEmail) && msgDate > sendDate) {
      newestReply = msg;
      break;
    }
  }

  return newestReply;
}

/**
 * HILFSFUNKTION: Parst das Datum aus dem Text in Spalte Q
 */
function extractDateFromStatus_(statusStr) {
  const match = statusStr.match(/(\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2})/);
  if (match && match[1]) {
    const parts = match[1].split(' ');
    const dateParts = parts[0].split('.');
    const timeParts = parts[1].split(':');

    return new Date(
      dateParts[2],
      dateParts[1] - 1,
      dateParts[0],
      timeParts[0],
      timeParts[1],
      timeParts[2]
    );
  }

  return null;
}