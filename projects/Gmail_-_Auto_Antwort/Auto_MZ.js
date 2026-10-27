function sendAutoResponseMonteurzimmer() {
  var scriptName = "Monteurzimmer.de Auto-Response";
  var cutoffDate = new Date("2026-10-03T12:57:00+02:00");
  var labelPath = "API_Buchung/MZ/aMZ";

  try {
    var searchQuery = 'from:monteurzimmer.de subject:"Buchungsanfrage für" -label:' + labelPath + ' (is:unread OR newer_than:7d)';
    var threads = GmailApp.search(searchQuery);
    if (threads.length === 0) return;

    var label = GmailApp.getUserLabelByName(labelPath) || GmailApp.createLabel(labelPath);
    var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
    var ss = SpreadsheetApp.openById(sheetId);

    // Bestehendes Tab "MZ"
    var mzSheet = ss.getSheetByName("MZ") || ss.insertSheet("MZ");
    if (mzSheet.getLastRow() === 0) {
      mzSheet.appendRow(["Zeit", "ID", "Name", "Telefonnummer", "Zeitraum", "Personen", "Titel", "Link", "Nachricht", "Adresse der Unterkunft", "JSON"]);
      mzSheet.getRange(1, 1, 1, 11).setFontWeight("bold");
    }

    var autoReplyText = "Hallo,\n\n" +
      "vielen Dank für Ihre Anfrage nach einer Monteurwohnung!\n" +
      "Wir prüfen die Anfrage jetzt und melden uns schnellstmöglich.\n\n" +
      "Für eine besonders schnelle Bearbeitung empfehlen wir Ihnen, uns direkt per WhatsApp Nachricht zu kontaktieren: +4917684801295\n\n" +
      "Sehr gerne können Sie bereits jetzt eine Übersicht der bei uns verfügbaren Wohnungen aus unserer Buchungswebsite a.L8Street.com erhalten. Sie können alle Wohnungen auch sofort auf der Website buchen, oder uns einen Link zu der gewünschten Wohnung per WhatsApp oder E-Mail zukommen lassen\n" +
      "Beste Grüße\n" +
      "Ihr Team von L8 Street";

    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var firstMessage = thread.getMessages()[0];
      if (firstMessage.getDate() < cutoffDate) continue;

      try {
        var bodyPlain = firstMessage.getPlainBody();
        var sender = firstMessage.getFrom();

        // Reply senden
        firstMessage.reply(autoReplyText);

        // --- PRÄZISES PARSING FÜR MONTEURZIMMER.DE ---
        var nameMatch = bodyPlain.match(/Anfrage von\s*:?\s*Herr\s*([^\r\n]+)/i) ||
                        bodyPlain.match(/Anfrage von\s*:?\s*Frau\s*([^\r\n]+)/i) ||
                        bodyPlain.match(/Anfrage von\s*:?\s*([^\r\n]+)/i);
        var cleanName = nameMatch ? nameMatch[1].trim() : "N/A";

        var periodMatch = bodyPlain.match(/An-\s*und\s*Abreise\s*:\s*([\d\.]+)\s*-\s*([\d\.]+)/i);
        var periodStr = periodMatch ? (periodMatch[1].trim() + " - " + periodMatch[2].trim()) : "N/A";

        var personsMatch = bodyPlain.match(/Anzahl Mieter\s*:\s*(\d+)/i) || bodyPlain.match(/Personen\s*:?\s*(\d+)/i);
        var personsCount = personsMatch ? personsMatch[1].trim() : "N/A";

        var phoneMatch = bodyPlain.match(/Telefon\s*:?\s*(\+?[\d\s\/|-]{8,})/i) || bodyPlain.match(/Mobil\s*:?\s*(\+?[\d\s\/|-]{8,})/i);
        var safePhone = phoneMatch ? "'" + phoneMatch[1].trim() : "N/A";
        var rawPhone = phoneMatch ? phoneMatch[1].trim() : "N/A";

        var titleMatch = bodyPlain.match(/Unterkunft\s*[\r\n]+([^(]+)/i) || firstMessage.getSubject().match(/für (.*)/i);
        var cleanTitle = titleMatch ? titleMatch[1].replace(/[\r\n]+/g, "").trim() : "N/A";

        var idMatch = bodyPlain.match(/\/eintrag\/([A-Za-z0-9]+)\?/i);
        var kundenNr = idMatch ? idMatch[1].trim() : "N/A";

        var msgMatch = bodyPlain.match(/Telefon[^\r\n]*[\r\n]+([\s\S]*?)(?=Jede Anfrage wird sicher|Mit freundlichen|Viele Grüße)/i);
        var cleanMsg = msgMatch ? msgMatch[1].replace(/[\r\n]{2,}/g, "\n").trim() : "Keine Nachricht";

        var threadLink = "https://mail.google.com/mail/u/0/#all/" + thread.getId();

        var senderEmailMatch = sender.match(/<([^>]+)>/);
        var cleanEmail = senderEmailMatch ? senderEmailMatch[1] : sender;

        // Einheitliches JSON-Schema
        var requestData = {
          portal: "monteurzimmer.de",
          kundenNummer: kundenNr,
          gastName: cleanName,
          telefonnummer: rawPhone,
          zeitraum: periodStr,
          personen: personsCount,
          unterkunftTitel: cleanTitle,
          link: threadLink,
          nachricht: cleanMsg,
          unterkunftAdresse: "N/A",
          email: cleanEmail,
          erfasstAm: new Date().toISOString()
        };

        var jsonString = JSON.stringify(requestData);

        // Im Tab "MZ" wie gewohnt eintragen (Schema A–K)
        mzSheet.appendRow([
          requestData.erfasstAm,
          requestData.kundenNummer,
          requestData.gastName,
          safePhone,
          requestData.zeitraum,
          requestData.personen,
          requestData.unterkunftTitel,
          requestData.link,
          requestData.nachricht,
          requestData.unterkunftAdresse,
          jsonString
        ]);

        thread.addLabel(label);
        logExecution(scriptName, "ERFOLG", "Antwort & Daten erfasst für Thread: " + cleanName, "");
      } catch (innerError) {
        logExecution(scriptName, "FEHLER", "Fehler beim Antworten an: " + firstMessage.getFrom(), innerError.stack || innerError.toString());
      }
    }
  } catch (globalError) {
    logExecution(scriptName, "SYSTEM_FEHLER", "Globaler Skriptfehler", globalError.stack || globalError.toString());
  }
}