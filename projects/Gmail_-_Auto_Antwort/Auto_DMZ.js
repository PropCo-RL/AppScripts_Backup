function sendAutoResponseDMZ() {
  var scriptName = "DMZ Auto-Response";
  var cutoffDate = new Date("2026-10-25T12:57:00+02:00");
  var labelPath = "API_Buchung/DMZ/aDMZ";
  var processedAny = false;

  try {
    var searchQuery = 'from:mail.dmz.de (subject:"Buchungsanfrage für" OR subject:"Neue Anfrage für") -label:' + labelPath + ' (is:unread OR newer_than:7d)';
    var threads = GmailApp.search(searchQuery);

    if (threads.length === 0) return;

    var label = GmailApp.getUserLabelByName(labelPath) || GmailApp.createLabel(labelPath);
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
        var actualSender = firstMessage.getFrom();
        var cleanEmail = actualSender.match(/<([^>]+)>/) ? actualSender.match(/<([^>]+)>/)[1] : actualSender;
        var subject = firstMessage.getSubject();

        // Sicherheits-Check: Nur echte Buchungsanfragen verarbeiten
        var isRealInquiry = bodyPlain.indexOf("Neue Anfrage für Ihren Eintrag") !== -1 ||
                            bodyPlain.indexOf("Sie haben eine neue Buchungsanfrage") !== -1;

        if (!isRealInquiry || cleanEmail.toLowerCase().indexOf("l8street.com") !== -1) {
          thread.addLabel(label);
          continue;
        }

        // Auto-Response senden
        if (!subject.toLowerCase().startsWith("re:")) {
          subject = "Re: " + subject;
        }
        GmailApp.sendEmail(cleanEmail, subject, autoReplyText, { replyTo: cleanEmail });

        // Daten extrahieren
        var kdNrMatch = bodyPlain.match(/Kundennummer\s*:\s*\*?(\d+)\*?/i);
        var nameMatch = bodyPlain.match(/Absender\s*:\s*\*?([^*|\r|\n]+)\*?/i) || bodyPlain.match(/Anfrage von\s*\*?([^*|\r|\n]+)\*?/i);
        var periodMatch = bodyPlain.match(/Zeitraum\s*:\s*von\s*\*?([\d\.]+)\*?\s*bis\s*\*?([\d\.]+)\*?/i);
        var periodStr = periodMatch ? (periodMatch[1] + " - " + periodMatch[2]) : "N/A";
        var personsMatch = bodyPlain.match(/Personen\s*:\s*\*?(\d+)\*?/i);
        var phoneMatch = bodyPlain.match(/Telefon\s*:\s*\*?(\+?[\d\s\/|-]{8,})\*?/i) || bodyPlain.match(/Mobil\s*:\s*\*?(\+?[\d\s\/|-]{8,})\*?/i);
        var titleMatch = bodyPlain.match(/TOP-Eintrag\s*\*?([^*|<|\r|\n]+)\*?/i) || bodyPlain.match(/Eintrag\s*\*?([^*|<|\r|\n]+)\*?/i) || subject.match(/Unterkunft\s+([^\r\n]+)/i);
        var addrMatch = bodyPlain.match(/\(([^)]+\d{5}[^)]+)\)/i);
        var msgMatch = bodyPlain.match(/Nachricht\s*:\s*([\s\S]*?)(?=\*?Sofort beantworten\*?|Nachrichtenverlauf|Mit freundlichen|Ihr Team|www\.Deutschland|$)/i);
        var cleanMsg = msgMatch ? msgMatch[1].replace(/<[^>]+>/g, "").trim() : "Keine Nachricht";

        var threadLink = "https://mail.google.com/mail/u/0/#all/" + thread.getId();

        var requestData = {
          portal: "deutschland-monteurzimmer.de",
          kundenNummer: kdNrMatch ? kdNrMatch[1].trim() : "N/A",
          gastName: nameMatch ? nameMatch[1].replace(/\*/g, "").trim() : "N/A",
          telefonnummer: phoneMatch ? phoneMatch[1].trim() : "N/A",
          zeitraum: periodStr,
          personen: personsMatch ? personsMatch[1].trim() : "N/A",
          unterkunftTitel: titleMatch ? titleMatch[1].replace(/\*/g, "").trim() : "N/A",
          link: threadLink,
          nachricht: cleanMsg,
          unterkunftAdresse: addrMatch ? addrMatch[1].trim() : "N/A",
          email: cleanEmail,
          erfasstAm: new Date().toISOString()
        };

        var jsonString = JSON.stringify(requestData, null, 2);

        // Im Tab "DMZ" eintragen (Schema A–K)
        var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
        var ss = SpreadsheetApp.openById(sheetId);
        var dmzSheet = ss.getSheetByName("DMZ") || ss.insertSheet("DMZ");

        if (dmzSheet.getLastRow() === 0) {
          dmzSheet.appendRow(["Zeit", "ID", "Name", "Telefonnummer", "Zeitraum", "Personen", "Titel", "Link", "Nachricht", "Adresse der Unterkunft", "JSON"]);
          dmzSheet.getRange(1, 1, 1, 11).setFontWeight("bold");
          dmzSheet.setRowHeight(1, 21); // Setzt auch die Zeilenhöhe des Headers auf 21
        }

        // Neue Zeile anhängen
        dmzSheet.appendRow([
          requestData.erfasstAm,
          requestData.kundenNummer,
          requestData.gastName,
          "'" + requestData.telefonnummer,
          requestData.zeitraum,
          requestData.personen,
          requestData.unterkunftTitel,
          requestData.link,
          requestData.nachricht,
          requestData.unterkunftAdresse,
          jsonString
        ]);

        // ÄNDERUNG: Genau für die neu angehängte Zeile die Höhe auf 21 Pixel erzwingen
        var lastRow = dmzSheet.getLastRow();
        dmzSheet.setRowHeight(lastRow, 21);

        thread.addLabel(label);
        logExecution(scriptName, "ERFOLG", "Antwort & Daten erfasst für: " + cleanEmail, "");
        processedAny = true;

      } catch (innerError) {
        logExecution(scriptName, "FEHLER", "Fehler bei E-Mail von: " + firstMessage.getFrom(), innerError.stack || innerError.toString());
      }
    }
  } catch (globalError) {
    logExecution(scriptName, "SYSTEM_FEHLER", "Globaler Skriptfehler", globalError.stack || globalError.toString());
  }
}