function sendAutoResponseDMZ2() {
  var cutoffDate = new Date("2026-10-03T12:57:00+02:00");
  var labelPath = "API_Buchung/DMZ/aDMZ2";

  try {
    var searchQuery = 'from:mail.dmz.de (subject:"Buchungsanfrage für" OR subject:"Neue Anfrage für") -label:' + labelPath + ' (is:unread OR newer_than:7d)';
    var threads = GmailApp.search(searchQuery);
    if (threads.length === 0) return;

    var label = GmailApp.getUserLabelByName(labelPath) || GmailApp.createLabel(labelPath);

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

        // Daten extrahieren für das JSON-Objekt
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

        // Kompakter Einzeiler ohne Zeilenumbrüche/Einrückung:
        var jsonString = JSON.stringify(requestData);

        // Im Tab "DMZ2" eintragen (nur Spalte A)
        var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
        var ss = SpreadsheetApp.openById(sheetId);
        var dmzSheet = ss.getSheetByName("DMZ2") || ss.insertSheet("DMZ2");

        if (dmzSheet.getLastRow() === 0) {
          dmzSheet.appendRow(["JSON"]);
          dmzSheet.getRange(1, 1).setFontWeight("bold");
          dmzSheet.setRowHeight(1, 21);
          // Setzt die gesamte Spalte A auf "Abschneiden", damit Zellenhöhe nie wächst
          dmzSheet.getRange("A:A").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        }

        // Neue Zeile anhängen
        dmzSheet.appendRow([jsonString]);

        var lastRow = dmzSheet.getLastRow();
        dmzSheet.setRowHeight(lastRow, 21);

        // Neues Label vergeben
        thread.addLabel(label);

      } catch (innerError) {
        // Fehler bei einzelner E-Mail wird ignoriert/übersprungen
      }
    }
  } catch (globalError) {
    // Globaler Skriptfehler wird abgefangen
  }
}