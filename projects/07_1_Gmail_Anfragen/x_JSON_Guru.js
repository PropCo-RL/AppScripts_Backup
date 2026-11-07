function extractDataFromMonteurzimmerGuru() {
  var labelPath = "JSON";
  var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
  var ss = SpreadsheetApp.openById(sheetId);
  var dmzSheet = ss.getSheetByName("DMZ2") || ss.insertSheet("DMZ2");
  
  if (dmzSheet.getLastRow() === 0) {
    dmzSheet.appendRow(["JSON"]);
    dmzSheet.getRange(1, 1).setFontWeight("bold");
    dmzSheet.setRowHeight(1, 21);
    dmzSheet.getRange("A:A").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  
  var label = GmailApp.getUserLabelByName(labelPath) || GmailApp.createLabel(labelPath);
  var query = 'from:monteurzimmerguru.de subject:MonteurzimmerGuru -label:' + labelPath + ' (is:unread OR newer_than:7d)';
  var threads = GmailApp.search(query);

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var message = thread.getMessages()[0];
    try {
      var bodyHtml = message.getBody();
      var text = bodyHtml
        .replace(/<style([\s\S]*?)<\/style>/gi, "")
        .replace(/<script([\s\S]*?)<\/script>/gi, "")
        .replace(/<br\s*[\/]?>/gi, "\n")
        .replace(/<\/p>/gi, "\n")
        .replace(/<\/tr>/gi, "\n")
        .replace(/<\/td>/gi, " ")
        .replace(/<\/div>/gi, "\n")
        .replace(/<[^>]+>/gi, "")
        .replace(/&nbsp;/gi, " ")
        .replace(/\u00A0/g, " ")
        .replace(/&amp;/gi, "&")
        .replace(/&#39;/gi, "'")
        .replace(/&quot;/gi, '"')
        .replace(/\r/g, "");

      var getField = function(header) {
        var m = text.match(new RegExp(header + "\\s*\\n+([^\\n]+)", "i"));
        return m ? m[1].trim() : "N/A";
      };

      var msgMatch = text.match(/Nachricht vom Gast\s*\n+([\s\S]*?)(?=\n*Anfrage beantworten|\n*Werten Sie)/i);
      var nachricht = msgMatch ? msgMatch[1].trim() : "Keine Nachricht";

      var requestData = {
        portal: "monteurzimmerguru.de",
        kundenNummer: "N/A",
        gastName: "N/A",
        telefonnummer: "Nicht angegeben",
        zeitraum: getField("Zeitraum"),
        personen: getField("Personen"),
        unterkunftTitel: getField("Unterkunft"),
        link: "https://mail.google.com/mail/u/0/#all/" + thread.getId(),
        nachricht: nachricht,
        unterkunftAdresse: getField("Adresse"),
        email: "Nicht angegeben",
        erfasstAm: new Date().toISOString()
      };

      dmzSheet.appendRow([JSON.stringify(requestData)]);
      dmzSheet.setRowHeight(dmzSheet.getLastRow(), 21);
      thread.addLabel(label);
    } catch (e) {
      Logger.log("Fehler bei Guru-ID " + thread.getId() + ": " + e.toString());
    }
  }
}