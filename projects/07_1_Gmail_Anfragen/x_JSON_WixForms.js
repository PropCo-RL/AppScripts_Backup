function extractDataFromWix() {
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
  var query = 'subject:AnfrageHomepageLanding -label:' + labelPath + ' (is:unread OR newer_than:7d)';
  var threads = GmailApp.search(query);

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var message = thread.getMessages()[0];
    try {
      var bodyPlain = message.getPlainBody() || "";
      var getField = function(regex) {
        var m = bodyPlain.match(regex);
        return m ? m[1].trim() : "N/A";
      };

      var city = getField(/Select City\s*:\s*([^\n\r]+)/i);
      var persons = getField(/Number of Persons\s*:\s*([^\n\r]+)/i);
      var email = getField(/E-Mail-Address\s*:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);

      if (email === "N/A" || email.indexOf("@") === -1) {
        var replyTo = message.getReplyTo();
        if (replyTo) {
          var rm = replyTo.match(/<([^>]+)>/) || [null, replyTo];
          email = rm[1] ? rm[1].trim() : "Nicht angegeben";
        }
      }

      var requestData = {
        portal: "wix",
        kundenNummer: "N/A",
        gastName: "N/A",
        telefonnummer: "Nicht angegeben",
        zeitraum: "N/A",
        personen: persons,
        unterkunftTitel: "N/A",
        link: "https://mail.google.com/mail/u/0/#all/" + thread.getId(),
        nachricht: "Formular: Anfrage Homepage Landing",
        unterkunftAdresse: city,
        email: email,
        erfasstAm: new Date().toISOString()
      };

      dmzSheet.appendRow([JSON.stringify(requestData)]);
      dmzSheet.setRowHeight(dmzSheet.getLastRow(), 21);
      thread.addLabel(label);
    } catch (e) {
      Logger.log("Fehler bei Wix-ID " + thread.getId() + ": " + e.toString());
    }
  }
}