function extractDataFromMonteurunterkunft() {
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
  var query = 'from:monteurunterkunft.de subject:"Anfrage vom" -label:' + labelPath + ' (is:unread OR newer_than:7d)';
  var threads = GmailApp.search(query);

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var message = thread.getMessages()[0];
    try {
      var bodyHtml = message.getBody();
      
      // HTML sauber in Text umwandeln
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

      var getVal = function(regex) {
        var m = text.match(regex);
        if (!m) return "";
        var val = m[1].trim();
        return (val === "-" || val === "" || val === "N/A") ? "" : val;
      };

      var firma = getVal(/Firma\s*:\s*\n*([^\n]+)/i);
      var anrede = getVal(/Anrede\s*:\s*\n*([^\n]+)/i);
      var vn = getVal(/Vorname\s*:\s*\n*([^\n]+)/i);
      var nn = getVal(/(?:^|\n)\s*Name\s*:\s*\n*([^\n]+)/i);

      var gastName = ((anrede ? anrede + " " : "") + vn + " " + nn).trim();
      if (firma) gastName += " (" + firma + ")";

      var tel = getVal(/Telefon\s*:\s*\n*([^\n]+)/i) || "Nicht angegeben";
      
      var emailMatch = text.match(/E-Mail\s*:\s*\n*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);
      var email = emailMatch ? emailMatch[1].trim() : "Nicht angegeben";
      if (email.indexOf("monteurunterkunft.de") !== -1) email = "Nicht angegeben";

      var anreise = getVal(/Anreisedatum\s*:\s*\n*([^\n]+)/i);
      var abreise = getVal(/Abreisedatum\s*:\s*\n*([^\n]+)/i);
      var zeitraumStr = (anreise && abreise) ? (anreise + " - " + abreise) : "N/A";
      var personen = getVal(/Anzahl Personen\s*:\s*\n*([^\n]+)/i) || "N/A";

      // Nachricht zwischen Header und Support-Bereich isolieren
      var msgMatch = text.match(/Nachricht des Zimmersuchenden:\s*\n+([\s\S]*?)(?=\n+\s*Support|\n+\s*Solltest Du)/i);
      var nachricht = msgMatch ? msgMatch[1].trim() : "Keine Nachricht";

      var uTitel = "N/A";
      var uMatch = text.match(/Unterkunft\s*["“]([^"”]+)["”]/i);
      if (uMatch) uTitel = uMatch[1].trim();

      var requestData = {
        portal: "monteurunterkunft.de",
        kundenNummer: "N/A",
        gastName: gastName || "N/A",
        telefonnummer: tel,
        zeitraum: zeitraumStr,
        personen: personen,
        unterkunftTitel: uTitel,
        link: "https://mail.google.com/mail/u/0/#all/" + thread.getId(),
        nachricht: nachricht,
        unterkunftAdresse: "N/A",
        email: email,
        erfasstAm: new Date().toISOString()
      };

      dmzSheet.appendRow([JSON.stringify(requestData)]);
      dmzSheet.setRowHeight(dmzSheet.getLastRow(), 21);
      thread.addLabel(label);
    } catch (e) {
      Logger.log("Fehler bei Monteurunterkunft-ID " + thread.getId() + ": " + e.toString());
    }
  }
}