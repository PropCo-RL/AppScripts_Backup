function extractDataFromKleinanzeigen() {
  var labelPath = "JSON";
  var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
  var ss = SpreadsheetApp.openById(sheetId);
  var dmzSheet = ss.getSheetByName("DMZ2") || ss.insertSheet("DMZ2");
  
  // Tab-Initialisierung mit Zeilenhöhe 21 und CLIP-Formatierung
  if (dmzSheet.getLastRow() === 0) {
    dmzSheet.appendRow(["JSON"]);
    dmzSheet.getRange(1, 1).setFontWeight("bold");
    dmzSheet.setRowHeight(1, 21);
    dmzSheet.getRange("A:A").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  
  var label = GmailApp.getUserLabelByName(labelPath) || GmailApp.createLabel(labelPath);
  var query = 'from:mail.kleinanzeigen.de subject:"Kleinanzeigen Anfrage zum Objekt" -label:' + labelPath + ' (is:unread OR newer_than:7d)';
  var threads = GmailApp.search(query);

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var message = thread.getMessages()[0];
    try {
      var attachments = message.getAttachments();
      var xmlContent = "";
      for (var j = 0; j < attachments.length; j++) {
        if (attachments[j].getName().indexOf(".xml") !== -1) {
          xmlContent = attachments[j].getDataAsString();
          break;
        }
      }

      var kundenNr = "N/A";
      var gastName = "N/A";
      var tel = "Nicht angegeben";
      var email = "Nicht angegeben";
      var nachricht = "Keine Nachricht";
      var unterkunftAdresse = "N/A";

      // 1. Primär den OpenImmo-XML-Anhang auswerten
      if (xmlContent) {
        var getSubTag = function(parentXml, tag) {
          var m = parentXml.match(new RegExp("<" + tag + ">([\\s\\S]*?)<\\/" + tag + ">", "i"));
          return m ? m[1].trim() : "";
        };

        var objektBlock = (xmlContent.match(/<objekt>([\s\S]*?)<\/objekt>/i) || ["", ""])[1];
        var interessentBlock = (xmlContent.match(/<interessent>([\s\S]*?)<\/interessent>/i) || ["", ""])[1];

        // Objektdaten (Inserats-ID & Adresse der Monteurwohnung)
        kundenNr = getSubTag(objektBlock, "portal_obj_id") || "N/A";
        var objStrasse = getSubTag(objektBlock, "strasse");
        var objOrt = getSubTag(objektBlock, "ort");
        unterkunftAdresse = ((objStrasse ? objStrasse + ", " : "") + objOrt).trim() || "N/A";

        // Gastdaten
        var anrede = getSubTag(interessentBlock, "anrede");
        var vn = getSubTag(interessentBlock, "vorname");
        var nn = getSubTag(interessentBlock, "nachname");
        gastName = ((anrede ? anrede + " " : "") + vn + " " + nn).trim() || "N/A";
        tel = getSubTag(interessentBlock, "tel") || "Nicht angegeben";
        email = getSubTag(interessentBlock, "email") || "Nicht angegeben";
        nachricht = getSubTag(interessentBlock, "anfrage") || "Keine Nachricht";

      } else {
        // 2. Fallback: Reiner Mail-Text, falls kein XML vorhanden ist
        var body = message.getPlainBody() || "";
        var kMatch = body.match(/Anzeige:\s*(\d+)/i);
        if (kMatch) kundenNr = kMatch[1].trim();

        var nMatch = body.match(/Name\s+([^\n\r]+)/i);
        if (nMatch) gastName = nMatch[1].trim();

        var eMatch = body.match(/E-Mail\s+([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);
        if (eMatch) email = eMatch[1].trim();

        var tMatch = body.match(/Telefon\s+(\+?[\d\s\/|-]{7,})/i);
        if (tMatch) tel = tMatch[1].trim();

        var mMatch = body.match(/Nachricht\s+([\s\S]*?)(?=Wenn Sie Fragen haben|kleinanzeigen\.de)/i);
        if (mMatch) nachricht = mMatch[1].trim();

        var adrMatch = body.match(/(\d{5}\s+[^\n\r]+)/i);
        if (adrMatch) unterkunftAdresse = adrMatch[1].trim();
      }

      // Titel aus Betreff (erkennt alle typografischen Anführungszeichen: „ “ ” " ' ‟)
      var uTitel = "N/A";
      var subMatch = message.getSubject().match(/zum Objekt\s*[\u201E\u201C\u201D"„“‟']([^"“'”\u201E\u201C\u201D\u201F]+)[\u201E\u201C\u201D\u201F"„“‟']/i);
      if (subMatch) uTitel = subMatch[1].trim();

      var requestData = {
        portal: "Kleinanzeigen",
        kundenNummer: kundenNr,
        gastName: gastName,
        telefonnummer: tel,
        zeitraum: "N/A",
        personen: "N/A",
        unterkunftTitel: uTitel,
        link: "https://mail.google.com/mail/u/0/#all/" + thread.getId(),
        nachricht: nachricht,
        unterkunftAdresse: unterkunftAdresse,
        email: email,
        erfasstAm: new Date().toISOString()
      };

      dmzSheet.appendRow([JSON.stringify(requestData)]);
      dmzSheet.setRowHeight(dmzSheet.getLastRow(), 21);
      thread.addLabel(label);

    } catch (e) {
      Logger.log("Fehler bei Kleinanzeigen-ID " + thread.getId() + ": " + e.toString());
    }
  }
}