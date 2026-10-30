function extractDataFromMeinMonteurzimmer() {
  var scriptName = "Mein-Monteurzimmer Data Extractor";
  var labelPath = "API_Buchung/MMZ/aMMZ";

  try {
    var searchQuery = 'from:mein-monteurzimmer.de subject:"Neue Buchungsanfrage auf mein-Monteurzimmer.de" -label:' + labelPath + ' (is:unread OR newer_than:7d)';
    var threads = GmailApp.search(searchQuery);
    if (threads.length === 0) return;

    var label = GmailApp.getUserLabelByName(labelPath) || GmailApp.createLabel(labelPath);
    var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
    var ss = SpreadsheetApp.openById(sheetId);
    
    // Tab "MMZ"
    var mmzSheet = ss.getSheetByName("MMZ") || ss.insertSheet("MMZ");
    if (mmzSheet.getLastRow() === 0) {
      // 11 Spalten: "E-Mail" wurde komplett aus den sichtbaren Spalten entfernt
      mmzSheet.appendRow([
        "Erfasst am", "Kundennummer", "Gast Name", "Telefonnummer",
        "Zeitraum", "Personen", "Unterkunft",
        "Gmail Link", "Nachricht", "Adresse", "JSON String"
      ]);
      mmzSheet.getRange(1, 1, 1, 11).setFontWeight("bold");
    }

    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var message = thread.getMessages()[0];
      var bodyHtml = message.getBody();

      // 1. E-Mail-Adresse gezielt aus der Kontaktdaten-Tabelle der Mail parsen
      var customerEmail = "Nicht gefunden";
      var emailMatch = bodyHtml.match(/>E-Mail<\/td>\s*<td[^>]*>([^<]+)<\/td>/i);
      if (emailMatch && emailMatch[1]) {
        customerEmail = emailMatch[1].replace(/=[\r\n]+/g, "").trim();
      }

      // 2. Link zur Detailseite extrahieren
      var linkMatch = bodyHtml.match(/https:\/\/mein-monteurzimmer\.de\/anfrage\/[a-zA-Z0-9\/]+/);
      if (!linkMatch) {
        thread.addLabel(label);
        continue;
      }
      var targetUrl = linkMatch[0];

      var phone = "Nicht gefunden";
      var firstName = "", lastName = "", company = "";
      var housingName = "", housingId = "";
      var street = "", number = "", zip = "", locality = "";
      var bookingFrom = "", bookingTo = "", maxPersons = "";
      var guestMessage = "Keine Nachricht hinterlassen", createdDate = "";

      try {
        var response = UrlFetchApp.fetch(targetUrl);
        var htmlContent = response.getContentText();

        var phoneMatch = htmlContent.match(/"phone"\s*:\s*"([^"]+)"/);
        if (phoneMatch && phoneMatch[1] && phoneMatch[1] !== "null") {
          phone = phoneMatch[1].replace(/\\/g, "").trim();
        }

        var fnMatch = htmlContent.match(/"firstname"\s*:\s*"([^"]+)"/);
        if (fnMatch) firstName = fnMatch[1];

        var lnMatch = htmlContent.match(/"lastname"\s*:\s*"([^"]+)"/);
        if (lnMatch) lastName = lnMatch[1];

        var compMatch = htmlContent.match(/"company"\s*:\s*"([^"]+)"/);
        if (compMatch && compMatch[1] !== "null") company = compMatch[1];

        var hNameMatch = htmlContent.match(/"housing"\s*:\s*\{[^}]*"name"\s*:\s*"([^"]+)"/);
        if (hNameMatch) housingName = hNameMatch[1];

        var hIdMatch = htmlContent.match(/"housing"\s*:\s*\{[^}]*"id"\s*:\s*"([^"]+)"/);
        if (hIdMatch) housingId = hIdMatch[1];

        var bfMatch = htmlContent.match(/"bookingFrom"\s*:\s*"([^"T]+)/);
        if (bfMatch) bookingFrom = bfMatch[1];

        var btMatch = htmlContent.match(/"bookingTo"\s*:\s*"([^"T]+)/);
        if (btMatch) bookingTo = btMatch[1];

        var mpMatch = htmlContent.match(/"maxPersons"\s*:\s*(\d+)/);
        if (mpMatch) maxPersons = mpMatch[1];

        var strMatch = htmlContent.match(/"street"\s*:\s*"([^"]+)"/);
        if (strMatch) street = strMatch[1];

        var numMatch = htmlContent.match(/"number"\s*:\s*"([^"]+)"/);
        if (numMatch) number = numMatch[1];

        var zipMatch = htmlContent.match(/"zip"\s*:\s*"([^"]+)"/);
        if (zipMatch) zip = zipMatch[1];

        var locMatch = htmlContent.match(/"locality"\s*:\s*"([^"]+)"/);
        if (locMatch) locality = locMatch[1];

        var msgMatch = htmlContent.match(/"message"\s*:\s*"((?:\\"|[^"])*?)"/);
        if (msgMatch && msgMatch[1]) {
          guestMessage = msgMatch[1].replace(/\\"/g, '"').replace(/\\\//g, '/').replace(/\\n/g, ' ');
        }

        var createdMatch = htmlContent.match(/"createdDate"\s*:\s*"([^"]+)"/);
        if (createdMatch) createdDate = createdMatch[1];

        // Fallback: E-Mail aus Detailseite holen, falls in Mail unauffindbar
        if (customerEmail === "Nicht gefunden") {
          var emailUrlMatch = htmlContent.match(/"email"\s*:\s*"([^"]+)"/);
          if (emailUrlMatch && emailUrlMatch[1] && emailUrlMatch[1] !== "null") {
            customerEmail = emailUrlMatch[1].replace(/\\/g, "").trim();
          }
        }
      } catch (fetchError) {
        logExecution(scriptName, "FEHLER", "URL konnte nicht geladen werden: " + targetUrl, fetchError.toString());
      }

      var fullGastName = decodeUnicodeString(firstName + " " + lastName).trim();
      if (company) fullGastName += " (" + decodeUnicodeString(company) + ")";

      var fullAddress = decodeUnicodeString((street + " " + number + ", " + zip + " " + locality).trim());
      if (fullAddress === ",") fullAddress = "N/A";

      var cleanedHousingName = decodeUnicodeString(housingName);
      var cleanedMessage = decodeUnicodeString(guestMessage);
      var zeitraumStr = (bookingFrom && bookingTo) ? (bookingFrom + " - " + bookingTo) : "N/A";

      if (!fullGastName && phone === "Nicht gefunden" && customerEmail === "Nicht gefunden") {
        thread.addLabel(label);
        continue;
      }

      // Schutz vor #ERROR! in Google Sheets: Nummern mit führendem + als reinen Text kennzeichnen
      var sheetPhone = phone;
      if (sheetPhone.indexOf("+") === 0) {
        sheetPhone = "'" + sheetPhone;
      }

      var gmailMessageLink = "https://mail.google.com/mail/u/0/#all/" + message.getId();

      // Einheitliches JSON-Schema (E-Mail bleibt hier wie gewünscht vollständig enthalten)
      var requestData = {
        portal: "mein-monteurzimmer.de",
        kundenNummer: housingId || "N/A",
        gastName: fullGastName || "N/A",
        telefonnummer: phone,
        zeitraum: zeitraumStr,
        personen: maxPersons || "N/A",
        unterkunftTitel: cleanedHousingName || "N/A",
        link: gmailMessageLink,
        nachricht: cleanedMessage,
        unterkunftAdresse: fullAddress,
        email: customerEmail,
        erfasstAm: new Date().toISOString(),
        anfrageUrl: targetUrl,
        anfrageErstelltAm: createdDate || "N/A"
      };

      var jsonString = JSON.stringify(requestData);

      // In Tab "MMZ" schreiben: Exakt 11 Spalten (Spalte E = Zeitraum)
      mmzSheet.appendRow([
        requestData.erfasstAm,
        requestData.kundenNummer,
        requestData.gastName,
        sheetPhone,
        requestData.zeitraum,
        requestData.personen,
        requestData.unterkunftTitel,
        requestData.link,
        requestData.nachricht,
        requestData.unterkunftAdresse,
        jsonString
      ]);

      thread.addLabel(label);
      logExecution(scriptName, "ERFOLG", "Daten erfasst für: " + requestData.gastName, "");
    }
  } catch (globalError) {
    logExecution(scriptName, "SYSTEM_FEHLER", "Globaler Skriptfehler", globalError.stack || globalError.toString());
  }
}

// Hilfsfunktion außerhalb platziert
function decodeUnicodeString(str) {
  if (!str) return "";
  return str.replace(/\\u([0-9a-fA-F]{4})/g, function(match, grp) {
    return String.fromCharCode(parseInt(grp, 16));
  });
}