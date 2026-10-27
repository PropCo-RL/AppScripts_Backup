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

    // Bestehendes Sheet "MMZ"
    var mmzSheet = ss.getSheetByName("MMZ") || ss.insertSheet("MMZ");
    if (mmzSheet.getLastRow() === 0) {
      mmzSheet.appendRow([
        "Erfasst am", "Kundennummer", "Gast Name", "Telefonnummer",
        "E-Mail", "Zeitraum", "Personen", "Unterkunft",
        "Gmail Link", "Nachricht", "Adresse", "JSON String"
      ]);
      mmzSheet.getRange(1, 1, 1, 12).setFontWeight("bold");
    }

    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var message = thread.getMessages()[0];
      var bodyHtml = message.getBody();

      // 1. E-Mail-Adresse extrahieren
      var emailMatch = bodyHtml.match(/mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i) ||
                       bodyHtml.match(/E-Mail[\s\S]*?>([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})</i);
      var customerEmail = emailMatch ? emailMatch[1].trim() : "Nicht gefunden";

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
        if (phoneMatch && phoneMatch[1] && phoneMatch[1] !== "null") phone = phoneMatch[1].replace(/\\/g, "");

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

        if (customerEmail === "Nicht gefunden") {
          var emailUrlMatch = htmlContent.match(/"email"\s*:\s*"([^"]+)"/);
          if (emailUrlMatch && emailUrlMatch[1] && emailUrlMatch[1] !== "null") {
            customerEmail = emailUrlMatch[1].replace(/\\/g, "");
          }
        }
      } catch (fetchError) {
        logExecution(scriptName, "FEHLER", "URL konnte nicht geladen werden: " + targetUrl, fetchError.toString());
      }

      function decodeUnicode(str) {
        if (!str) return "";
        return str.replace(/\\u([0-9a-fA-F]{4})/g, function(match, grp) {
          return String.fromCharCode(parseInt(grp, 16));
        });
      }

      var fullGastName = decodeUnicode((firstName + " " + lastName).trim());
      if (company) fullGastName += " (" + decodeUnicode(company) + ")";

      var fullAddress = decodeUnicode((street + " " + number + ", " + zip + " " + locality).trim());
      if (fullAddress === ",") fullAddress = "N/A";

      var cleanedHousingName = decodeUnicode(housingName);
      var cleanedMessage = decodeUnicode(guestMessage);
      var zeitraumStr = (bookingFrom && bookingTo) ? (bookingFrom + " - " + bookingTo) : "N/A";

      if (!fullGastName && phone === "Nicht gefunden" && customerEmail === "Nicht gefunden") {
        thread.addLabel(label);
        continue;
      }

      var gmailMessageLink = "https://mail.google.com/mail/u/0/#all/" + message.getId();

      // Einheitliches JSON-Schema
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

      // In Tab "MMZ" wie gewohnt mit allen Spalten schreiben
      mmzSheet.appendRow([
        requestData.erfasstAm,
        requestData.kundenNummer,
        requestData.gastName,
        requestData.telefonnummer,
        requestData.email,
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