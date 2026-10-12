function sendOfferEmail() {
  const sheetId = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const ss = SpreadsheetApp.openById(sheetId);
  const dmzSheet = ss.getSheetByName('DMZ');
  if (!dmzSheet) return;

  const data = dmzSheet.getDataRange().getValues();
  if (data.length <= 1) return; // Nur Header vorhanden

  // Test-E-Mail Empfänger
  const testRecipient = "info@l8street.com";

  // Standard Fußzeile & Abschluss definieren
  const footerText = 
    "\n\nAGB (terms): L8Street.com/agb\n\n" +
    "Eine Übersicht aller unserer Wohnungen inklusive dem Datum ab wann diese frei werden, Fotos, Adressen und Preisen finden Sie unter https://a.l8street.com/\n\n" +
    "Für eine besonders schnelle Bearbeitung empfehlen wir Ihnen, die Buchung direkt auf unserer Website unter a.L8Street.com abzuschließen oder uns direkt per WhatsApp Nachricht zu kontaktieren: +4917684801295 \n\n" +
    "Freundliche Grüße\n\n" +
    "L8 Street \n\n" +
    "© 2017-2026\n" +
    "L8 Street GmbH\n" +
    "Postfach CAYA595616\n" +
    "96035 Bamberg\n\n" +
    "WhatsApp: +4917684801295\n" +
    "E-Mail: support@L8Street.com\n" +
    "Sofort-Buchung: https://a.l8street.com/\n" +
    "Web: https://L8Street.com\n" +
    "Geschäftsführer: Raul Leneweit\n" +
    "Sitz der Gesellschaft: Forststraße 65,\n" +
    "75223 Niefern-Öschelbronn\n" +
    "Amtsgericht Mannheim HRB: 72 85 52\n" +
    "USt-IdNr.: DE 314 603 427";

  for (let i = 1; i < data.length; i++) {
    let row = data[i];

    // Vermeide Mehrfachversand (Spalte P / Index 15 als Status-Marker)
    let status = row[15] || "";
    if (status === "GESENDET") continue;

    let gastName = row[2] || "Damen und Herren";
    let linksStr = row[14] || ""; // Spalte O (Index 14)
    let id = row[1] || "";

    let mailSubject = "Angebotsübersicht - Monteurwohnung L8 Street";
    if (id && id !== "N/A") {
      mailSubject += " (Anfrage ID: " + id + ")";
    }

    let mailBody = "";

    // FALL 1: Mindestens ein Link in Spalte O vorhanden
    if (linksStr && linksStr !== "Keine verfügbar" && !linksStr.startsWith("Fehler") && !linksStr.startsWith("Unter Mindestaufenthalt")) {
      let linkList = linksStr.split(',').map(l => l.trim());
      let formattedLinks = linkList.map((link, index) => "Option " + (index + 1) + ": " + link).join("\n");

      mailBody = "Hallo " + gastName + ",\n\n" +
        "vielen Dank für Ihre Anfrage!\n\n" +
        "Wir haben passende Unterkünfte für Ihren Zeitraum verfügbar:\n" +
        formattedLinks + "\n\n" +
        "Eckdaten & Inklusivleistungen:\n" +
        "• Waschmaschine, WLAN und Bettwäsche sind kostenfrei inklusive.\n" +
        "• Parkmöglichkeiten vorhanden.\n\n" +
        "Preisbeispiel & Konditionen:\n" +
        "• Ab €89 pro Nacht netto gesamt für alle Personen zusammen.\n" +
        "• Einmalige Endreinigung 89€ und 7% MwSt. kommen dazu, sonst keine weiteren Kosten.\n\n" +
        "Gefällt Ihnen eines der Angebote? Antworten Sie mir einfach kurz auf diese E-Mail, dann reservieren wir die Wohnung direkt für Sie.";
    } 
    // FALL 2: Keine passende Wohnung frei
    else {
      mailBody = "Hallo " + gastName + ",\n\n" +
        "vielen Dank für Ihre Buchungsanfrage.\n\n" +
        "Leider haben wir im von Ihnen gewünschten Zeitraum direkt vor Ort keine exakt passende Kapazität frei.\n\n" +
        "Auf unserer Website finden Sie eine Übersicht weiterer Wohnungen die in den kommenden Tagen frei werden: a.L8Street.com\n\n" +
        "Schauen Sie dort gerne durch oder geben Sie uns kurz Bescheid, falls auch ein leicht verschobener Zeitraum für Sie infrage kommt.";
    }

    // Fußzeile an den jeweiligen E-Mail-Text anhängen
    mailBody += footerText;

    try {
      GmailApp.sendEmail(testRecipient, mailSubject, mailBody);

      // Markiert die Zeile in Spalte P als GESENDET
      dmzSheet.getRange(i + 1, 16).setValue("GESENDET");

      if (typeof logExecution === "function") {
        logExecution("DMZ Angebote", "ERFOLG", "Angebot (TEST) gesendet an info@l8street.com für: " + gastName, "");
      }
    } catch (err) {
      if (typeof logExecution === "function") {
        logExecution("DMZ Angebote", "FEHLER", "Fehler beim Angebot-Versand für: " + gastName, err.toString());
      }
    }
  }
}