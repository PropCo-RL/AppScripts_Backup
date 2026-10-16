function sendOfferEmail() {
  const sheetId = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const ss = SpreadsheetApp.openById(sheetId);
  const dmzSheet = ss.getSheetByName('DMZ');
  if (!dmzSheet) return;

  const data = dmzSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const footerText =
    "\n\nAGB (terms): L8Street.com/agb\n\n" +
    "Für eine besonders schnelle Bearbeitung empfehlen wir Ihnen uns direkt per WhatsApp Nachricht zu kontaktieren: +4917684801295.  \n\n" +
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
    let rowNum = i + 1;
    let row = data[i];

    // Prüft Spalte AA (Index 26): Wenn WIRKLICH IRGENDWAS drin steht -> Zeile überspringen
    let emailStatus = String(row[26] || "").trim();
    if (emailStatus !== "") continue; 

    // E-Mail-Adresse ermitteln (Prio 1: Spalte L [Index 11], Prio 2: JSON in Spalte K [Index 10])
    let recipientEmail = String(row[11] || "").trim(); 
    if (!recipientEmail) {
      const jsonString = String(row[10] || "").trim();
      if (jsonString) {
        try {
          const jsonData = JSON.parse(jsonString);
          if (jsonData && jsonData.email) {
            recipientEmail = jsonData.email.trim();
          }
        } catch (e) {
          Logger.log(`Zeile ${rowNum}: Fehler beim Parsen des JSON in Spalte K.`);
        }
      }
    }

    // Ohne Empfänger-E-Mail kann kein Entwurf erstellt werden -> überspringen
    if (!recipientEmail) {
      Logger.log(`Zeile ${rowNum}: Keine E-Mail-Adresse gefunden. Überspringe Entwurf-Erstellung.`);
      continue; 
    }

    let gastName = row[2] || "Damen und Herren"; // Spalte C
    
    // Liest Wohnungs-Links aus Spalte Z (Index 25)
    let linksStr = row[25] || ""; 
    
    let id = row[1] || ""; // Spalte B

    let mailSubject = "Angebotsübersicht - Monteurwohnung L8 Street";
    if (id && id !== "N/A") {
      mailSubject += " (Anfrage ID: " + id + ")";
    }

    let mailBody = "";

    if (linksStr && linksStr !== "Keine verfügbar" && !linksStr.startsWith("Fehler") && !linksStr.startsWith("Unter Mindestaufenthalt")) {
      let linkList = linksStr.split(',').map(l => l.trim());
      let formattedLinks = linkList.map((link, index) => "Option " + (index + 1) + ": " + link).join("\n");

      mailBody = "Hallo " + gastName + ",\n\n" +
        "vielen Dank für Ihre Anfrage!\n\n" +
        "Wir haben passende Unterkünfte für Ihren Zeitraum verfügbar:\n" +
        formattedLinks + "\n\n" +
        "Kostenfrei inklusive:\n" +
        "• Waschmaschine, WLAN und Bettwäsche\n" +
        "• Parkmöglichkeiten vorhanden.\n\n";
        
    } else {
      mailBody = "Hallo " + gastName + ",\n\n" +
        "vielen Dank für Ihre Buchungsanfrage.\n\n" +
        "Leider haben wir im von Ihnen gewünschten Zeitraum direkt vor Ort keine exakt passende Kapazität frei.\n\n" +
        "Auf unserer Website finden Sie eine Übersicht weiterer Wohnungen die in den kommenden Tagen frei werden: a.L8Street.com\n\n" +
        "Schauen Sie dort gerne durch oder geben Sie uns kurz Bescheid, falls auch ein leicht verschobener Zeitraum für Sie infrage kommt.";
    }

    mailBody += footerText;

    try {
      // Erstellt NUR einen Entwurf anstelle der direkten E-Mail-Versendung
      GmailApp.createDraft(recipientEmail, mailSubject, mailBody);

      // Trägt "ENTWURF" in Spalte AA ein (Spalte 27 in 1-basierter Zählung)
      dmzSheet.getRange(rowNum, 27).setValue("ENTWURF");
      SpreadsheetApp.flush(); // ZWINGEND NÖTIG GEGEN MEHRFACHVERARBEITUNG

      if (typeof logExecution === "function") {
        logExecution("DMZ Angebote", "ERFOLG", "Entwurf Angebot erstellt für: " + recipientEmail + " (" + gastName + ")", "");
      }
    } catch (err) {
      if (typeof logExecution === "function") {
        logExecution("DMZ Angebote", "FEHLER", "Fehler bei Entwurf-Erstellung für: " + gastName, err.toString());
      }
    }
  }
}