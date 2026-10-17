function filterAndProcessSmallGroups() {
  const SHEET1_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  let ss, dmzSheet;

  try {
    ss = SpreadsheetApp.openById(SHEET1_ID);
    dmzSheet = ss.getSheetByName('DMZ');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabelle: " + e.toString());
    return;
  }

  if (!dmzSheet) return;

  const data = dmzSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  Logger.log("--- VORAB-FILTERUNG START ---");

  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    const row = data[i];

    // 1. Prüfe Spalte AJ (Index 35) und Spalte S (Index 18)
    const emailStatus = String(row[35] || "").trim(); // Spalte AJ
    const dopplungsStatus = String(row[18] || "").trim(); // Spalte S (Dopplungen)

    // Wenn in Spalte AJ ODER in Spalte S irgendetwas steht -> Zeile überspringen
    if (emailStatus !== "" || dopplungsStatus !== "") continue;

    const rawFormPersonen = String(row[5] || "").trim(); // Spalte F
    const rawGeminiPersonen = String(row[13] || "").trim(); // Spalte N
    const currentPfad = String(row[19] || "").trim(); // Spalte T

    const formPersonen = parseInt(rawFormPersonen, 10);
    const geminiPersonen = parseInt(rawGeminiPersonen, 10);

    // 2. Nur verarbeiten, wenn genau 1 Person angegeben wurde
    if (formPersonen === 1) {
      const hasHigherGeminiValue = !isNaN(geminiPersonen) && geminiPersonen > formPersonen;

      // Wenn Gemini keinen höheren Wert geliefert hat
      if (!hasHigherGeminiValue) {
        // Marker '1P' in Spalte T (Spalte 20) setzen, falls noch nicht vorhanden
        if (currentPfad !== "1P") {
          dmzSheet.getRange(rowNum, 20).setValue("1P");
          Logger.log(`Zeile ${rowNum}: Pfad '1P' in Spalte T eingetragen.`);
        }

        const gastName = row[2] || "Damen und Herren"; // Spalte C
        const id = row[1] || ""; // Spalte B

        // E-Mail Adresse ermitteln (Prio 1: Spalte L, Prio 2: JSON in Spalte K)
        let recipientEmail = String(row[11] || "").trim(); // Spalte L (Index 11)
        if (!recipientEmail) {
          // Fallback auf Spalte K (Index 10) mit dem JSON
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

        if (!recipientEmail) {
          Logger.log(`Zeile ${rowNum}: Keine E-Mail-Adresse gefunden. Überspringe Entwurf-Erstellung.`);
          continue; // Ohne Empfänger können wir keinen Entwurf anlegen
        }

        let mailSubject = "Buchungsanfrage - Monteurwohnung L8 Street";
        if (id && id !== "N/A") {
          mailSubject += " (Anfrage ID: " + id + ")";
        }

        const footerText =
          "\n\nL8 Street \n\n" +
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

        const mailBody = "Hallo " + gastName + ",\n\n" +
          "vielen Dank für Ihre Buchungsanfrage.\n\n" +
          "Tut mir schrecklich leid, wir können aktuell leider garnichts anbieten.\n\n" +
          "Vielen Dank für Ihr Verständnis." + footerText;

        try {
          // Erstellt NUR einen Entwurf anstatt ihn sofort zu senden
          GmailApp.createDraft(recipientEmail, mailSubject, mailBody);
          // Markiere in Spalte AJ als ENTWURF
          dmzSheet.getRange(rowNum, 36).setValue("ENTWURF");
          SpreadsheetApp.flush();
          if (typeof logExecution === "function") {
            logExecution("DMZ Vorab-Filter", "ERFOLG", "Entwurf (1P) erstellt für: " + recipientEmail, "");
          }
        } catch (err) {
          if (typeof logExecution === "function") {
            logExecution("DMZ Vorab-Filter", "FEHLER", "Fehler bei Entwurf-Erstellung (1P) für: " + gastName, err.toString());
          }
        }
      }
    }
  }
}