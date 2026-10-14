// =================================================================
// VORAB-FILTERUNG (1P / 2P DETEKTION & AUTO-ABSAGE FÜR 1 PERSON)
// =================================================================
function filterAndProcessSmallGroups() {
  const SHEET1_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const testRecipient = "info@l8street.com";

  let ss, dmzSheet;
  try {
    ss = SpreadsheetApp.openById(SHEET1_ID);
    dmzSheet = ss.getSheetByName('DMZ');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabelle: " + e.toString());
    return;
  }

  if (!dmzSheet) {
    Logger.log("FEHLER: Tabellenblatt 'DMZ' nicht gefunden.");
    return;
  }

  const data = dmzSheet.getDataRange().getValues();
  if (data.length <= 1) return; // Nur Header vorhanden

  Logger.log("--- VORAB-FILTERUNG START ---");

  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    const row = data[i];

    // Spalten-Indizes (0-basiert für JavaScript Array `row`):
    // F (Index 5): Personen laut Formular
    // N (Index 13): Personen laut Gemini
    // T (Index 19): Pfad-Zuordnung (1P / 2P)
    // AC (Index 28): Status E-Mail Versand

    const rawFormPersonen = String(row[5] || "").trim();
    const rawGeminiPersonen = String(row[13] || "").trim();
    const currentPfad = String(row[19] || "").trim(); // Spalte T (Index 19)
    const emailStatus = String(row[28] || "").trim(); // Spalte AC (Index 28)

    // Ziffern aus den Feldern parsen
    const formPersonen = parseInt(rawFormPersonen, 10);
    const geminiPersonen = parseInt(rawGeminiPersonen, 10);

    // =============================================================
    // SCHRITT 1: KLASSIFIZIERUNG IN SPALTE T (falls noch nicht gesetzt)
    // =============================================================
    if (currentPfad === "") {
      if (formPersonen === 1 || formPersonen === 2) {
        // Prüfen, ob Gemini eine höhere Zahl als im Formular erkannt hat
        const hasHigherGeminiValue = !isNaN(geminiPersonen) && geminiPersonen > formPersonen;

        if (!hasHigherGeminiValue) {
          const pfadMarker = formPersonen === 1 ? "1P" : "2P";
          dmzSheet.getRange(rowNum, 20).setValue(pfadMarker); // Spalte T = Spalte 20
          Logger.log(`Zeile ${rowNum}: Pfad '${pfadMarker}' in Spalte T eingetragen.`);
        }
      }
    }

    // Status nach der Klassifizierung neu einlesen
    const updatedPfad = String(dmzSheet.getRange(rowNum, 20).getValue() || "").trim();

    // =============================================================
    // SCHRITT 2: ABSAGE-E-MAIL BEI 1P VERSENDEN
    // =============================================================
    if (updatedPfad === "1P" && emailStatus !== "ABSAGE GESENDET (1P)") {
      const gastName = row[2] || "Damen und Herren"; // Spalte C (Index 2)
      const id = row[1] || ""; // Spalte B (Index 1)

      let mailSubject = "Buchungsanfrage - Monteurwohnung L8 Street";
      if (id && id !== "N/A") {
        mailSubject += " (Anfrage ID: " + id + ")";
      }

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

      const mailBody = "Hallo " + gastName + ",\n\n" +
        "vielen Dank für Ihre Buchungsanfrage.\n\n" +
        "Bitte haben Sie Verständnis dafür, dass wir unsere Unterkünfte vorrangig an Gruppen vermieten und für Einzelpersonen leider keine passenden Kapazitäten anbieten können.\n\n" +
        "Auf unserer Website finden Sie eine Übersicht unserer Wohnungen sowie Verfügbarkeiten: a.L8Street.com\n\n" +
        "Vielen Dank für Ihr Verständnis." + footerText;

      try {
        GmailApp.sendEmail(testRecipient, mailSubject, mailBody);

        // Markiert die Zeile in Spalte AC (Spalte 29) als versendet
        dmzSheet.getRange(rowNum, 29).setValue("ABSAGE GESENDET (1P)");

        if (typeof logExecution === "function") {
          logExecution("DMZ Vorab-Filter", "ERFOLG", "Absage (1P) gesendet an " + testRecipient + " für: " + gastName, "");
        }
        Logger.log(`Zeile ${rowNum}: Absage-E-Mail (1P) erfolgreich an ${testRecipient} gesendet.`);
      } catch (err) {
        if (typeof logExecution === "function") {
          logExecution("DMZ Vorab-Filter", "FEHLER", "Fehler beim Absage-Versand (1P) für: " + gastName, err.toString());
        }
        Logger.log(`FEHLER Zeile ${rowNum}: ${err.toString()}`);
      }
    }
  }

  Logger.log("--- VORAB-FILTERUNG BEENDET ---");
}