// =================================================================
// 2. GEMINI DATEN-KORREKTUR + EXTRAKTION (Sonderwünsche & Details)
// =================================================================
function processApartmentRequestsWithGemini() {
  const SHEET1_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');

  let s1;
  try {
    s1 = SpreadsheetApp.openById(SHEET1_ID).getSheetByName('DMZ');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabelle: " + e.toString());
    return;
  }

  if (!s1) {
    Logger.log("FEHLER: Tabellenblatt 'DMZ' nicht gefunden.");
    return;
  }

  const dmzData = s1.getDataRange().getValues();
  Logger.log("--- GEMINI ANFRAGE-PRÜFUNG & EXTRAKTION START ---");

  for (let i = 1; i < dmzData.length; i++) {
    const rowNum = i + 1;
    const currentStatus = String(dmzData[i][17] || "").trim(); // Spalte R (Index 17)
    const dupStatus = String(dmzData[i][18] || "").trim();     // Spalte S (Index 18)

    // A. Bereits verarbeitete Zeilen ODER Dopplungen überspringen
    if (currentStatus !== "") {
      continue;
    }
    if (dupStatus === "Dopplung") {
      s1.getRange(rowNum, 18).setValue("Übersprungen (Dopplung)");
      continue;
    }

    let rawLead = {
      gastName: String(dmzData[i][2] || "").trim(),      // Spalte C (Index 2)
      telefonnummer: String(dmzData[i][3] || "").trim(), // Spalte D (Index 3)
      zeitraum: String(dmzData[i][4] || "").trim(),      // Spalte E (Index 4)
      personen: String(dmzData[i][5] || "").trim(),      // Spalte F (Index 5)
      nachricht: String(dmzData[i][8] || "").trim(),     // Spalte I (Index 8)
      addressStr: String(dmzData[i][9] || "").trim(),    // Spalte J (Index 9)
      email: String(dmzData[i][11] || "").trim()         // Spalte L (Index 11)
    };

    if (!rawLead.nachricht && !rawLead.zeitraum) {
      s1.getRange(rowNum, 18).setValue("Fehler: Keine Daten/Nachricht");
      continue;
    }

    Logger.log("Verarbeite Zeile " + rowNum + " für " + (rawLead.gastName || "Unbekannt"));
    let verifiedData = verifyAndUpdateLeadWithGemini(rawLead, GEMINI_API_KEY);

    if (verifiedData) {
      // B. Gemini-Daten in die separaten Spalten L bis Q schreiben (C, D, E, F bleiben unverändert)
      s1.getRange(rowNum, 12).setValue(verifiedData.email || rawLead.email);                 // Spalte L (12): E-Mail
      s1.getRange(rowNum, 13).setValue(verifiedData.standort || "");                         // Spalte M (13): Gemini Standort aus Nachricht/Adresse
      s1.getRange(rowNum, 14).setValue(verifiedData.personen || rawLead.personen);         // Spalte N (14): Gemini Personen
      s1.getRange(rowNum, 15).setValue(verifiedData.zeitraum || rawLead.zeitraum);         // Spalte O (15): Gemini Zeitraum
      s1.getRange(rowNum, 16).setValue(verifiedData.gastName || rawLead.gastName);         // Spalte P (16): Gemini Name
      s1.getRange(rowNum, 17).setValue(verifiedData.telefonnummer || rawLead.telefonnummer); // Spalte Q (17): Gemini Telefon

      // C. Detaillierte Änderungen für das Logging in Spalte R ermitteln
      let changes = [];
      if (verifiedData.gastName && verifiedData.gastName !== rawLead.gastName) {
        changes.push(`Name (${rawLead.gastName || 'leer'} -> ${verifiedData.gastName})`);
      }
      if (verifiedData.telefonnummer && verifiedData.telefonnummer !== rawLead.telefonnummer) {
        changes.push(`Telefon (${rawLead.telefonnummer || 'leer'} -> ${verifiedData.telefonnummer})`);
      }
      if (verifiedData.zeitraum && verifiedData.zeitraum !== rawLead.zeitraum) {
        changes.push(`Zeitraum (${rawLead.zeitraum || 'leer'} -> ${verifiedData.zeitraum})`);
      }
      if (String(verifiedData.personen) !== String(rawLead.personen)) {
        changes.push(`Personen (${rawLead.personen || 'leer'} -> ${verifiedData.personen})`);
      }
      if (verifiedData.email && verifiedData.email !== rawLead.email) {
        changes.push(`E-Mail (${rawLead.email || 'leer'} -> ${verifiedData.email})`);
      }
      if (verifiedData.standort) {
        changes.push(`Standort (${verifiedData.standort})`);
      }

      // D. Detailliertes Status-Log in Spalte R schreiben
      let statusLog = changes.length > 0 ? "Korrektur: " + changes.join(", ") : "Keine Änderung";
      s1.getRange(rowNum, 18).setValue(statusLog); // Spalte R

      // E. Erweiterte Extraktionen in Spalten U, V, W, X, Y eintragen
      s1.getRange(rowNum, 21).setValue(verifiedData.rueckfragen || "Keine");               // Spalte U (21)
      s1.getRange(rowNum, 22).setValue(verifiedData.sonderwuensche || "Keine");             // Spalte V (22)
      s1.getRange(rowNum, 23).setValue(verifiedData.einzelzimmerWunsch || "Nein");          // Spalte W (23)
      s1.getRange(rowNum, 24).setValue(verifiedData.optionAufVerlaengerung || "Nein");      // Spalte X (24)
      s1.getRange(rowNum, 25).setValue(verifiedData.abrechnungsZyklus || "Keine Angabe");  // Spalte Y (25)
    } else {
      s1.getRange(rowNum, 18).setValue("Fehler: Gemini API");
    }
  }
  Logger.log("--- GEMINI ANFRAGE-PRÜFUNG BEENDET ---");
}

// =================================================================
// GEMINI PROMPT & API CALL
// =================================================================
function verifyAndUpdateLeadWithGemini(rawLead, apiKey) {
  const promptText = `
Du bist ein präziser Daten-Korrektur- und Analyse-Assistent für Buchungsanfragen von Monteurunterkünften.
Analysiere die Eingangsdaten und vergleiche sie mit dem Feld "nachricht" sowie "addressStr".

REGELN FÜR STAMMDATEN-KORREKTUR & EXTRAKTION:
1. Das Feld "nachricht" enthält die absolute WAHRHEIT.
2. Wenn in der Nachricht korrigierte oder abweichende Angaben zu Name, Telefonnummer, Zeitraum, Personenanzahl oder E-Mail stehen, korrigiere diese.
3. Falls "email" leer ist, prüfe ob in "nachricht" eine E-Mail genannt wird.
4. "standort": Extrahiere den Ziel-Standort bevorzugt aus "nachricht". Achte dabei besonders auf genaue Orts- oder Standortbeschreibungen, Industriegebiete, Firmen oder relative Angaben (z.B. "Nähe Audi Werk Neckarsulm", "Raum Stuttgart", "Baustelle in Karlsruhe", "Nähe Flughafen"). Falls in der Nachricht überhaupt kein Ort/Standort erwähnt wird, nutze den Wert aus "addressStr".

REGELN FÜR DETAIL-EXTRAKTIONEN:
5. "rueckfragen": Kommagetrennte Fragen des Kunden (z.B. "Check-In Zeit", "Rechnungsstellung", "Haustiere erlaubt?"). Falls keine Fragen: "Keine".
6. "sonderwuensche": Kommagetrennte Liste von Wünschen (Achte besonders auf: "WLAN", "Parkplatz/Transporter", "Keine Durchgangszimmer", "Getrennte Betten", "Waschmaschine", "Balkon", "Erdgeschoss"). Falls keine: "Keine".
7. "einzelzimmerWunsch": Falls Zimmeraufteilung oder Einzelzimmer gewünscht sind, liefere die genauen Details (z.B. "2 DZ + 1 EZ", "Nur Einzelzimmer" oder "Nein").
8. "optionAufVerlaengerung": Falls erwähnt, gib an ob und wie lange (z.B. "Ja (bis Ende November)", "Ja (voraussichtlich +2 Wochen)" oder "Nein").
9. "abrechnungsZyklus": Falls der Kunde Wünsche zur Abrechnung/Zahlungsintervall nennt (z.B. "Erstmal 7 Tage", "Monatsweise (30 Tage)", "14-tägig", "Bis 31.12."). Falls nicht erwähnt: "Keine Angabe".

Eingangs-Daten:
"""${JSON.stringify(rawLead)}"""

ANTWORTE AUSSCHLIESSLICH ALS EIN EINZELNES JSON-OBJEKT IN FOLGENDEM FORMAT (KEIN MARKDOWN, KEIN EXTRA TEXT):
{
  "gastName": "Korrigierter oder ursprünglicher Name",
  "telefonnummer": "Korrigierte oder ursprüngliche Telefonnummer",
  "standort": "Extrahierter Standort/Ort aus Nachricht oder Adresse",
  "zeitraum": "Korrigierter oder ursprünglicher Zeitraum",
  "personen": 3,
  "email": "Korrigierte oder ergänzte E-Mail-Adresse",
  "rueckfragen": "Kommagetrennte Fragen oder Keine",
  "sonderwuensche": "Kommagetrennte Wünsche oder Keine",
  "einzelzimmerWunsch": "Genaue Zimmeraufteilung oder Nein",
  "optionAufVerlaengerung": "Ja mit Detail oder Nein",
  "abrechnungsZyklus": "Intervall/Datum oder Keine Angabe"
}`;

  return executeGeminiInteraction(promptText, apiKey);
}

function executeGeminiInteraction(inputText, apiKey) {
  if (!apiKey) {
    Logger.log("FEHLER: Kein Gemini API Key in den Script Properties gefunden ('Gemini-Key').");
    return null;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/interactions`;
  const payload = {
    "model": "gemini-3.8-flash",
    "input": inputText,
    "store": false
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      Logger.log("Gemini API Fehler Code " + response.getResponseCode() + ": " + response.getContentText());
      return null;
    }

    const json = JSON.parse(response.getContentText());
    let rawText = "";
    if (json.output_text) {
      rawText = json.output_text.trim();
    } else if (json.steps && json.steps.length > 0) {
      const lastStep = json.steps[json.steps.length - 1];
      if (lastStep.content && lastStep.content[0] && lastStep.content[0].text) {
        rawText = lastStep.content[0].text.trim();
      }
    }

    const jsonMatch = rawText.match(/\[[\s\S]*\]|\{[\s\S]*\}/);
    if (jsonMatch) rawText = jsonMatch[0];
    return safeParseJSON(rawText);
  } catch (e) {
    Logger.log("Fehler bei Gemini API Aufruf: " + e.toString());
  }
  return null;
}

function safeParseJSON(rawText) {
  try {
    return JSON.parse(rawText);
  } catch (e) {
    try {
      let cleanText = rawText
        .replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
        .replace(/\r?\n/g, " ")
        .replace(/",\s*}/g, '"}')
        .replace(/",\s*]/g, '"]');
      return JSON.parse(cleanText);
    } catch (e2) {
      return null;
    }
  }
}