// =================================================================
// SCRIPT 5: KUNDENANTWORTEN MIT GEMINI ANALYSIEREN
// =================================================================
function processRepliesWithGemini() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName('output');

  if (!outputSheet) {
    Logger.log("FEHLER: Tabellenblatt 'output' nicht gefunden.");
    return;
  }

  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  if (!GEMINI_API_KEY) {
    Logger.log("FEHLER: Kein Gemini API Key in den Script Properties gefunden ('Gemini-Key').");
    return;
  }

  const lastRow = outputSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("Keine Daten im Tab 'output' vorhanden.");
    return;
  }

  const outputData = outputSheet.getDataRange().getValues();
  Logger.log("--- GEMINI ANTWORT-ANALYSE START ---");

  for (let i = 2; i < outputData.length; i++) {
    const rowNum = i + 1;
    
    const replyText = String(outputData[i][18] || "").trim();    // Spalte S (Index 18): Kundenantwort
    const currentStatus = String(outputData[i][23] || "").trim(); // Spalte X (Index 23): Gemini Status

    // A. Bereits verarbeitete Zeilen ODER leere Antworten überspringen
    if (currentStatus !== "" || replyText === "") {
      continue;
    }

    const customerData = {
      gastName: String(outputData[i][0] || "").trim(),      // Spalte A
      apartment: String(outputData[i][1] || "").trim(),     // Spalte B
      geplantesAbreiseDatum: String(outputData[i][2] || "").trim(), // Spalte C: Bisheriges Abreisedatum
      email: String(outputData[i][3] || "").trim(),         // Spalte D
      buchungsNr: String(outputData[i][5] || "").trim(),    // Spalte F
      antwortText: replyText                                // Spalte S
    };

    Logger.log("Verarbeite Zeile " + rowNum + " für " + (customerData.gastName || "Unbekannt"));

    const verifiedData = analyzeCustomerReplyWithGemini(customerData, GEMINI_API_KEY);

    if (verifiedData) {
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');

      // B. Ergebnisse in die Spalten X, Y und Z schreiben
      outputSheet.getRange(rowNum, 24).setValue(`Analysiert (${timestamp})`);                 // Spalte X (24): Status
      outputSheet.getRange(rowNum, 25).setValue(verifiedData.entscheidung || "Keine Verlängerung"); // Spalte Y (25): Datum, Zahl oder Status
      outputSheet.getRange(rowNum, 26).setValue(verifiedData.begruendung || "");              // Spalte Z (26): Begründung

      Logger.log(`Zeile ${rowNum} erfolgreich analysiert: ${verifiedData.entscheidung}`);
    } else {
      outputSheet.getRange(rowNum, 24).setValue("Fehler: Gemini API");
    }
  }

  Logger.log("--- GEMINI ANTWORT-ANALYSE BEENDET ---");
}

// =================================================================
// GEMINI PROMPT & API CALL
// =================================================================
function analyzeCustomerReplyWithGemini(customerData, apiKey) {
  const promptText = `
Du bist ein präziser Analyse-Assistent für Antworten von Gästen/Monteuren bezüglich ihrer geplanten Abreise aus einer Unterkunft.
Analysiere die Kundenantwort ("antwortText") im Kontext des bisherigen geplanten Abreisedatums ("geplantesAbreiseDatum").

REGELN FÜR DEN WERT IN "entscheidung":
Gib im Feld "entscheidung" AUSSCHLIESSLICH einen der folgenden Werte zurück (kein zusätzlicher Text!):

1. Ein konkretes Datum im Format "DD.MM.YYYY":
   - Wenn ein konkretes Datum genannt wird ODER eine Aussage wie "bis Ende des Monats" / "bis Monatsende" gemacht wird, rechne dies in ein exaktes Datum um (z. B. "31.10.2026").

2. Eine reine Zahl (z. B. "7" oder "30"):
   - Wenn der Kunde eine Zeitspanne nennt wie "1 Woche", "noch 10 Tage" oder "einen Monat", rechne dies in reine TAGE um und gib NUR die Zahl aus (z. B. "7" bei einer Woche, "30" bei einem Monat).

3. "Ohne Datum":
   - Wenn der Kunde fest verlängern möchte, aber absolut keine Dauer, Tage oder Daten nennt.

4. "Unentschlossen":
   - Wenn der Kunde unsicher ist (z. B. "weiß noch nicht", "muss Chef fragen", "melden uns später").

5. "Keine Verlängerung":
   - Wenn der Kunde nicht verlängert oder seine Abreise wie geplant bestätigt.

Gib im Feld "begruendung" eine kurze Erklärung (1 Satz auf Deutsch), wie du die Antwort interpretiert hast.

Eingangs-Daten:
"""${JSON.stringify(customerData)}"""

ANTWORTE AUSSCHLIESSLICH ALS EIN EINZELNES JSON-OBJEKT IN FOLGENDEM FORMAT (KEIN MARKDOWN, KEIN EXTRA TEXT):
{
  "entscheidung": "DD.MM.YYYY ODER reine Zahl ODER Ohne Datum ODER Unentschlossen ODER Keine Verlängerung",
  "begruendung": "Kurze Begründung für den extrahierten Wert"
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