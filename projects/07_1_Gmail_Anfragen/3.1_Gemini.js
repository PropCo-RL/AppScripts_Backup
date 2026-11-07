// =================================================================
// SKRIPT 1: GEMINI EXTRAKTOR (Spalte B -> Puffer in Spalte C)
// Vollständig mit No-Message-Bypass & Zeilenhöhen-Fix
// =================================================================

const GEMINI_CONFIG = {
  MODEL_NAME: "gemini-3.8-flash",
  PROMPT_SYSTEM_ROLE: "Du bist ein extrem präziser Buchungs-Analyst für Monteurzimmer-Vermietung.",
  INSTRUCTIONS: [
    "NACHRICHTEN-FREITEXT GEWINNT IMMER: Formularfelder werden ignoriert, wenn der Freitext abweicht.",
    "ZITATE-PFLICHT (ANTI-HALLUZINATION):",
    " - Zitiere unter 'zitate.personenFundstelle' die exakte Textstelle zur Personenzahl aus dem Freitext (oder 'Nur im Formular').",
    " - Zitiere unter 'zitate.zeitraumFundstelle' die exakte Textstelle zu Daten/Zeitraum aus dem Freitext (oder 'Nur im Formular').",
    "SCORES:",
    " - 'kiConfidenceScore' (0-100): Wie sicher bist du, dass du den Text grammatikalisch/inhaltlich fehlerfrei verstanden hast? (Klares Deutsch = 95-100; Kauderwelsch < 80).",
    " - 'gastConfidenceScore' (0-100): 100 = Feste Daten/Personen; 70-85 = 'ca.', '3-4 Pers.', 'Mitte des Monats' (gastIstVage: true); < 70 = Reine Schätzungen/Open-End.",
    "suchortText: Konkreter Zielort/Baustelle/Arbeitsort aus dem Freitext (z.B. Eberdingen, Jena, Weil der Stadt). Falls kein abweichender Ort genannt: leer lassen.",
    "personenData: 'Pers' = Startanzahl als reine Zahl (bei Spannen die untere Zahl). 'max' = geplante Aufstockung/Maximalzahl. 'aufteilung' & 'zimmer' aus Text übernehmen.",
    "zeitraumData.start / ende: Im Format YYYY-MM-DD. Zähle KEINE Tage.",
    "zeitraumData.maxTage: Geschätzte Gesamtdauer in Tagen NUR WENN Projekt/Verlängerung/Monatswunsch genannt (z.B. 30, 90, 180). Falls keine Verlängerung erwähnt: null.",
    "zeitraumData.abrechnung: Abrechnungs- oder Buchungszyklus als reine Zahl in Tagen (z.B. 7 bei wöchentlich, 14 bei zweiwöchentlich, 30 bei monatlich). Sonst null.",
    "zeitraumData.verlaengerung: true (wenn Option auf Verlängerung/Projekt erwähnt), sonst false.",
    "zeitraumData.muster: 'mo_do' (falls nur Mo-Do gewünscht), sonst 'durchgehend'.",
    "zeitraumData.info: Kurzer Text zur Verlängerungsoption/Projekt, sonst leer.",
    "spezialData.flags: Exakte Booleans für alleinnutzung, waschmaschine, wlan, parkplatz, einzelbetten, eigenesBad.",
    "spezialData.direktEmail: Echte Kontakt-E-Mail aus dem Freitext/Signatur (falls vorhanden).",
    "spezialData.kontaktKanal: 'nur_email' (falls keine Anrufe/WhatsApp erlaubt), 'whatsapp' (falls WhatsApp gewünscht), sonst 'egal'.",
    "spezialData.unserZeichen: Aktenzeichen/Projektnummer des Gastes, sonst leer.",
    "spezialData.budgetMonatGesamt: Budgetangabe als reine Zahl, sonst null.",
    "spezialData.fragenUndSonstiges: Alle sonstigen Wünsche, Fragen oder Notizen aus dem Text."
  ],
  OUTPUT_SCHEMA: {
    "kiConfidenceScore": 95,
    "zitate": {
      "personenFundstelle": "",
      "zeitraumFundstelle": ""
    },
    "gastConfidenceScore": 90,
    "gastIstVage": false,
    "vageGruende": "",
    "suchortText": "",
    "personenData": {
      "Pers": 1,
      "max": 1,
      "aufteilung": "",
      "zimmer": ""
    },
    "zeitraumData": {
      "start": "YYYY-MM-DD",
      "ende": "YYYY-MM-DD",
      "maxTage": null,
      "abrechnung": null,
      "verlaengerung": false,
      "muster": "durchgehend",
      "info": ""
    },
    "spezialData": {
      "unterkunftsTyp": "",
      "zimmerAnzahl": null,
      "flags": {
        "alleinnutzung": false,
        "waschmaschine": false,
        "wlan": false,
        "parkplatz": false,
        "einzelbetten": false,
        "eigenesBad": false
      },
      "direktEmail": "",
      "kontaktKanal": "egal",
      "unserZeichen": "",
      "budgetMonatGesamt": null,
      "fragenUndSonstiges": ""
    },
    "istStaffelAnfrage": false,
    "zeitraumSegmente": []
  }
};

function runGeminiExtractorDMZ2() {
  const SHEET_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DMZ2');
  if (!sheet) {
    Logger.log("FEHLER: Sheet 'DMZ2' nicht gefunden.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  ensureExtractorHeadersDMZ2(sheet);

  const numRows = lastRow - 1;
  const range = sheet.getRange(2, 1, numRows, 3);
  const rows = range.getValues();
  let updatesMade = false;

  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    const cellB = String(rows[i][1] || "").trim(); // Spalte B: Bereinigter Lead
    const cellC = String(rows[i][2] || "").trim(); // Spalte C: Gemini-Puffer

    if (!cellB || cellB === "-" || cellC !== "") {
      continue;
    }

    let leadJson;
    try {
      leadJson = JSON.parse(cellB);
    } catch (e) {
      rows[i][2] = "Fehler: Rohdaten-JSON ungültig";
      updatesMade = true;
      continue;
    }

    // --- BYPASS: Nur wenn WIRKLICH GAR KEINE Nachricht da ist ---
    const msg = String(leadJson.nachricht || "").trim();
    if (msg === "") {
      Logger.log(`Zeile ${rowNum} (${leadJson.gastName || 'Gast'}): Nachricht komplett leer -> Schneller Formular-Bypass.`);
      const fallbackResult = buildFormularFallbackJSON(leadJson);
      rows[i][2] = JSON.stringify(fallbackResult);
      updatesMade = true;
      continue;
    }

    // Sobald Text existiert -> Mit Gemini analysieren
    if (!GEMINI_API_KEY) {
      Logger.log("FEHLER: Kein Gemini API Key vorhanden.");
      return;
    }

    Logger.log(`Extrahiere Zeile ${rowNum} (${leadJson.gastName || 'Gast'}) mit Gemini...`);
    const resultJson = executeGeminiInteractionsCall(leadJson, GEMINI_API_KEY);

    if (resultJson) {
      resultJson.bypassOhneNachricht = false;
      rows[i][2] = JSON.stringify(resultJson);
    } else {
      rows[i][2] = "Fehler: API-Aufruf fehlgeschlagen";
    }
    updatesMade = true;
  }

  if (updatesMade) {
    range.setValues(rows);
    for (let r = 2; r <= lastRow; r++) {
      sheet.setRowHeight(r, 21);
    }
    Logger.log("Gemini Extraktor erfolgreich abgeschlossen.");
  }
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN FÜR BYPASS & API
// -----------------------------------------------------------------
function buildFormularFallbackJSON(leadJson) {
  const dates = extractDatesFromFormular(leadJson.zeitraum);
  const persCount = Number(leadJson.personen) || 1;

  return {
    bypassOhneNachricht: true,
    kiConfidenceScore: 100,
    zitate: {
      personenFundstelle: "Nur im Formular",
      zeitraumFundstelle: "Nur im Formular"
    },
    gastConfidenceScore: 100,
    gastIstVage: false,
    vageGruende: "",
    suchortText: "",
    personenData: {
      Pers: persCount,
      max: persCount,
      aufteilung: "",
      zimmer: String(leadJson.unterkunftTitel || "")
    },
    zeitraumData: {
      start: dates.start,
      ende: dates.ende,
      maxTage: null,
      abrechnung: null,
      verlaengerung: false,
      muster: "durchgehend",
      info: ""
    },
    spezialData: {
      unterkunftsTyp: String(leadJson.unterkunftTitel || ""),
      zimmerAnzahl: null,
      flags: {
        alleinnutzung: false,
        waschmaschine: false,
        wlan: false,
        parkplatz: false,
        einzelbetten: false,
        eigenesBad: false
      },
      direktEmail: "",
      kontaktKanal: "egal",
      unserZeichen: "",
      budgetMonatGesamt: null,
      fragenUndSonstiges: ""
    },
    istStaffelAnfrage: false,
    zeitraumSegmente: []
  };
}

function extractDatesFromFormular(zeitraumStr) {
  if (!zeitraumStr) return { start: "", ende: "" };
  const m = zeitraumStr.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})\s*-\s*(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (m) {
    const sTag = m[1].padStart(2, '0');
    const sMon = m[2].padStart(2, '0');
    const sJahr = m[3];
    const eTag = m[4].padStart(2, '0');
    const eMon = m[5].padStart(2, '0');
    const eJahr = m[6];
    return {
      start: `${sJahr}-${sMon}-${sTag}`,
      ende: `${eJahr}-${eMon}-${eTag}`
    };
  }
  return { start: "", ende: "" };
}

function executeGeminiInteractionsCall(leadJson, apiKey) {
  const todayIso = new Date().toISOString().split("T")[0];

  const dynamicInstructions = [
    `HEUTIGES BEZUGSDATUM: ${todayIso}. Verwende dieses Datum und Jahr als Referenz für relative Angaben.`,
    ...GEMINI_CONFIG.INSTRUCTIONS
  ];

  const instructionsList = dynamicInstructions.map((instr, idx) => `${idx + 1}. ${instr}`).join("\n");
  const schemaString = JSON.stringify(GEMINI_CONFIG.OUTPUT_SCHEMA, null, 2);

  const promptText = `
${GEMINI_CONFIG.PROMPT_SYSTEM_ROLE}
Analysiere die Eingangsdaten und vergleiche die Formular-Felder mit dem Freitext ("nachricht").

REGELN & INSTRUKTIONEN:
${instructionsList}

Eingangsdaten:
"""${JSON.stringify(leadJson)}"""

ANTWORTE AUSSCHLIESSLICH ALS EIN VALIDES JSON-OBJEKT ENTSPRECHEND FOLGENDER STRUKTUR (KEIN MARKDOWN, KEINE CODE-BLÖCKE):
${schemaString}
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/interactions`;
  const payload = {
    "model": GEMINI_CONFIG.MODEL_NAME,
    "input": promptText,
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
      Logger.log("API Fehler " + response.getResponseCode() + ": " + response.getContentText());
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
    Logger.log("Fehler bei Gemini API: " + e.toString());
    return null;
  }
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

function ensureExtractorHeadersDMZ2(sheet) {
  const cell = sheet.getRange(1, 3);
  if (!cell.getValue()) {
    cell.setValue("Gemini Raw (JSON)");
    cell.setFontWeight("bold");
    sheet.getRange("C:C").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}