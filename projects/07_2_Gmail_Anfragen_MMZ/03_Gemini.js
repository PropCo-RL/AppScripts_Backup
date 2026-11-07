// =================================================================
// GEMINI EXTRAKTOR FÜR TAB "MMZ" (Spalte B -> Puffer in Spalte C)
// Verwendet gemini-3.8-flash & interactions-Endpunkt
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
    "ZIMMER- & BELEGUNGSLOGIK (EXTREM WICHTIG):",
    " - Analysiere Freitext und Formular auf die gewünschte Zimmeraufteilung.",
    " - 'spezialData.belegungsTyp': Setze exakt 'EINZELZIMMER' wenn Einzelzimmer, getrennte Zimmer oder 1 Person pro Zimmer gewünscht sind. Setze 'WUNSCH_ZIMMER' wenn eine konkrete Zimmeranzahl genannt wird. Setze sonst 'STANDARD' (Standard = 2 Personen pro Zimmer).",
    " - 'spezialData.minZimmer': Die tatsächliche Mindestanzahl an separaten Schlafzimmern als reine Zahl (mindestens 1). Wenn Einzelzimmer: minZimmer = pers (oder höher, falls mehr gewünscht). Wenn konkrete Zimmerzahl gefordert: diese Zahl. Im Standardfall: aufgerundet (Pers / 2).",
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
      "minZimmer": 1,
      "belegungsTyp": "STANDARD",
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

function runGeminiExtractorMMZ() {
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MMZ');
  if (!sheet) {
    Logger.log("FEHLER: Sheet 'MMZ' nicht gefunden.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  ensureExtractorHeadersMMZ(sheet);

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

    // --- BYPASS: Wenn keine Nachricht oder nur "Keine Nachricht" vorhanden ist ---
    const msg = String(leadJson.nachricht || "").trim();
    if (msg === "" || msg.toLowerCase() === "keine nachricht") {
      Logger.log(`Zeile ${rowNum} (${leadJson.gastName || 'Gast'}): Keine Nachricht -> Schneller Formular-Bypass.`);
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
    Logger.log("Gemini Extraktor für MMZ erfolgreich abgeschlossen.");
  }
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN FÜR BYPASS & API
// -----------------------------------------------------------------
function buildFormularFallbackJSON(leadJson) {
  const dates = extractDatesFromFormular(leadJson.zeitraum);
  const persCount = Math.max(1, Number(leadJson.personen) || 1);
  const zimmerart = String(leadJson.zimmerart || "").toLowerCase();
  const titel = String(leadJson.unterkunftTitel || "").toLowerCase();

  // Prüfung 1: Einzelzimmer-Erkennung
  const istEinzelzimmer = zimmerart.includes("einzelzimmer") || titel.includes("einzelzimmer");

  // Prüfung 2: Explizite Zimmeranzahl im Titel oder der Zimmerart suchen (z.B. "3 Zimmer", "4-Zi", "2-Zimmer-Wohnung")
  let gefundeneZimmer = null;
  const combinedText = `${zimmerart} ${titel}`;
  const ziMatch = combinedText.match(/(\d+)\s*(?:[- ]?zimmer|[- ]?zi\b)/i);
  if (ziMatch && ziMatch[1]) {
    gefundeneZimmer = parseInt(ziMatch[1], 10);
  }

  // Ermittlung von minZimmer und belegungsTyp
  let minZimmer = Math.max(1, Math.ceil(persCount / 2));
  let belegungsTyp = "STANDARD";

  if (istEinzelzimmer) {
    minZimmer = persCount;
    belegungsTyp = "EINZELZIMMER";
  } else if (gefundeneZimmer && gefundeneZimmer > 0) {
    minZimmer = Math.max(minZimmer, gefundeneZimmer);
    belegungsTyp = "WUNSCH_ZIMMER";
  }

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
      zimmerAnzahl: gefundeneZimmer,
      minZimmer: Math.max(1, minZimmer),
      belegungsTyp: belegungsTyp,
      flags: {
        alleinnutzung: false,
        waschmaschine: false,
        wlan: false,
        parkplatz: false,
        einzelbetten: false,
        eigenesBad: false
      },
      direktEmail: leadJson.email || "",
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

function ensureExtractorHeadersMMZ(sheet) {
  const cell = sheet.getRange(1, 3);
  if (!cell.getValue()) {
    cell.setValue("Gemini Raw (JSON)");
    cell.setFontWeight("bold");
    sheet.getRange("C:C").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}