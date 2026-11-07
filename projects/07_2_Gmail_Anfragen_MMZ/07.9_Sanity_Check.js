// =================================================================
// STUFE 4: VOLLSTÄNDIGER SANITY-CHECK & RETTUNGS-AUDITOR (MMZ)
// Liest dynamisch alle Spalten ab B ein, ermittelt Wohnungen & Buchungen
// und schreibt das Audit-Ergebnis als JSON in Spalte I ("Sanity Check (JSON)")
// Verwendet gemini-3.8-flash & interactions-Endpunkt
// =================================================================

const GEMINI_AUDIT_CONFIG = {
  MODEL_NAME: "gemini-3.8-flash",
  PROMPT_SYSTEM_ROLE: "Du bist der finale Chef-Revisor und Qualitätsprüfer für Monteurzimmer-Buchungsangebote."
};

function runDynamicSanityCheckMMZ() {
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  const SPREADSHEET_INVENTORY_ID = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';
  const SHEET_NAME = 'MMZ';

  if (!GEMINI_API_KEY) {
    Logger.log("FEHLER: Kein Gemini API Key hinterlegt.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`FEHLER: Sheet '${SHEET_NAME}' nicht gefunden.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return;

  // Header-Zeile prüfen / Spalte I für Sanity Check sicherstellen
  const headers = sheet.getRange(1, 1, 1, Math.max(lastCol, 9)).getValues()[0];
  let auditColIdx = headers.indexOf("Sanity Check (JSON)") + 1;
  if (auditColIdx === 0) {
    auditColIdx = 9; // Spalte I als Standard
    sheet.getRange(1, auditColIdx).setValue("Sanity Check (JSON)").setFontWeight("bold");
    sheet.getRange(1, auditColIdx, sheet.getMaxRows(), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Inventar und Kalender laden
  let invSheetM, invSheetCal;
  try {
    const invSS = SpreadsheetApp.openById(SPREADSHEET_INVENTORY_ID);
    invSheetM = invSS.getSheetByName('M');
    invSheetCal = invSS.getSheetByName('cal');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen des Inventory-Sheets: " + e.toString());
    return;
  }

  if (!invSheetM || !invSheetCal) {
    Logger.log("FEHLER: Tabs 'M' oder 'cal' im Inventory-Sheet fehlen.");
    return;
  }

  const apartments = loadApartmentsForAudit(invSheetM);
  const bookingsMap = loadBookingsForAudit(invSheetCal);

  // Alle relevanten Datenzeilen einlesen (bis zur Audit-Spalte)
  const numRows = lastRow - 1;
  const dataRange = sheet.getRange(2, 1, numRows, auditColIdx);
  const rows = dataRange.getValues();
  let updatesMade = false;

  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    const cellB = String(rows[i][1] || "").trim(); // Spalte B: Lead
    const existingAudit = String(rows[i][auditColIdx - 1] || "").trim();

    // Überspringen wenn: kein Lead da, Duplikat "-" oder Audit bereits vorhanden
    if (!cellB || cellB === "-" || existingAudit !== "") {
      continue;
    }

    // Dynamisch alle Spalten ab B (Index 1) bis vor die Audit-Spalte sammeln
    const leadGesamtKontext = {};
    for (let c = 1; c < auditColIdx - 1; c++) {
      const headerName = headers[c] || `Spalte_${c + 1}`;
      const cellVal = String(rows[i][c] || "").trim();
      try {
        leadGesamtKontext[headerName] = JSON.parse(cellVal);
      } catch (e) {
        leadGesamtKontext[headerName] = cellVal;
      }
    }

    // Suchort & Zeitraum für den Kalender-Snapshot bestimmen
    const rohOrt = leadGesamtKontext["Validiert (JSON)"]?.ort?.suchortText 
                || leadGesamtKontext["Validiert (JSON)"]?.ort?.Ort 
                || leadGesamtKontext["Lead (JSON)"]?.unterkunftStadtHtml 
                || leadGesamtKontext["Lead (JSON)"]?.unterkunftStadtTitel || "";
    const suchStadt = extractPureCityName(rohOrt);

    const startStr = leadGesamtKontext["Routing Status (JSON)"]?.kennzahlen?.start 
                  || leadGesamtKontext["Validiert (JSON)"]?.dauer?.start || "";
    const endeStr = leadGesamtKontext["Routing Status (JSON)"]?.kennzahlen?.ende 
                 || leadGesamtKontext["Validiert (JSON)"]?.dauer?.ende || "";

    // Kalender-Snapshot mit Buchungsintervallen & Toleranzfenster holen
    const verfuegbarkeitsSnapshot = erstelleDetailliertenWohnungsSnapshot(suchStadt, startStr, endeStr, apartments, bookingsMap);

    Logger.log(`Sanity-Check Zeile ${rowNum} für Ort '${suchStadt}' (${startStr} bis ${endeStr})...`);

    const auditRes = executeGeminiFullAuditCall(leadGesamtKontext, verfuegbarkeitsSnapshot, GEMINI_API_KEY);
    if (auditRes) {
      auditRes.geprueftAm = new Date().toISOString();
      rows[i][auditColIdx - 1] = JSON.stringify(auditRes);
    } else {
      rows[i][auditColIdx - 1] = JSON.stringify({
        freigabe: false,
        status: "FEHLER_API",
        begruendung: "Gemini Audit-Aufruf fehlgeschlagen"
      });
    }
    updatesMade = true;
  }

  if (updatesMade) {
    dataRange.setValues(rows);
    for (let r = 2; r <= lastRow; r++) {
      sheet.setRowHeight(r, 21);
    }
    Logger.log("Dynamischer Sanity-Check erfolgreich abgeschlossen.");
  }
}

// -----------------------------------------------------------------
// GEMINI CALL MIT INTERACTION-ENDPUNKT & KOMPLETTEM ZEILEN-DUMP
// -----------------------------------------------------------------
function executeGeminiFullAuditCall(alleSpaltenAbB, verfuegbarkeitsSnapshot, apiKey) {
  const promptText = `
${GEMINI_AUDIT_CONFIG.PROMPT_SYSTEM_ROLE}

Hier ist der VOLLSTÄNDIGE Datensatz einer Buchungsanfrage (alle Pipeline-Stufen von Spalte B bis zum generierten E-Mail-Draft) sowie der echte Kalender-Snapshot der relevanten Wohnungen mit bestehenden Buchungen (inklusive Puffer-Zeitfenster).

1. GESAMTER DATENSATZ DER ANFRAGE:
${JSON.stringify(alleSpaltenAbB, null, 2)}

2. WOHNUNGEN IN DER STADT & BESTEHENDE BUCHUNGEN (AUS TAB 'M' & 'CAL'):
${JSON.stringify(verfuegbarkeitsSnapshot, null, 2)}

REGELN & INSTRUKTIONEN FÜR DEINE PRÜFUNG:
1. VERLORENER UMSATZ / FALSCHE ABSAGE:
   - Wurde ein Draft mit Absage ("Keine Verfügbarkeit") generiert?
   - Prüfe anhand der "bestehendeBuchungen" der Wohnungen, ob im angefragten Zeitraum (oder bei vagen Angaben mit wenigen Tagen Versatz) Kapazitäten frei waren.
   - Hätten wir freie Einheiten für ein Vollangebot ODER ein wirtschaftliches Teilangebot gehabt?
   -> Wenn JA: "freigabe": false, "status": "ALARM_FALSCHE_ABSAGE".
2. KOLLISION / DOPPELBUCHUNG / ÜBERBELEGUNG:
   - Kollidiert eine im Angebot genannte Wohnung zeitlich mit einer bestehenden Buchung?
   - Reichen Betten und Zimmer für die im Text versprochene Personenzahl?
   -> Wenn JA: "freigabe": false, "status": "ALARM_KOLLISION_ODER_UEBERBELEGUNG".
3. ZEITRAUM-KONSISTENZ:
   - Passt der Zeitraum im Draft zur Anfrage (unter Beachtung des Mindestaufenthalts von 6 Tagen bei Verlängerungswunsch)?
   - Wenn ein Teilangebot vorliegt: Wird dem Gast transparent erklärt, wie viele Personen Platz finden?
4. FREIGABE:
   - Wenn Angebot oder Absage sachlich, mathematisch und kalendarisch vollkommen korrekt sind: "freigabe": true, "status": "PASS".

ANTWORTE AUSSCHLIESSLICH ALS EIN VALIDES JSON-OBJEKT ENTSPRECHEND FOLGENDER STRUKTUR (KEIN MARKDOWN, KEINE CODE-BLÖCKE):
{
  "freigabe": true,
  "status": "PASS",
  "risikoScore": 0,
  "moeglicheAlternativeApts": [],
  "begruendung": "Kurze Begründung"
}
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/interactions`;
  const payload = {
    "model": GEMINI_AUDIT_CONFIG.MODEL_NAME,
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
      Logger.log("API Fehler Audit " + response.getResponseCode() + ": " + response.getContentText());
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
    Logger.log("Fehler bei Gemini Audit API: " + e.toString());
    return null;
  }
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN FÜR INVENTAR & DETAILLIERTEN SNAPSHOT
// -----------------------------------------------------------------
function loadApartmentsForAudit(invSheetM) {
  const mData = invSheetM.getDataRange().getValues();
  const list = [];
  for (let r = 1; r < mData.length; r++) {
    const name = String(mData[r][0] || "").trim();
    if (!name) continue;
    const rawShort = String(mData[r][6] || "").trim();
    list.push({
      name: name,
      code: cleanShortCode(rawShort) || name,
      schlafzimmer: Number(mData[r][14]) || 1,
      maxPersonen: Number(mData[r][15]) || 2,
      regionen: String(mData[r][22] || "").toLowerCase(),
      dauerhaftBelegt: String(mData[r][44] || "").trim().toLowerCase() === "x"
    });
  }
  return list;
}

function loadBookingsForAudit(invSheetCal) {
  const calData = invSheetCal.getDataRange().getValues();
  const map = {};
  for (let c = 1; c < calData.length; c++) {
    const prop = String(calData[c][0] || "").trim().toLowerCase();
    const s = parseIsoDateString(calData[c][1]);
    const e = parseIsoDateString(calData[c][2]);
    if (!prop || !s || !e) continue;
    if (!map[prop]) map[prop] = [];
    map[prop].push({
      s: s,
      e: e,
      sMs: new Date(s + "T00:00:00Z").getTime(),
      eMs: new Date(e + "T00:00:00Z").getTime()
    });
  }
  return map;
}

function erstelleDetailliertenWohnungsSnapshot(suchStadt, anfrageStartStr, anfrageEndeStr, apartments, bookingsMap) {
  if (!suchStadt) return [];

  // Relevantes Zeitfenster berechnen (7 Tage vor Start bis 14 Tage nach Ende)
  let windowStartMs = 0;
  let windowEndMs = Infinity;

  if (anfrageStartStr) {
    const sDate = new Date(anfrageStartStr + "T00:00:00Z");
    if (!isNaN(sDate.getTime())) {
      windowStartMs = sDate.getTime() - (7 * 86400000);
    }
  }

  if (anfrageEndeStr) {
    const eDate = new Date(anfrageEndeStr + "T00:00:00Z");
    if (!isNaN(eDate.getTime())) {
      windowEndMs = eDate.getTime() + (14 * 86400000);
    }
  } else if (windowStartMs > 0) {
    windowEndMs = windowStartMs + (37 * 86400000); // 30 Tage Fallback ab Start
  }

  // 1. Nur Wohnungen der Zielstadt filtern
  const stadtApts = apartments.filter(a => {
    const r = a.regionen || "";
    return r.includes(suchStadt) || suchStadt.includes(r);
  });

  // 2. Bestehende Buchungen im erweiterten Puffer-Zeitfenster sammeln
  return stadtApts.map(apt => {
    const aptKey = apt.name.toLowerCase();
    const allBookings = bookingsMap[aptKey] || [];

    const relevanteBuchungen = allBookings
      .filter(b => b.eMs > windowStartMs && b.sMs < windowEndMs)
      .map(b => ({ von: b.s, bis: b.e }));

    return {
      code: apt.code,
      name: apt.name,
      betten: apt.maxPersonen,
      zimmer: apt.schlafzimmer,
      dauerhaftBelegt: apt.dauerhaftBelegt,
      bestehendeBuchungen: relevanteBuchungen // Leer = komplett frei im Zeitfenster
    };
  });
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