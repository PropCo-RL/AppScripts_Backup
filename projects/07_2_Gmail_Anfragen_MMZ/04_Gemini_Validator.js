// =================================================================
// SKRIPT: VALIDATOR & GESCHÄFTSLOGIK FÜR TAB "MMZ"
// Schreibt alle Ergebnisse in EINEM Gesamt-JSON nur in Spalte D
// =================================================================

function runBusinessValidatorMMZ() {
  const SHEET_NAME = 'MMZ';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`FEHLER: Sheet "${SHEET_NAME}" nicht gefunden.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Header für Spalte D setzen
  ensureValidatorHeaderMMZ(sheet);

  // Spalten A bis D einlesen (4 Spalten: Index 0=A, 1=B, 2=C, 3=D)
  const numRows = lastRow - 1;
  const dataRange = sheet.getRange(2, 1, numRows, 4);
  const rows = dataRange.getValues();
  let updatesMade = false;

  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    const cellB = String(rows[i][1] || "").trim(); // Spalte B: Bereinigtes Unique-JSON
    const cellC = String(rows[i][2] || "").trim(); // Spalte C: Gemini-Roh-JSON
    const cellD = String(rows[i][3] || "").trim(); // Spalte D: Gesamt-JSON-Ausgabe

    // Überspringen wenn: kein Gemini-Puffer da ist, ein Fehler vorliegt oder Spalte D bereits befüllt ist
    if (!cellC || cellC.startsWith("Fehler") || cellD !== "") {
      continue;
    }

    let leadJson, geminiData;
    try {
      leadJson = JSON.parse(cellB);
      geminiData = JSON.parse(cellC);
    } catch (e) {
      rows[i][3] = JSON.stringify({ fehler: "JSON-Parsing fehlgeschlagen" });
      updatesMade = true;
      continue;
    }

    Logger.log(`Validiere und verarbeite Zeile ${rowNum} (${leadJson.gastName || 'Gast'})...`);

    // 1. ZITAT-AUDIT GEGEN DEN ORIGINALTEXT
    const zitatPruefung = auditZitateGegenFreitext(leadJson, geminiData);

    // 2. EXAKTE TAGEBERECHNUNG (UTC-basiert)
    let tageExakt = 0;
    let startD = null;
    let endeD = null;
    if (geminiData.zeitraumData?.start && geminiData.zeitraumData?.ende) {
      startD = parseISODate(geminiData.zeitraumData.start);
      endeD = parseISODate(geminiData.zeitraumData.ende);
      if (startD && endeD && endeD >= startD) {
        const diffMs = endeD.getTime() - startD.getTime();
        tageExakt = Math.round(diffMs / (1000 * 60 * 60 * 24));
        if (tageExakt === 0) tageExakt = 1;
      }
    }

    // maxTage Absicherung
    const hatVerlaengerung = Boolean(geminiData.zeitraumData?.verlaengerung);
    const geminiMaxTage = Number(geminiData.zeitraumData?.maxTage) || 0;
    let finalMaxTage = Math.max(geminiMaxTage, tageExakt);
    if (hatVerlaengerung && finalMaxTage < 30) {
      finalMaxTage = 30; // Hebt Verlängerungsoptionen mindestens auf Monatsbasis
    }

    const pers = Number(geminiData.personenData?.Pers || leadJson.personen || 1);
    const maxPers = Number(geminiData.personenData?.max || pers);
    const berechnungsPersonen = Math.max(pers, maxPers);

    // 3. DETERMINISTISCHE VIP-BERECHNUNG
    let vip = "-";
    if (berechnungsPersonen >= 3 && finalMaxTage >= 120) {
      vip = "VIP2";
    } else if (berechnungsPersonen >= 3 && finalMaxTage >= 30) {
      vip = "VIP1";
    }

    // 4. DATUMSVORSCHAU "Z" & ORT BAUEN
    const zVorschau = buildZVorschau(startD, endeD, tageExakt);
    const stadtBasis = extractCityFromAddress(leadJson.unterkunftAdresse || leadJson.unterkunftTitel || "");

    // 5. SCORES BERECHNEN
    let kiScore = typeof geminiData.kiConfidenceScore === "number" ? geminiData.kiConfidenceScore : 80;
    let kiStatus = "OK";
    if (!zitatPruefung.passed) {
      kiScore = 0;
      kiStatus = "ALARM: ZITAT_FEHLT";
    } else if (kiScore < 85) {
      kiStatus = "UNSICHER";
    }

    const gastScore = typeof geminiData.gastConfidenceScore === "number" ? geminiData.gastConfidenceScore : 100;
    const gastStatus = geminiData.gastIstVage ? `Vage (${geminiData.vageGruende || 'Unpräzise'})` : "Verbindlich";

    // =============================================================
    // EINZIGES GESAMT-JSON-OBJEKT FÜR SPALTE D
    // =============================================================
    const gesamtErgebnis = {
      vip: vip,
      ort: {
        Ort: stadtBasis,
        suchortText: geminiData.suchortText || ""
      },
      personen: {
        Pers: pers,
        max: maxPers,
        aufteilung: geminiData.personenData?.aufteilung || "",
        zimmer: geminiData.personenData?.zimmer || ""
      },
      dauer: {
        Z: zVorschau,
        tage: tageExakt,
        maxTage: finalMaxTage,
        abrechnung: geminiData.zeitraumData?.abrechnung || null,
        verlaengerung: hatVerlaengerung,
        muster: geminiData.zeitraumData?.muster || "durchgehend",
        start: geminiData.zeitraumData?.start || "",
        ende: geminiData.zeitraumData?.ende || "",
        info: geminiData.zeitraumData?.info || ""
      },
      spezial: geminiData.spezialData || {},
      confidence: {
        ki: `${kiScore}% | ${kiStatus}`,
        gast: `${gastScore}% | ${gastStatus}`
      },
      details: {
        zitate: geminiData.zitate || {},
        zitatFehler: zitatPruefung.issues,
        istStaffelAnfrage: Boolean(geminiData.istStaffelAnfrage),
        zeitraumSegmente: geminiData.zeitraumSegmente || []
      }
    };

    // Ausschließlich in Spalte D schreiben (Index 3)
    rows[i][3] = JSON.stringify(gesamtErgebnis);
    updatesMade = true;
  }

  // Zurückschreiben und Zeilenhöhe fixieren
  if (updatesMade) {
    dataRange.setValues(rows);
    for (let r = 2; r <= lastRow; r++) {
      sheet.setRowHeight(r, 21);
    }
    Logger.log("Business Validator für MMZ erfolgreich abgeschlossen.");
  }
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN
// -----------------------------------------------------------------

function auditZitateGegenFreitext(leadJson, geminiData) {
  const issues = [];
  const rawMsg = String(leadJson.nachricht || "");
  if (!rawMsg || rawMsg.trim().length < 5 || rawMsg.toLowerCase() === "keine nachricht") {
    return { passed: true, issues: [] };
  }

  const normalize = (str) => {
    return String(str || "")
      .toLowerCase()
      .replace(/[\r\n\t]+/g, " ")
      .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?"'„“•]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  };

  const cleanRaw = normalize(rawMsg);
  const zitate = geminiData.zitate || {};
  const persZitat = normalize(zitate.personenFundstelle);
  const zeitZitat = normalize(zitate.zeitraumFundstelle);

  if (persZitat && !persZitat.includes("nur im formular") && persZitat.length > 3) {
    if (!cleanRaw.includes(persZitat)) {
      issues.push(`Personen-Zitat nicht im Freitext: "${zitate.personenFundstelle}"`);
    }
  }

  if (zeitZitat && !zeitZitat.includes("nur im formular") && zeitZitat.length > 3) {
    if (!cleanRaw.includes(zeitZitat)) {
      issues.push(`Zeitraum-Zitat nicht im Freitext: "${zitate.zeitraumFundstelle}"`);
    }
  }

  return {
    passed: issues.length === 0,
    issues: issues
  };
}

function parseISODate(dateStr) {
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return null;
  const [y, m, d] = dateStr.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}

function buildZVorschau(startD, endeD, tage) {
  if (!startD || !endeD || isNaN(startD.getTime()) || isNaN(endeD.getTime())) return "";
  const monate = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
  const d1 = String(startD.getUTCDate()).padStart(2, '0');
  const m1 = monate[startD.getUTCMonth()];
  const d2 = String(endeD.getUTCDate()).padStart(2, '0');
  const m2 = monate[endeD.getUTCMonth()];
  const yy = String(endeD.getUTCFullYear()).slice(-2);
  return `${d1} ${m1} - ${d2} ${m2} ${yy} (${tage}T)`;
}

function extractCityFromAddress(addressStr) {
  if (!addressStr || addressStr === "N/A") return "";
  const match = addressStr.match(/\b\d{5}\s+([A-Za-zÄÖÜäöüß\-.\s]+)$/);
  if (match && match[1]) return match[1].trim();
  const parts = addressStr.split(",");
  return parts.length > 1 ? parts[parts.length - 1].replace(/\b\d{5}\b/, "").trim() : addressStr.trim();
}

function ensureValidatorHeaderMMZ(sheet) {
  const cell = sheet.getRange(1, 4); // Spalte D
  if (!cell.getValue()) {
    cell.setValue("Validiert (Gesamt JSON)");
    cell.setFontWeight("bold");
    sheet.getRange("D:D").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}