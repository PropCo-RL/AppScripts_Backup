// =================================================================
// SKRIPT 2: VALIDATOR & GESCHÄFTSLOGIK (Deterministisch & Bereinigt)
// Schreibt Spalten D bis K basierend auf Gemini-Puffer in Spalte C
// =================================================================

function runBusinessValidatorDMZ2() {
  const SHEET_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DMZ2');
  
  if (!sheet) {
    Logger.log("FEHLER: Sheet 'DMZ2' nicht gefunden.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  ensureValidatorHeadersDMZ2(sheet);

  // Spalten A bis K einlesen (11 Spalten: Index 0=A bis 10=K)
  const numRows = lastRow - 1;
  const dataRange = sheet.getRange(2, 1, numRows, 11);
  const rows = dataRange.getValues();
  let updatesMade = false;

  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    const cellB = String(rows[i][1] || "").trim(); // Original Lead
    const cellC = String(rows[i][2] || "").trim(); // Gemini-Roh-JSON Puffer
    const cellD = String(rows[i][3] || "").trim(); // VIP / Status-Marker

    // Nur verarbeiten, wenn Spalte C Daten hat und Spalte D noch leer ist
    if (!cellC || cellC.startsWith("Fehler") || cellD !== "") {
      continue;
    }

    let leadJson, geminiData;
    try {
      leadJson = JSON.parse(cellB);
      geminiData = JSON.parse(cellC);
    } catch (e) {
      rows[i][3] = "Fehler: JSON-Parsing";
      updatesMade = true;
      continue;
    }

    Logger.log(`Validiere und verarbeite Zeile ${rowNum} (${leadJson.gastName || 'Gast'})...`);

    // 1. ZITAT-AUDIT GEGEN DEN ORIGINALTEXT (Mit robuster Normalisierung)
    const zitatPruefung = auditZitateGegenFreitext(leadJson, geminiData);

    // 2. EXAKTE TAGEBERECHNUNG (UTC-basiert, verhindert Sommer-/Winterzeit-Versatz)
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

    // maxTage Absicherung:
    // 1. maxTage darf nie kleiner sein als die fest gebuchten Tage
    // 2. Wenn eine Verlängerung/Projekt genannt wurde, mindestens 30 Tage ansetzen
    const hatVerlaengerung = Boolean(geminiData.zeitraumData?.verlaengerung);
    const geminiMaxTage = Number(geminiData.zeitraumData?.maxTage) || 0;
    
    let finalMaxTage = Math.max(geminiMaxTage, tageExakt);
    if (hatVerlaengerung && finalMaxTage < 30) {
      finalMaxTage = 30; // Hebt Warndt & Co. sauber auf Monats-/Projektstatus
    }

    const pers = Number(geminiData.personenData?.Pers || leadJson.personen || 1);
    const maxPers = Number(geminiData.personenData?.max || pers);
    const berechnungsPersonen = Math.max(pers, maxPers); // Berücksichtigt Aufstockung

    // 3. DETERMINISTISCHE VIP-BERECHNUNG
    let vip = "-";
    if (berechnungsPersonen >= 3 && finalMaxTage >= 120) {
      vip = "VIP2";
    } else if (berechnungsPersonen >= 3 && finalMaxTage >= 30) {
      vip = "VIP1";
    }

    // 4. DATUMSVORSCHAU "Z" BAUEN
    const zVorschau = buildZVorschau(startD, endeD, tageExakt);
    const stadtBasis = extractCityFromAddress(leadJson.unterkunftAdresse);

    // =============================================================
    // SCHREIBEN IN DIE SPALTEN D BIS K
    // =============================================================

    // Spalte D (4): VIP
    rows[i][3] = vip;

    // Spalte E (5): Ort
    rows[i][4] = JSON.stringify({
      Ort: stadtBasis,
      suchortText: geminiData.suchortText || ""
    });

    // Spalte F (6): Personen
    rows[i][5] = JSON.stringify({
      Pers: pers,
      max: maxPers,
      aufteilung: geminiData.personenData?.aufteilung || "",
      zimmer: geminiData.personenData?.zimmer || ""
    });

    // Spalte G (7): Zeitraum & Abrechnung
    rows[i][6] = JSON.stringify({
      Z: zVorschau,
      tage: tageExakt,
      maxTage: finalMaxTage,
      abrechnung: geminiData.zeitraumData?.abrechnung || null,
      verlaengerung: hatVerlaengerung,
      muster: geminiData.zeitraumData?.muster || "durchgehend",
      start: geminiData.zeitraumData?.start || "",
      ende: geminiData.zeitraumData?.ende || "",
      info: geminiData.zeitraumData?.info || ""
    });

    // Spalte H (8): Spezial
    rows[i][7] = JSON.stringify(geminiData.spezialData || {});

    // Spalte I (9): KI-Confidence (Audit-Prüfung)
    let kiScore = typeof geminiData.kiConfidenceScore === "number" ? geminiData.kiConfidenceScore : 80;
    let kiStatus = "OK";
    if (!zitatPruefung.passed) {
      kiScore = 0;
      kiStatus = "ALARM: ZITAT_FEHLT";
    } else if (kiScore < 85) {
      kiStatus = "UNSICHER";
    }
    rows[i][8] = `${kiScore}% | ${kiStatus}`;

    // Spalte J (10): Gast-Confidence (Verbindlichkeit)
    const gastScore = typeof geminiData.gastConfidenceScore === "number" ? geminiData.gastConfidenceScore : 100;
    const gastStatus = geminiData.gastIstVage ? `Vage (${geminiData.vageGruende || 'Unpräzise'})` : "Verbindlich";
    rows[i][9] = `${gastScore}% | ${gastStatus}`;

    // Spalte K (11): Details, Zitate & Segmente
    rows[i][10] = JSON.stringify({
      zitate: geminiData.zitate || {},
      zitatFehler: zitatPruefung.issues,
      istStaffelAnfrage: Boolean(geminiData.istStaffelAnfrage),
      zeitraumSegmente: geminiData.zeitraumSegmente || []
    });

    updatesMade = true;
  }

  if (updatesMade) {
    dataRange.setValues(rows);
    for (let r = 2; r <= lastRow; r++) {
      sheet.setRowHeight(r, 21);
    }
    Logger.log("Business Validator erfolgreich durchgelaufen.");
  }
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN
// -----------------------------------------------------------------
function auditZitateGegenFreitext(leadJson, geminiData) {
  const issues = [];
  const rawMsg = String(leadJson.nachricht || "");
  if (!rawMsg || rawMsg.trim().length < 5) {
    return { passed: true, issues: [] };
  }

  // Vereinheitlicht ALLES: Zeilenumbrüche weg, Satzzeichen weg, Mehrfachleerzeichen weg
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

  // Personen-Zitat prüfen
  if (persZitat && !persZitat.includes("nur im formular") && persZitat.length > 3) {
    if (!cleanRaw.includes(persZitat)) {
      issues.push(`Personen-Zitat nicht im Freitext: "${zitate.personenFundstelle}"`);
    }
  }

  // Zeitraum-Zitat prüfen
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
  return new Date(Date.UTC(y, m - 1, d)); // UTC verhindert Sommerzeit-/Zeitzonenversatz
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

function ensureValidatorHeadersDMZ2(sheet) {
  const headers = [
    [1, "JSON (Roh)"],
    [2, "Bereinigt (Unique)"],
    [3, "Gemini Raw (JSON)"],
    [4, "VIP"],
    [5, "Ort (JSON)"],
    [6, "Pers (JSON)"],
    [7, "Dauer (JSON)"],
    [8, "Spezial (JSON)"],
    [9, "KI-Confidence"],
    [10, "Gast-Confidence"],
    [11, "Details & Zitate (JSON)"]
  ];

  headers.forEach(([col, title]) => {
    const cell = sheet.getRange(1, col);
    if (!cell.getValue()) {
      cell.setValue(title);
      cell.setFontWeight("bold");
    }
  });

  sheet.getRange("A:K").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}