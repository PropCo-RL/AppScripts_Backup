// =================================================================
// MAIN ORCHESTRATOR: 5-MINUTEN-TRIGGER (Alle Funktionen in Reihenfolge)
// =================================================================

function runCompleteAutomationPipelineDMZ2() {
  const SHEET_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const COL_AI_INDEX = 34; // Spalte AI (0-basiert: 34 = 35. Spalte)

  let sheet;
  try {
    sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DMZ2');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabelle: " + e.toString());
    return;
  }

  // 1. ZUERST NEUE E-MAILS AUS GMAIL EINLESEN
  // (Muss immer laufen, damit neue Leads überhaupt erst in Spalte A ankommen)
  Logger.log("--- 1. GMAIL IMPORT ---");
  if (typeof sendAutoResponseDMZ2 === "function") {
    sendAutoResponseDMZ2();
  }

  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // Noch keine Daten vorhanden

  // 2. DUPLIKATE ERMITTELN & IN GMAIL LABELN
  Logger.log("--- 2.1 DUPLIKATE PRÜFEN ---");
  if (typeof processDMZ2Duplicates === "function") {
    processDMZ2Duplicates();
  }

  Logger.log("--- 2.2 DUPLIKATE LABELN ---");
  if (typeof markiereDuplikateMitLabel === "function") {
    markiereDuplikateMitLabel();
  }

  // 3. BLITZ-CHECK: Gibt es Zeilen mit Spalte B, die noch kein "x" in AI haben?
  const numRows = lastRow - 1;
  const checkData = sheet.getRange(2, 2, numRows, COL_AI_INDEX).getValues();
  let hasPendingWork = false;

  for (let i = 0; i < checkData.length; i++) {
    const cellB = String(checkData[i][0] || "").trim();
    const cellAI = String(checkData[i][COL_AI_INDEX - 1] || "").trim();

    if (cellB !== "" && cellB !== "-" && cellAI !== "x") {
      hasPendingWork = true;
      break;
    }
  }

  // Wenn keine offenen Leads zur weiteren Verarbeitung anstehen -> sofortiger Abbruch
  if (!hasPendingWork) {
    Logger.log("Keine offenen Leads zur Weiterverarbeitung. Beende Durchlauf.");
    return;
  }

  // 4. RESTLICHE PIPELINE AUSFÜHREN
  try {
    Logger.log("--- 3.1 GEMINI EXTRAKTOR ---");
    if (typeof runGeminiExtractorDMZ2 === "function") {
      runGeminiExtractorDMZ2();
    }

    Logger.log("--- 3.2 GEMINI VALIDATOR ---");
    if (typeof runBusinessValidatorDMZ2 === "function") {
      runBusinessValidatorDMZ2();
    }

    Logger.log("--- 4. ROUTING VORFILTER ---");
    if (typeof runLeadPreFilterDMZ2 === "function") {
      runLeadPreFilterDMZ2();
    }

    Logger.log("--- 5.1 ENGINE A (bis 6P) ---");
    if (typeof runApartmentMatchingEngineA === "function") {
      runApartmentMatchingEngineA();
    }

    Logger.log("--- 6.1 ENGINE B (ab 7P) ---");
    if (typeof runApartmentMatchingEngineB === "function") {
      runApartmentMatchingEngineB();
    }

    Logger.log("--- 7. GMAIL ENTWURF ---");
    if (typeof generateEmailDraftsDMZ2 === "function") {
      generateEmailDraftsDMZ2();
    }

    Logger.log("--- 9.9 ENTWURF LABELN ---");
    if (typeof markiereEntwuerfeMitLabel === "function") {
      markiereEntwuerfeMitLabel();
    }

  } catch (err) {
    Logger.log("Fehler während Pipeline: " + err.toString());
  }

  // 5. ABSCHLUSS-MARKER: 'x' IN SPALTE AI SETZEN
  markCompletedRowsInColAI(sheet);

  Logger.log("=== PIPELINE VOLLSTÄNDIG ABGESCHLOSSEN ===");
}

// -----------------------------------------------------------------
// HILFSFUNKTION: Setzt 'x' in Spalte AI
// -----------------------------------------------------------------
function markCompletedRowsInColAI(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const headerAI = sheet.getRange(1, 35);
  if (!headerAI.getValue()) {
    headerAI.setValue("Pipeline Fertig");
    headerAI.setFontWeight("bold");
    sheet.getRange("AI:AI").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Spalten T (Index 19 / Spalte 20) bis AI (Index 34 / Spalte 35) = 16 Spalten
  const numRows = lastRow - 1;
  const range = sheet.getRange(2, 20, numRows, 16);
  const rows = range.getValues();
  let updates = false;

  for (let i = 0; i < rows.length; i++) {
    const mailStatus = String(rows[i][0] || "").trim(); // Spalte T
    const currentAI = String(rows[i][15] || "").trim();  // Spalte AI

    if (mailStatus !== "" && currentAI !== "x") {
      rows[i][15] = "x";
      updates = true;
    }
  }

  if (updates) {
    range.setValues(rows);
    SpreadsheetApp.flush();
    Logger.log("Marker 'x' in Spalte AI eingetragen.");
  }
}