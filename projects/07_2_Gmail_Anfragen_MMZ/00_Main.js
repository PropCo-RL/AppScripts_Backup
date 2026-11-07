// =================================================================
// 0_MAIN PIPELINE ORCHESTRATOR FÜR MMZ (HIGH-PERFORMANCE & DETAIL-LOG)
// =================================================================
function runFullMMZPipeline() {
  const pipelineStartTime = new Date().getTime();
  Logger.log("==================================================");
  Logger.log("🚀 START VOLLSTÄNDIGER MMZ-PIPELINE-DURCHLAUF");
  Logger.log("==================================================");

  // 1.0 E-Mails auslesen, weiterleiten, archivieren -> Spalte A
  executeStep("1.0_JSON_MMZ (Mail-Import)", function() {
    if (typeof verarbeiteMMZEmails === "function") {
      verarbeiteMMZEmails();
    } else {
      Logger.log("⚠️ Funktion verarbeiteMMZEmails() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 1.1 Selbstheilung: Fehlende Mails nachträglich importieren -> Spalte A
  // (Ohne alten Fallback auf auditMMZPipeline, um Doppel-Logs zu verhindern)
  executeStep("1.1_Selbstheilung (Nachträglicher Mail-Import)", function() {
    if (typeof recoverMissingMMZEmails === "function") {
      recoverMissingMMZEmails();
    } else {
      Logger.log("⚠️ Funktion recoverMissingMMZEmails() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 2. Duplikate filtern -> Spalte B
  executeStep("2_Duplikate (Unique Filter)", function() {
    if (typeof verarbeiteUndBereinigeMMZ === "function") {
      verarbeiteUndBereinigeMMZ();
    } else {
      Logger.log("⚠️ Funktion verarbeiteUndBereinigeMMZ() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 3. Gemini KI-Extraktion -> Spalte C
  executeStep("3_Gemini (KI Extraktion)", function() {
    if (typeof runGeminiExtractorMMZ === "function") {
      runGeminiExtractorMMZ();
    } else if (typeof extrahiereMitGeminiMMZ === "function") {
      extrahiereMitGeminiMMZ();
    } else if (typeof runGeminiMMZ === "function") {
      runGeminiMMZ();
    } else if (typeof processGeminiRows === "function") {
      processGeminiRows();
    } else {
      Logger.log("❌ FEHLER: Keine Gemini-Extraktor-Funktion gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 4. Business Validator & Scoring -> Spalte D
  executeStep("4_Gemini_Validator (Validierung)", function() {
    if (typeof runBusinessValidatorMMZ === "function") {
      runBusinessValidatorMMZ();
    } else if (typeof validateMMZRows === "function") {
      validateMMZRows();
    } else {
      Logger.log("⚠️ Validator-Funktion prüfen und ggf. Namen anpassen.");
    }
  });
  SpreadsheetApp.flush();

  // 5. K.O.- & Routing-Vorfilter -> Spalte E
  executeStep("5_Pre_Filter (Routing Vorfilter)", function() {
    if (typeof runLeadPreFilterMMZ === "function") {
      runLeadPreFilterMMZ();
    } else {
      Logger.log("⚠️ Funktion runLeadPreFilterMMZ() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 6.0 Engine A: Einzelapartments (<= 6P) Wunschtermin -> Spalte F
  executeStep("6.0_Bis 6 Personen (Engine A)", function() {
    if (typeof runApartmentMatchingEngineA_MMZ === "function") {
      runApartmentMatchingEngineA_MMZ();
    } else {
      Logger.log("⚠️ Funktion runApartmentMatchingEngineA_MMZ() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 6.1 Engine A-Alt: Alternative Einzelapartments (nächster freier Slot) -> Spalte I
  executeStep("6.1_Bis 6 Personen Alternativen (Engine A-Alt)", function() {
    if (typeof runApartmentMatchingEngineA_Alternative_MMZ === "function") {
      runApartmentMatchingEngineA_Alternative_MMZ();
    } else {
      Logger.log("⚠️ Funktion runApartmentMatchingEngineA_Alternative_MMZ() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 7.0 Engine B: Kombi-Apartments (> 6P) Wunschtermin -> Spalte G
  executeStep("7.0_Ab 7 Personen (Engine B)", function() {
    if (typeof runApartmentMatchingEngineB_MMZ === "function") {
      runApartmentMatchingEngineB_MMZ();
    } else {
      Logger.log("⚠️ Funktion runApartmentMatchingEngineB_MMZ() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 7.1 Engine B-Alt: Alternative Kombis (nächster freier Slot) -> Spalte J
  executeStep("7.1_Ab 7 Personen Alternativen (Engine B-Alt)", function() {
    if (typeof runApartmentMatchingEngineB_Alternative_MMZ === "function") {
      runApartmentMatchingEngineB_Alternative_MMZ();
    } else {
      Logger.log("⚠️ Funktion runApartmentMatchingEngineB_Alternative_MMZ() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 8. Gmail Draft Generator -> Spalte H & Gmail Entwürfe
  executeStep("8_Gmail (Draft Generator)", function() {
    if (typeof generateEmailDraftsMMZ === "function") {
      generateEmailDraftsMMZ();
    } else {
      Logger.log("⚠️ Funktion generateEmailDraftsMMZ() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  // 10. Audit & Hänger-Prüfung -> Schreibt erst JETZT ins Tab "Kontrolle"
  executeStep("10_Audit Kontrolle (Tab Kontrolle)", function() {
    if (typeof logMMZPipelineAudit === "function") {
      logMMZPipelineAudit();
    } else {
      Logger.log("⚠️ Funktion logMMZPipelineAudit() nicht gefunden.");
    }
  });
  SpreadsheetApp.flush();

  const totalDurationSec = ((new Date().getTime() - pipelineStartTime) / 1000).toFixed(2);
  Logger.log("==================================================");
  Logger.log(`🏁 PIPELINE ERFOLGREICH BEENDET (Gesamtdauer: ${totalDurationSec}s)`);
  Logger.log("==================================================");
}

// -----------------------------------------------------------------
// HIGH-SPEED EXECUTION WRAPPER MIT DETAIL-TIMING
// -----------------------------------------------------------------
function executeStep(stepName, fn) {
  const stepStartTime = new Date().getTime();
  Logger.log(`--- [START] ${stepName} ---`);
  try {
    fn();
    const stepDurationMs = new Date().getTime() - stepStartTime;
    const durationLabel = stepDurationMs > 1000
      ? `${(stepDurationMs / 1000).toFixed(2)}s`
      : `${stepDurationMs}ms`;
    Logger.log(`✅ [FERTIG] ${stepName} (${durationLabel})`);
  } catch (error) {
    Logger.log(`❌ [FEHLER] in ${stepName}: ${error.toString()}`);
    if (error.stack) {
      Logger.log(`Stacktrace: ${error.stack}`);
    }
  }
}