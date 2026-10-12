// =================================================================
// MAIN SCRIPT - HAUPTSTEUERUNG DER PIPELINE
// =================================================================
function mainPipeline() {
  Logger.log("==================================================");
  Logger.log(">>> START DER GESAMTEN AUTOMATISIERUNG PIPELINE <<<");
  Logger.log("==================================================");

  // SCHRITT 1: Duplikatsprüfung (Setzt "Dopplung" in Spalte S)
  try {
    Logger.log("\n--- [1/4] Duplikatsprüfung ---");
    checkDuplicateRequests();
  } catch (e) {
    Logger.log("FEHLER in Schritt 1 (Duplikate): " + e.toString());
  }

  // SCHRITT 2: Gemini Datenkorrektur (Prüft Nachricht & korrigiert C, D, E, F, L, loggt in R)
  try {
    Logger.log("\n--- [2/4] Gemini Datenkorrektur ---");
    processApartmentRequestsWithGemini();
  } catch (e) {
    Logger.log("FEHLER in Schritt 2 (Gemini Korrektur): " + e.toString());
  }

  // SCHRITT 3: Verfügbarkeiten prüfen (DEIN UNVERÄNDERTES ORIGINAL-SKRIPT -> Schreibt in Spalte O)
  try {
    Logger.log("\n--- [3/4] Verfügbare Wohnungen suchen (Original) ---");
    findAvailableApartments();
  } catch (e) {
    Logger.log("FEHLER in Schritt 3 (Wohnungssuche): " + e.toString());
  }

  // SCHRITT 4: Priorisierung der verfügbaren Links (Schreibt in Spalte P)
  try {
    Logger.log("\n--- [4/4] Priorisierung der Links ---");
    prioritizeAvailableApartments();
  } catch (e) {
    Logger.log("FEHLER in Schritt 4 (Priorisierung): " + e.toString());
  }

  Logger.log("\n==================================================");
  Logger.log(">>> PIPELINE ERFOLGREICH VOLLSTÄNDIG DURCHGELAUFEN <<<");
  Logger.log("==================================================");
}