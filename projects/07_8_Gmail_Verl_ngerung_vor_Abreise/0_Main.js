/**
 * SCRIPT 0: Haupt-Workflow / Orchestrator
 * Führt alle Teil-Prozesse (Skripte 1 bis 5) nacheinander in der richtigen Reihenfolge aus.
 * 
 * AUSLÖSER (Trigger): Richte einen zeitgesteuerten Auslöser für diese Funktion ein 
 * (z. B. Alle 30 Minuten oder Jede Stunde).
 */
function mainWorkflow() {
  Logger.log("==================================================");
  Logger.log("START: AUTOMATISCHER ABREISE-WORKFLOW");
  Logger.log("==================================================");

  try {
    // 1. Abreisen aus Quell-Sheet laden, sortieren & nach 'input' schreiben
    Logger.log("--- SCHRITT 1: updateInputSheet() gestartet ---");
    updateInputSheet();

    // 2. Neue/Aktuelle Einträge von 'input' nach 'output' übertragen & veraltete löschen
    Logger.log("--- SCHRITT 2: transferInputToOutput() gestartet ---");
    transferInputToOutput();

    // 4. In Gmail nach eingehenden Kundenantworten suchen
    Logger.log("--- SCHRITT 4: checkCustomerReplies() gestartet ---");
    checkCustomerReplies();

    // 5. Neue Kundenantworten mit der Gemini API analysieren
    Logger.log("--- SCHRITT 5: processRepliesWithGemini() gestartet ---");
    processRepliesWithGemini();

    Logger.log("==================================================");
    Logger.log("ENDE: WORKFLOW ERFOLGREICH ABGESCHLOSSEN");
    Logger.log("==================================================");

  } catch (e) {
    Logger.log("FEHLER IM HAUPT-WORKFLOW: " + e.toString());
  }
}