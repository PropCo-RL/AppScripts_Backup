/**
 * Hauptfunktion für die automatisierte Verarbeitung.
 * Kann über einen zeitgesteuerten Trigger alle 5 Minuten ausgeführt werden.
 */
function main() {
  Logger.log("=== START AUTOMATISCHER WORKFLOW ===");

  // 1. Duplikats-Prüfung
  try {
    Logger.log("--> Step 1: Duplikats-Prüfung gestartet");
    checkDuplicateRequests();
  } catch (e) {
    Logger.log("FEHLER in Step 1 (checkDuplicateRequests): " + e.toString());
  }

  // 2. Gemini Daten-Korrektur + Extraktion
  try {
    Logger.log("--> Step 2: Gemini Daten-Korrektur gestartet");
    processApartmentRequestsWithGemini();
  } catch (e) {
    Logger.log("FEHLER in Step 2 (processApartmentRequestsWithGemini): " + e.toString());
  }

  // 3. Vorab-Filterung kleine Gruppen (1P)
  try {
    Logger.log("--> Step 3: Vorab-Filterung (1P) gestartet");
    filterAndProcessSmallGroups();
  } catch (e) {
    Logger.log("FEHLER in Step 3 (filterAndProcessSmallGroups): " + e.toString());
  }

  // 4. Verfügbare Wohnungen suchen (MVP 1)
  try {
    Logger.log("--> Step 4: Angebotssuche (findAvailableApartments) gestartet");
    findAvailableApartments();
  } catch (e) {
    Logger.log("FEHLER in Step 4 (findAvailableApartments): " + e.toString());
  }

  // 5. MVP 2 - Verfeinerte Sortierung & Preisberechnung (Spalte AA)
  try {
    Logger.log("--> Step 5: MVP 2 Sortierung gestartet");
    runMVP2_ApartmentRefinement();
  } catch (e) {
    Logger.log("FEHLER in Step 5 (runMVP2_ApartmentRefinement): " + e.toString());
  }

  // 6. MVP 3 - Sortierung mit SST-Risiko-Check (Spalte AB)
  try {
    Logger.log("--> Step 6: MVP 3 Risiko-Check gestartet");
    runMVP3_ApartmentRefinementWithRisk();
  } catch (e) {
    Logger.log("FEHLER in Step 6 (runMVP3_ApartmentRefinementWithRisk): " + e.toString());
  }

  // 7. MVP 4 - Kombi-Angebote Fallback (Spalte AC)
  try {
    Logger.log("--> Step 7: MVP 4 Kombi-Angebote gestartet");
    runMVP4_ApartmentCombinations();
  } catch (e) {
    Logger.log("FEHLER in Step 7 (runMVP4_ApartmentCombinations): " + e.toString());
  }

  // 8. PDF-Angebote erstellen (Spalte AI)
  try {
    Logger.log("--> Step 8: PDF-Erstellung gestartet");
    createPdfOffersFromColAC(); // <-- HIER GEÄNDERT
  } catch (e) {
    Logger.log("FEHLER in Step 8 (createPdfOffersFromColAC): " + e.toString());
  }

  // 9. E-Mail Entwürfe senden/erstellen
  try {
    Logger.log("--> Step 9: E-Mail Entwürfe senden gestartet");
    sendOfferEmailFromColAC(); // <-- HIER GEÄNDERT
  } catch (e) {
    Logger.log("FEHLER in Step 9 (sendOfferEmailFromColAC): " + e.toString());
  }

  Logger.log("=== WORKFLOW ERFOLGREICH BEENDET ===");
}