function main() {
  Logger.log("=== STARTE GESAMTEN CHECK-IN PROZESS ===");

  // 1. Quelldaten aus dem zentralen Sheet abrufen
  Logger.log("Schritt 1: Hole Quelldaten aus Tab 'd'...");
  holeDatenAusQuelldatensatz();
  SpreadsheetApp.flush();

  // 2. Input-Daten filtern & aufbereiten
  Logger.log("Schritt 2: Verarbeite Input-Daten...");
  fetchInputDataDataOnly();
  SpreadsheetApp.flush();

  // 3. Input in Output übertragen, säubern & sortieren
  Logger.log("Schritt 3: Übertrage Input zu Output...");
  processInputToOutput();
  SpreadsheetApp.flush();

  // 4. E-Mail-Posteingang nach neuen Zahlungsbelegen scannen
  Logger.log("Schritt 4: Scanne Gmail nach Zahlungsnachweisen...");
  scanGmailForPaymentProofs();
  SpreadsheetApp.flush();

  // 5. WhatsApp-Webhooks via 360dialog nach neuen Zahlungsbelegen scannen
  Logger.log("Schritt 5: Scanne WhatsApp nach Zahlungsnachweisen...");
  scanWhatsAppForPaymentProofs();
  SpreadsheetApp.flush();

  // 6. Gemini KI-Prüfung für alle gesammelten Belege (E-Mail & WhatsApp) ausführen
  Logger.log("Schritt 6: Starte Gemini KI-Prüfung für Belege...");
  processPaymentProofsWithGemini();
  SpreadsheetApp.flush();

  // 7. E-Mail-Entwürfe / Check-in Mails in Gmail erstellen (bei Status bezahlt oder Gemini ERFOLGREICH)
  Logger.log("Schritt 7: Erstelle E-Mail-Entwürfe in Gmail...");
  processEmailDrafts();
  SpreadsheetApp.flush();

  Logger.log("=== GESAMTPROZESS ERFOLGREICH BEENDET ===");
}