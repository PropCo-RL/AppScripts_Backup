/**
 * Master Controller: Führt die drei spezifischen Skripte nacheinander aus.
 * Ohne LockService und nur mit den definierten Funktionen.
 */
function runLodgifySyncCycle() {
  console.log("--- START: Lodgify Sync Cycle ---");

  try {
    // 1. Verfügbarkeiten (Blockierungen) abrufen -> schreibt in 'dbX'
    console.log("Schritt 1: Rufe Lodgify Verfügbarkeiten ab...");
    getLodgifyAvailability();
    
    // 2. Buchungen abrufen -> schreibt in 'db'
    console.log("Schritt 2: Rufe Lodgify Buchungen ab...");
    getLodgifyBookings();
    
    // 3. Daten aufbereiten und exportieren -> prüft 'ava' & 'wCl' und schreibt in externes Sheet
    console.log("Schritt 3: Starte Export und Formatierung...");
    exportCleaningWithFormat();

    console.log("--- ERFOLG: Alle Prozesse nacheinander ausgeführt. ---");
  } catch (e) {
    console.error("Fehler im Master-Prozess: " + e.toString());
  }
}