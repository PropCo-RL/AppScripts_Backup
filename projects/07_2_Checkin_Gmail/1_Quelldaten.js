function holeDatenAusQuelldatensatz() {
  // 1. Quellsheet definieren (01.1_WA API Buchung) und Tab 'd' abrufen
  const sourceSpreadsheetId = "1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU";
  const sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
  const sourceSheet = sourceSS.getSheetByName("d");

  if (!sourceSheet) {
    Logger.log("Fehler: Das Tab 'd' wurde im Quellsheet nicht gefunden!");
    return;
  }

  // 2. Aktives Zielsheet (07.2_Checkin Gmail) direkt abrufen
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = activeSS.getSheetByName("d");

  // Falls das Tab 'd' lokal noch nicht existiert, wird es angelegt
  if (!targetSheet) {
    targetSheet = activeSS.insertSheet("d");
  }

  // 3. Datenbereich im Quellsheet ermitteln
  const sourceRange = sourceSheet.getDataRange();
  const sourceValues = sourceRange.getValues();

  if (sourceValues.length === 0 || (sourceValues.length === 1 && sourceValues[0][0] === "")) {
    Logger.log("Hinweis: Keine Daten im Quellsheet 'd' gefunden.");
    return;
  }

  // 4. Ziel-Tab d vorher vollständig leeren
  targetSheet.clear();

  // 5. Daten und Datums-/Zahlenformate aus dem Quellsheet eintragen
  const targetRange = targetSheet.getRange(
    1, 1, 
    sourceRange.getNumRows(),
    sourceRange.getNumColumns()
  );
  
  targetRange.setValues(sourceValues);
  targetRange.setNumberFormats(sourceRange.getNumberFormats());

  Logger.log("Daten wurden erfolgreich aus dem WA API Buchung Sheet geholt!");
}