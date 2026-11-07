function uebertrageLodgifyDaten() {
  // 1. Ursprungssheet definieren und Daten auslesen
  const sourceSpreadsheetId = "18581fxYcLaw-HpSAHctL23gEabi88wl9";
  const sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
  const sourceSheet = sourceSS.getSheets()[0]; // Verwendet das erste Tab; alternativ Namen angeben, z.B. sourceSS.getSheetByName("Sheet1")
  
  const sourceLastRow = sourceSheet.getLastRow();
  if (sourceLastRow < 2) {
    Logger.log("Keine Daten im Ursprungssheet gefunden.");
    return;
  }
  
  // Daten aus Spalte A bis T im Ursprungssheet holen
  const sourceValues = sourceSheet.getRange(2, 1, sourceLastRow - 1, 20).getValues();
  
  // Lookup Map erstellen: Lodgify ID (Spalte T / Index 19) -> { colN: Spalte N (Index 13), colP: Spalte P (Index 15) }
  const dataMap = new Map();
  
  for (let i = 0; i < sourceValues.length; i++) {
    const lodgifyId = String(sourceValues[i][19]).trim(); // Spalte T ist Index 19
    const valN = sourceValues[i][13]; // Spalte N ist Index 13
    const valP = sourceValues[i][15]; // Spalte P ist Index 15
    
    if (lodgifyId !== "") {
      dataMap.set(lodgifyId, { valN: valN, valP: valP });
    }
  }

  // 2. Zielsheet definieren und Tab 'd' abrufen
  const targetSpreadsheetId = "1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU";
  const targetSS = SpreadsheetApp.openById(targetSpreadsheetId);
  const targetSheet = targetSS.getSheetByName("d");
  
  if (!targetSheet) {
    SpreadsheetApp.getUi().alert("Fehler: Das Tab 'd' wurde im Zielsheet nicht gefunden!");
    return;
  }
  
  const targetLastRow = targetSheet.getLastRow();
  if (targetLastRow < 2) {
    Logger.log("Keine Daten im Zielsheet Tab 'd' gefunden.");
    return;
  }
  
  // Lodgify IDs aus Zielsheet Spalte A holen (Zeile 2 bis Ende)
  const targetIds = targetSheet.getRange(2, 1, targetLastRow - 1, 1).getValues();
  
  // Ziel-Arrays für Spalte X und Spalte Y vorbereiten
  const outputX = [];
  const outputY = [];
  
  for (let i = 0; i < targetIds.length; i++) {
    const targetId = String(targetIds[i][0]).trim();
    
    if (dataMap.has(targetId)) {
      const matchedData = dataMap.get(targetId);
      outputX.push([matchedData.valN]);
      outputY.push([matchedData.valP]);
    } else {
      // Wenn keine ID gefunden wurde, leeres Feld beibehalten (oder Vorhandenes nicht überschreiben)
      outputX.push([""]); 
      outputY.push([""]);
    }
  }
  
  // 3. Ergebnisse in Spalte X (24) und Spalte Y (25) im Zielsheet eintragen
  targetSheet.getRange(2, 24, outputX.length, 1).setValues(outputX);
  targetSheet.getRange(2, 25, outputY.length, 1).setValues(outputY);
  
  SpreadsheetApp.getUi().alert("Datenübertragung erfolgreich abgeschlossen!");
}