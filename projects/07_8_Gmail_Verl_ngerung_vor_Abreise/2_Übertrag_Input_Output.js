/**
 * SCRIPT 2: Liest die Daten aus Tab 'input' aus und überträgt neue Einträge 
 * nach Tab 'output'. Veraltete Zeilen (falls ein Gast verlängert hat und noch 
 * keine E-Mail gesendet wurde) werden automatisch in 'output' bereinigt.
 */
function transferInputToOutput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('input');
  const outputSheet = ss.getSheetByName('output');

  if (!inputSheet || !outputSheet) {
    Logger.log('Fehler: "input" oder "output" Tab wurde nicht gefunden.');
    return;
  }

  const inputLastRow = inputSheet.getLastRow();
  if (inputLastRow < 3) {
    Logger.log('Keine Daten im Tab "input" vorhanden.');
    return;
  }

  // 1. Daten aus Tab "input" auslesen (14 Spalten)
  const inputData = inputSheet.getRange(3, 1, inputLastRow - 2, 14).getValues();
  
  // Map zur Abfrage der aktuellen Input-Daten nach Buchungsnummer
  const currentInputMap = new Map();
  inputData.forEach(row => {
    const bookingNo = String(row[5]).trim(); // Spalte F: Buchungsnummer
    const uniqueId = String(row[7]).trim();  // Spalte H: Unique ID
    if (bookingNo) {
      currentInputMap.set(bookingNo, {
        uniqueId: uniqueId,
        fullRow: row
      });
    }
  });

  // 2. Vorhandene Daten im Tab "output" auslesen und prüfen
  const outputLastRow = outputSheet.getLastRow();
  const existingOutputIds = new Set();

  if (outputLastRow >= 3) {
    const outputRange = outputSheet.getRange(3, 1, outputLastRow - 2, 14);
    const outputData = outputRange.getValues();
    
    // Rückwärts durch 'output' iterieren, damit das Löschen von Zeilen die Zeilenindizes nicht verschiebt
    for (let i = outputData.length - 1; i >= 0; i--) {
      const row = outputData[i];
      const bookingNo = String(row[5]).trim();       // Spalte F: Buchungsnummer
      const outputUniqueId = String(row[7]).trim();  // Spalte H: Unique ID
      const emailSentStatus = String(row[11]).trim();// Spalte L: Gmail 1 Gesendet
      const rowIndexInSheet = i + 3;

      if (bookingNo) {
        // Bereinigungsprüfung: Wenn noch KEINE E-Mail gesendet wurde, das Datum der Buchung sich 
        // in 'input' aber geändert hat (Verlängerung), wird die veraltete Zeile gelöscht.
        if (!emailSentStatus && currentInputMap.has(bookingNo)) {
          const freshInput = currentInputMap.get(bookingNo);
          if (freshInput.uniqueId !== outputUniqueId) {
            outputSheet.deleteRow(rowIndexInSheet);
            Logger.log(`Zeile ${rowIndexInSheet} gelöscht: Buchung ${bookingNo} hat sich verlängert.`);
            continue;
          }
        }

        existingOutputIds.add(outputUniqueId);
      }
    }
  }

  // 3. Neue Datensätze ermitteln, die noch nicht in 'output' stehen
  const newRowsToAppend = [];
  inputData.forEach(row => {
    const uniqueId = String(row[7]).trim(); // Spalte H: Unique ID
    if (uniqueId && !existingOutputIds.has(uniqueId)) {
      newRowsToAppend.push(row);
      existingOutputIds.add(uniqueId); // Verhindert auch Duplikate innerhalb desselben Durchlaufs
    }
  });

  // 4. Neue Zeilen unten an 'output' anfügen
  if (newRowsToAppend.length > 0) {
    const nextRow = Math.max(outputSheet.getLastRow() + 1, 3);
    outputSheet.getRange(nextRow, 1, newRowsToAppend.length, 14).setValues(newRowsToAppend);
    Logger.log(`${newRowsToAppend.length} neue Zeile(n) erfolgreich nach "output" übertragen.`);
  } else {
    Logger.log('Keine neuen Übertragungen erforderlich.');
  }
}