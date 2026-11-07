function exportCleaningWithFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("wCl");
  const avaSheet = ss.getSheetByName("ava"); // Das AVA-Sheet definieren
  
  const targetFileId = "1RiMepCwbPsVGD70m_tEeDTjEsSkBTdba003cc2vFGCs"; 
  const targetSheetName = "Import";

  try {
    // --- 1. C5 FEHLER-CHECK (Bestehend) ---
    const checkValue = sourceSheet.getRange("C5").getValue();
    if (checkValue instanceof Error || checkValue === "#N/A" || checkValue === "#REF!" || (typeof checkValue === "string" && checkValue.includes("#N/A"))) {
      console.log("Abbruch: Zelle C5 enthält einen Fehler.");
      return;
    }

    // Ab hier folgt der normale Export...
    const targetSS = SpreadsheetApp.openById(targetFileId);
    let targetSheet = targetSS.getSheetByName(targetSheetName);
    if (!targetSheet) targetSheet = targetSS.insertSheet(targetSheetName);

    // 2. HEADER (A4:I4)
    const headerRange = sourceSheet.getRange(4, 1, 1, 9);
    const hValues = headerRange.getValues();
    const hBacks = headerRange.getBackgrounds();
    const hWeights = headerRange.getFontWeights();
    const hColors = headerRange.getFontColors();
    const hAligns = headerRange.getHorizontalAlignments();
    const hFormats = headerRange.getNumberFormats();

    // 3. DATEN (A5:I146)
    const sourceRange = sourceSheet.getRange(5, 1, 142, 9);
    const dValues = sourceRange.getValues();
    const dBacks = sourceRange.getBackgrounds();
    const dWeights = sourceRange.getFontWeights();
    const dColors = sourceRange.getFontColors();
    const dAligns = sourceRange.getHorizontalAlignments();
    const dFormats = sourceRange.getNumberFormats();

    // 4. FILTER & DUPLIKAT-PRÜFUNG + ZEITFILTER (Max 7 Tage Zukunft)
    let fValues = [], fBacks = [], fWeights = [], fColors = [], fAligns = [], fFormats = [];
    let seenKeys = new Set(); 

    const today = new Date();
    today.setHours(0, 0, 0, 0); // Zeit auf Mitternacht setzen für sauberen Vergleich

    const maxFutureDate = new Date();
    maxFutureDate.setDate(today.getDate() + 8);
    maxFutureDate.setHours(23, 59, 59, 999); // Bis zum Ende des 7. Tages

    for (let i = 0; i < dValues.length; i++) {
      const colA = dValues[i][0]; // Datum (Cleaning Datum)
      const colB = dValues[i][1]; // Name
      
      // 1. Check: Ist die Zeile leer?
      if (colA !== "" && colA !== "-" && colA !== null && colA instanceof Date) {
        
        // 2. Check: Liegt das Datum im Bereich (Heute bis heute + 7 Tage)?
        // Wenn das Datum kleiner als heute ODER größer als maxFutureDate ist -> Überspringen
        if (colA < today || colA > maxFutureDate) {
          continue; 
        }

        // 3. Duplikat-Prüfung
        const rowKey = colA.getTime() + "|" + colB;

        if (!seenKeys.has(rowKey)) {
          fValues.push(dValues[i]);
          fBacks.push(dBacks[i]);
          fWeights.push(dWeights[i]);
          fColors.push(dColors[i]);
          fAligns.push(dAligns[i]);
          fFormats.push(dFormats[i]);
          
          seenKeys.add(rowKey);
        }
      }
    }

    // 5. CLEAN
    targetSheet.clearContents();
    targetSheet.clearFormats();

    // 6. HEADER schreiben
    const targetHeaderRange = targetSheet.getRange(1, 1, 1, 9);
    targetHeaderRange.setValues(hValues);
    targetHeaderRange.setBackgrounds(hBacks);
    targetHeaderRange.setFontWeights(hWeights);
    targetHeaderRange.setFontColors(hColors);
    targetHeaderRange.setHorizontalAlignments(hAligns);
    targetHeaderRange.setNumberFormats(hFormats);

    // 7. DATEN schreiben
    const targetDataRange = targetSheet.getRange(2, 1, fValues.length, 9);
    targetDataRange.setValues(fValues);
    targetDataRange.setFontWeights(fWeights);
    targetDataRange.setFontColors(fColors);
    targetDataRange.setHorizontalAlignments(fAligns);
    targetDataRange.setNumberFormats(fFormats);

    // 8. SORTIEREN
    targetDataRange.sort([
    {column: 1, ascending: true}, // 1. Sortierung: Datum (Spalte A)
    {column: 2, ascending: true}  // 2. Sortierung: Name (Spalte B)
    ]);

    // 9. FIXE BREITE
    for (let col = 1; col <= 9; col++) {
      targetSheet.setColumnWidth(col, 50);
    }

    // --- 10. TIMESTAMP IN L1 ---
    const now = new Date();
    // Format: 15. Apr (10:53 AM)
    const formattedDate = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "dd. MMM (hh:mm a)");
    targetSheet.getRange("L1").setValue(formattedDate);

    console.log("Export sauber abgeschlossen: " + fValues.length + " Zeilen.");

  } catch (e) {
    console.error("Fehler: " + e.message);
  }
}