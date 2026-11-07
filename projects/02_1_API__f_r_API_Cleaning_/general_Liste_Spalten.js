/**
 * Erstellt eine Übersicht aller Tabs, deren Spalten und Formeln/Beispielwerte.
 * Zusätzlich wird am Ende für jeden Tab eine Zusammenfassung als Fließtext in einer Zelle erstellt.
 */
function generateColumnOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = "Liste Spalten";
  let targetSheet = ss.getSheetByName(targetSheetName);

  if (targetSheet) {
    targetSheet.clear();
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }

  const results = [["Tab", "Spalte", "Überschrift", "Formel / Beispielwert"]];
  const formats = [["@", "@", "@", "@"]]; 
  const summaryTexts = []; // Speicher für den Fließtext pro Tab

  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === targetSheetName) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) return;

    // 1. Header finden (Suche in den ersten 10 Zeilen)
    let headerRowIndex = 1;
    let maxColsFilled = 0;
    const searchRange = sheet.getRange(1, 1, Math.min(10, lastRow), lastCol).getValues();

    for (let r = 0; r < searchRange.length; r++) {
      let filledCount = searchRange[r].filter(cell => cell !== "").length;
      if (filledCount > maxColsFilled) {
        maxColsFilled = filledCount;
        headerRowIndex = r + 1;
      }
    }

    const headers = sheet.getRange(headerRowIndex, 1, 1, lastCol).getValues()[0];
    const dataRowIndex = headerRowIndex + 1;
    let formulaRange = [];
    let valueRange = [];
    let formatRange = [];

    if (lastRow >= dataRowIndex) {
      const targetRange = sheet.getRange(dataRowIndex, 1, 1, lastCol);
      formulaRange = targetRange.getFormulas()[0];
      valueRange = targetRange.getValues()[0];
      formatRange = targetRange.getNumberFormats()[0];
    }

    let tabSummary = "Tab '" + sheetName + "': ";
    let colDetails = [];

    for (let c = 0; c < lastCol; c++) {
      let colLetter = columnToLetter(c + 1);
      let headerName = headers[c] || "(Leer)";
      let displayValue = "";
      let cellFormat = "@"; 

      if (formulaRange[c] !== "") {
        displayValue = "'" + formulaRange[c];
      } else if (valueRange[c] !== "") {
        displayValue = valueRange[c];
        cellFormat = formatRange[c] || "@"; 
      } else {
        displayValue = "-";
      }

      results.push([sheetName, colLetter, headerName, displayValue]);
      formats.push(["@", "@", "@", cellFormat]);

      // Fließtext-Komponente für diese Spalte bauen
      colDetails.push("Spalte " + colLetter + " (" + headerName + ") hat die Funktion/Wert: " + displayValue);
    }
    
    // Den gesamten Text für diesen Tab zusammenfügen
    summaryTexts.push([tabSummary + colDetails.join("; ")]);

    results.push(["", "", "", ""]);
    formats.push(["@", "@", "@", "@"]);
  });

  // 2. Tabelle schreiben
  const range = targetSheet.getRange(1, 1, results.length, 4);
  targetSheet.setColumnWidth(1, 60);  
  targetSheet.setColumnWidth(2, 60);  
  targetSheet.setColumnWidth(3, 150); 
  targetSheet.setColumnWidth(4, 400); 

  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  range.setHorizontalAlignment("left");
  range.setVerticalAlignment("middle");
  range.setNumberFormats(formats);
  range.setValues(results);
  
  targetSheet.setRowHeights(1, results.length, 21);
  targetSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f3f3");

  // 3. Fließtext-Ergebnisse GANZ AM ENDE anfügen
  if (summaryTexts.length > 0) {
    const startRowSummary = results.length + 2;
    // Header für die Zusammenfassung
    targetSheet.getRange(startRowSummary - 1, 1).setValue("Zusammenfassung als Fließtext:").setFontWeight("bold");
    
    // Die Zusammenfassungen untereinander schreiben (jeder Tab eine Zelle)
    const summaryRange = targetSheet.getRange(startRowSummary, 1, summaryTexts.length, 1);
    summaryRange.setValues(summaryTexts);
    
    // Optik für die Textzellen: Zeilenumbruch erlauben, damit man alles lesen kann
    summaryRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    targetSheet.setColumnWidth(1, 600); // Spalte A breiter machen für den Text
    // Hier keine feste Zeilenhöhe, damit der Fließtext lesbar bleibt
  }

  targetSheet.setFrozenRows(1);
}

function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}