function fetchInputDataDataOnly() {
  var sourceSpreadsheetId = "1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU";
  var sourceSs = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSs.getSheetByName("RE");
  if (!sourceSheet) return;
  
  var sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length <= 1) return;
  
  // Heutiges Datum auf 00:00:00 Uhr setzen
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Gewünschte Quellspalten-Indizes (0-basiert):
  var targetIndices = [0, 1, 8, 10, 11, 13, 18, 19, 21, 27, 28, 31, 36];
  var rowsToInsert = [];
  
  // Zeilen filtern (ab Zeile 2 / Index 1 im Quellsheet)
  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    var anreiseDatumVal = row[10]; // Spalte K (Index 10)
    
    // Robuste Datumskonvertierung für Anreise und Abreise
    var anreiseDatum = parseDateSafe(anreiseDatumVal);
    var abreiseDatum = parseDateSafe(row[11]); // Spalte L (Index 11)
    
    if (anreiseDatum && !isNaN(anreiseDatum.getTime())) {
      // Temporäres Date-Objekt für den Vergleicher (ohne Uhrzeit)
      var compareDate = new Date(anreiseDatum.getTime());
      compareDate.setHours(0, 0, 0, 0);
      
      // Nur Daten ab heute (inklusive heute) übernehmen
      if (compareDate >= today) {
        var extractedRow = [];
        for (var j = 0; j < targetIndices.length; j++) {
          var val = row[targetIndices[j]];
          
          // Datumsformate beim Einfügen direkt vereinheitlichen (DD.MM.YYYY)
          if (targetIndices[j] === 10 && anreiseDatum) {
            val = formatDateDE(anreiseDatum);
          } else if (targetIndices[j] === 11 && abreiseDatum) {
            val = formatDateDE(abreiseDatum);
          }
          
          extractedRow.push(val);
        }
        
        // Füge das echte Anreise-Date-Objekt als Hilfseigenschaft für die Sortierung an
        rowsToInsert.push({
          dateObj: compareDate,
          rowArray: extractedRow
        });
      }
    }
  }
  
  // Sortierung: Nächstes Anreisedatum ganz oben (aufsteigend)
  rowsToInsert.sort(function(a, b) {
    return a.dateObj - b.dateObj;
  });
  
  // Array wieder in flache Zeilen-Struktur für setValues umwandeln
  var finalData = rowsToInsert.map(function(item) {
    return item.rowArray;
  });
  
  // Ziel-Tabellenblatt verarbeiten
  var targetSs = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = targetSs.getSheetByName("Input");
  if (!targetSheet) {
    targetSheet = targetSs.insertSheet("Input");
  } else {
    // Löscht nur den Inhalt AB Zeile 2
    var lastRow = targetSheet.getLastRow();
    if (lastRow > 1) {
      targetSheet.getRange(2, 1, lastRow - 1, targetSheet.getLastColumn()).clearContent();
    }
  }
  
  // Neue Daten unter den Header schreiben
  if (finalData.length > 0) {
    targetSheet.getRange(2, 1, finalData.length, finalData[0].length).setValues(finalData);
  }
}

/**
 * Hilfsfunktion zur sicheren Konvertierung verschiedener Datumsformate in ein Date-Objekt
 */
function parseDateSafe(value) {
  if (!value) return null;
  
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return value;
  }
  
  if (typeof value === "string") {
    value = value.trim();
    
    // Format TT.MM.JJJJ (z.B. "12.08.2026")
    var deMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (deMatch) {
      return new Date(parseInt(deMatch[3], 10), parseInt(deMatch[2], 10) - 1, parseInt(deMatch[1], 10));
    }
    
    // Format MM/DD/YYYY oder M/D/YYYY (z.B. "8/12/2026")
    var usMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (usMatch) {
      return new Date(parseInt(usMatch[3], 10), parseInt(usMatch[1], 10) - 1, parseInt(usMatch[2], 10));
    }
    
    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }
  
  return null;
}

/**
 * Hilfsfunktion zur Formatierung eines Date-Objekts in DD.MM.YYYY
 */
function formatDateDE(date) {
  var day = ("0" + date.getDate()).slice(-2);
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var year = date.getFullYear();
  return day + "." + month + "." + year;
}