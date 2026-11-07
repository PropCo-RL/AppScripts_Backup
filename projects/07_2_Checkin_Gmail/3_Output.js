function processInputToOutput() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("Input");
  var outputSheet = ss.getSheetByName("Output");
  if (!inputSheet || !outputSheet) {
    Logger.log("Fehler: 'Input' oder 'Output' Sheet existiert nicht.");
    return;
  }

  var inputData = inputSheet.getDataRange().getValues();
  var outputData = outputSheet.getDataRange().getValues();
  if (inputData.length <= 1) return; // Keine Daten im Input

  var outputHeaders = outputData[0];
  var inputHeaders = inputData[0];

  // 1. Spalten dynamisch ermitteln
  var buchungColIdx = outputHeaders.indexOf("Buchung Nummer");
  var anreiseColIdx = outputHeaders.indexOf("Anreise");
  if (anreiseColIdx === -1) anreiseColIdx = 3; // Fallback Spalte D

  // Rechnungsstatus-Spalte ermitteln (Spalte L)
  var outputStatusColIdx = outputHeaders.indexOf("Rechnung Status") !== -1 
    ? outputHeaders.indexOf("Rechnung Status") 
    : outputHeaders.indexOf("Rechnungsstatus");
  if (outputStatusColIdx === -1) outputStatusColIdx = 11;

  var inputStatusColIdx = inputHeaders.indexOf("Rechnung Status") !== -1 
    ? inputHeaders.indexOf("Rechnung Status") 
    : inputHeaders.indexOf("Rechnungsstatus");
  if (inputStatusColIdx === -1) inputStatusColIdx = 11;

  if (buchungColIdx === -1) {
    Logger.log("Fehler: Spalte 'Buchung Nummer' wurde im Output-Sheet nicht gefunden.");
    return;
  }

  var totalOutputCols = outputHeaders.length;

  // 2. Bestehende Output-Zeilen vollständig erfassen und indizieren
  // Key = Buchungsnummer, Value = Index im masterRows-Array
  var existingBookingsMap = new Map();
  var masterRows = outputData.slice(1);

  for (var i = 0; i < masterRows.length; i++) {
    var bNummer = String(masterRows[i][buchungColIdx] || "").trim();
    if (bNummer !== "") {
      // Fehlende Spalten auffüllen, falls Header breiter als die Datenzeile ist
      while (masterRows[i].length < totalOutputCols) {
        masterRows[i].push("");
      }
      existingBookingsMap.set(bNummer, i);
    }
  }

  // 3. Input-Daten verarbeiten
  var inputBuchungColIdx = inputHeaders.indexOf("Buchung Nummer");
  if (inputBuchungColIdx === -1) inputBuchungColIdx = buchungColIdx;

  for (var j = 1; j < inputData.length; j++) {
    var inputRow = inputData[j];
    var inputBuchungNummer = String(inputRow[inputBuchungColIdx] || "").trim();
    if (!inputBuchungNummer) continue;

    if (existingBookingsMap.has(inputBuchungNummer)) {
      // 🔒 FALL 1: BUCHUNG EXISTIERT BEREITS IM OUTPUT
      // Die Zeile wird KOMPLETT ÜBERSPRUNGEN und bleibt erhalten!
      // Einzige Ausnahme: Wir aktualisieren rein den Zahlungs-/Rechnungsstatus in Spalte L,
      // falls er im Input inzwischen auf "bezahlt" gesetzt wurde.
      var rowIdx = existingBookingsMap.get(inputBuchungNummer);
      var currentStatus = String(masterRows[rowIdx][outputStatusColIdx] || "").trim();
      var newStatus = String(inputRow[inputStatusColIdx] || "").trim();

      // Falls sich der Rechnungsstatus geändert hat:
      if (newStatus && currentStatus.toLowerCase() !== newStatus.toLowerCase()) {
        masterRows[rowIdx][outputStatusColIdx] = newStatus;
      }

    } else {
      // 🆕 FALL 2: BUCHUNG IST NEU
      // Wir bauen eine neue Zeile mit exakt der vollen Spaltenbreite von Output
      var newRow = new Array(totalOutputCols).fill("");
      
      // Nur die Spalten aus Input (A bis M) in die neue Zeile übernehmen
      for (var col = 0; col < inputRow.length && col < totalOutputCols; col++) {
        newRow[col] = inputRow[col];
      }

      masterRows.push(newRow);
      existingBookingsMap.set(inputBuchungNummer, masterRows.length - 1);
    }
  }

  // 4. ARCHIVIERUNG (Ganze Zeilen prüfen: Anreise 16 Uhr > 12h her)
  var now = new Date().getTime();
  var rowsToKeep = [];
  var rowsToArchive = [];

  for (var k = 0; k < masterRows.length; k++) {
    var row = masterRows[k];
    var anreiseVal = row[anreiseColIdx];
    var anreiseDatum = parseDateSafe(anreiseVal);

    if (anreiseDatum && !isNaN(anreiseDatum.getTime())) {
      anreiseDatum.setHours(16, 0, 0, 0);
      var archiveThreshold = anreiseDatum.getTime() + (12 * 60 * 60 * 1000);
      if (now > archiveThreshold) {
        rowsToArchive.push(row);
        continue;
      }
    }
    rowsToKeep.push(row);
  }

  // Archiv-Sheet beschreiben (falls Zeilen alt sind)
  if (rowsToArchive.length > 0) {
    var archiveSheet = ss.getSheetByName("Archive");
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet("Archive");
      archiveSheet.appendRow(outputHeaders);
    }
    var lastRowArchive = archiveSheet.getLastRow();
    if (lastRowArchive === 0) {
      archiveSheet.appendRow(outputHeaders);
      lastRowArchive = 1;
    }
    archiveSheet.getRange(lastRowArchive + 1, 1, rowsToArchive.length, totalOutputCols).setValues(rowsToArchive);
  }

  // 5. SORTIERUNG (Immer die GANZE Zeile sortieren, damit nie ein Spaltenversatz entsteht)
  rowsToKeep.sort(function(a, b) {
    var dateA = parseDateSafe(a[anreiseColIdx]);
    var dateB = parseDateSafe(b[anreiseColIdx]);
    var timeA = (dateA && !isNaN(dateA.getTime())) ? dateA.getTime() : 0;
    var timeB = (dateB && !isNaN(dateB.getTime())) ? dateB.getTime() : 0;
    return timeA - timeB;
  });

  // 6. OUTPUT-SHEET AKTUALISIEREN
  outputSheet.clearContents();
  var finalOutputData = [outputHeaders].concat(rowsToKeep);
  outputSheet.getRange(1, 1, finalOutputData.length, totalOutputCols).setValues(finalOutputData);
}

function parseDateSafe(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return new Date(value.getTime());
  }
  if (typeof value === "string") {
    value = value.trim();
    var deMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (deMatch) {
      return new Date(parseInt(deMatch[3], 10), parseInt(deMatch[2], 10) - 1, parseInt(deMatch[1], 10));
    }
    var usMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (usMatch) {
      return new Date(parseInt(usMatch[3], 10), parseInt(usMatch[2], 10) - 1, parseInt(usMatch[1], 10));
    }
    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }
  return null;
}