/**
 * Verfügbarkeits-Skript (gefiltert auf Blockierungen):
 * Schreibt NUR Zeiträume mit einer closed_period.id in den Tab "dbX".
 * Spalten: Property ID, Start, End, Closed Period ID
 */
function getLodgifyAvailability() {
  const props = PropertiesService.getScriptProperties();
  const loKey = props.getProperty('loKey');
  const fixieUrl = props.getProperty('fixieUrl');
  
  if (!loKey || !fixieUrl) {
    Logger.log("FEHLER: 'loKey' oder 'fixieUrl' fehlen!");
    return;
  }

  // --- DYNAMISCHE DATUM-BERECHNUNG ---
  const today = new Date();
  const startDate = new Date();
  startDate.setDate(today.getDate() - 30);
  const endDate = new Date();
  endDate.setDate(today.getDate() + 365);

  const startStr = Utilities.formatDate(startDate, "GMT+2", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const endStr = Utilities.formatDate(endDate, "GMT+2", "yyyy-MM-dd'T'HH:mm:ss'Z'");

  const targetUrl = `https://api.lodgify.com/v2/availability?start=${encodeURIComponent(startStr)}&end=${encodeURIComponent(endStr)}&includeDetails=true`;
  
  Logger.log("Rufe Verfügbarkeiten ab...");
  const responseData = sendThroughFixie(targetUrl, loKey, fixieUrl);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("dbX") || ss.insertSheet("dbX");

  // Inhalte ab Zeile 2 löschen (jetzt nur noch 4 Spalten relevant)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  if (!responseData) {
    Logger.log("Keine Daten erhalten.");
    return;
  }

  const items = Array.isArray(responseData) ? responseData : [responseData];
  
  // Header schreiben, falls Sheet neu ist (4 Spalten)
  if (sheet.getLastRow() === 0) {
    const headers = ["property_id", "periods.start", "periods.end", "periods.closed_period.id"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }

  let filteredRows = [];

  items.forEach(item => {
    if (item.periods && Array.isArray(item.periods)) {
      item.periods.forEach(p => {
        const cpId = (p.closed_period && p.closed_period.id) ? p.closed_period.id : "";
        
        // FILTER: Nur hinzufügen, wenn eine closed_period.id existiert
        if (cpId !== "") {
          filteredRows.push([
            item.property_id || "",
            p.start || "",
            p.end || "",
            cpId // Jetzt in Spalte D (Index 3)
          ]);
        }
      });
    }
  });

  // Daten ab Zeile 2 schreiben
  if (filteredRows.length > 0) {
    sheet.getRange(2, 1, filteredRows.length, 4).setValues(filteredRows);
    Logger.log(filteredRows.length + " Blockierungen (4 Spalten) in dbX eingetragen.");
  } else {
    Logger.log("Keine Blockierungen gefunden.");
  }
}

/**
 * Hilfsfunktion Fixie
 */
function sendThroughFixie(targetUrl, loKey, fixieUrl) {
  const authCredentials = fixieUrl.split('@')[0].replace('http://', '');
  const options = {
    "method": "get",
    "headers": {
      "X-ApiKey": loKey,
      "accept": "application/json",
      "Proxy-Authorization": "Basic " + Utilities.base64Encode(authCredentials)
    },
    "muteHttpExceptions": true
  };
  const response = UrlFetchApp.fetch(targetUrl, options);
  return JSON.parse(response.getContentText());
}