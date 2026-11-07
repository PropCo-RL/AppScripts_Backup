/**
 * Webhook-Handler für Google Sheets
 * Speichert eintreffende JSON-Daten in das Tabellenblatt "WH"
 */
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Blatt "Raw" suchen oder erstellen, falls es nicht existiert
  let sheet = ss.getSheetByName("WH");
  if (!sheet) {
    sheet = ss.insertSheet("WH");
  }

  try {
    const rawData = e && e.postData ? e.postData.contents : "";
    const timestamp = new Date();

    // 👉 Header und Layout beim ersten Mal einrichten
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Zeitstempel", "JSON Payload"]);
      sheet.setColumnWidth(1, 180); // Datum braucht etwas mehr Platz
      sheet.setColumnWidth(2, 800);
      sheet.getRange("1:1").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // JSON validieren
    let dataToStore;
    try {
      dataToStore = rawData ? JSON.parse(rawData) : { "info": "Kein Inhalt empfangen" };
    } catch (parseError) {
      dataToStore = { "raw_text": rawData, "error": "Ungültiges JSON" };
    }

    // 👉 Den Eintrag hinzufügen
    sheet.appendRow([
      timestamp,
      JSON.stringify(dataToStore)
    ]);

    // Zeilenhöhe auf Standard setzen (verhindert riesige Zellen bei langen JSONs)
    sheet.setRowHeight(sheet.getLastRow(), 21);

  } catch (err) {
    // Fehler-Logging direkt im Sheet
    sheet.appendRow([
      new Date(),
      "SYSTEM ERROR: " + err.toString()
    ]);
  }

  // Erfolgsmeldung an den Absender
  return ContentService.createTextOutput("OK");
}