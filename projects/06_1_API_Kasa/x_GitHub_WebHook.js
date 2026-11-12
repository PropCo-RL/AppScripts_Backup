/**
 * Empfängt die fertigen Gerätedaten von GitHub Actions
 * und schreibt sie direkt in das Sheet.
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Sicherheitsprüfung über ein Secret Token
    const EXPECTED_TOKEN = "MeinKasaSecret2026"; // frei wählbares Passwort
    if (data.token !== EXPECTED_TOKEN) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Unauthorized" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Heizung_Status");
    if (!sheet) {
      sheet = ss.insertSheet("Heizung_Status");
    }

    const headers = [
      "Zuletzt geprüft",
      "Account (E-Mail)",
      "Gerätename / Standort",
      "Modell",
      "Status",
      "Ist-Temp (°C)",
      "Soll-Temp (°C)",
      "Batterie"
    ];

    sheet.clearContents();
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);

    if (data.rows && data.rows.length > 0) {
      sheet.getRange(2, 1, data.rows.length, data.rows[0].length).setValues(data.rows);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", count: data.rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}