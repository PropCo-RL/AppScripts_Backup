function sendDriveFilesAsEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 2; // Startet bei Zeile 2
  var lastRow = sheet.getLastRow();
  var recipient = "l8@inbox.lexware.email";

  if (lastRow < startRow) {
    Logger.log("Keine Daten vorhanden.");
    return;
  }

  var range = sheet.getRange(startRow, 4, lastRow - startRow + 1, 3);
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    var url = values[i][0]; // Spalte D: Drive Link
    var status = values[i][2]; // Spalte F: Status
    var currentRow = startRow + i;

    if (url && url.toString().trim() !== "" && (!status || status.toString().trim() === "")) {
      try {
        var fileId = extractFileIdFromUrl(url.toString());
        if (fileId) {
          var file = DriveApp.getFileById(fileId);
          var fileName = file.getName().toLowerCase();
          var fileSizeKB = file.getSize() / 1024;

          // SICHERHEITSPRÜFUNG: Nur PDFs ab 30 KB versenden
          if (fileSizeKB < 30 || !fileName.endsWith(".pdf")) {
            sheet.getRange(currentRow, 6).setValue("Übersprungen: Kein PDF oder unter 30 KB");
            Logger.log("Zeile " + currentRow + " übersprungen (" + fileName + ", " + fileSizeKB.toFixed(1) + " KB)");
            continue;
          }

          // E-Mail senden
          GmailApp.sendEmail(recipient, "Rechnung: " + file.getName(), "Anbei befindet sich die angehängte Datei.", {
            attachments: [file.getAs(file.getMimeType())],
            name: "Automatischer Belegversand"
          });

          var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
          sheet.getRange(currentRow, 6).setValue("Anhang versendet am " + timestamp);
          Logger.log("Erfolgreich gesendet für Zeile " + currentRow);
        } else {
          sheet.getRange(currentRow, 6).setValue("Fehler: Ungültige Google Drive URL");
        }
      } catch (e) {
        sheet.getRange(currentRow, 6).setValue("Fehler: " + e.message);
        Logger.log("Fehler bei Zeile " + currentRow + ": " + e.message);
      }
    }
  }
}

function extractFileIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}