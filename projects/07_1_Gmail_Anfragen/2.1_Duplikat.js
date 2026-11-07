function processDMZ2Duplicates() {
  var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("DMZ2");
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // Nur Header oder leer

  // Header für Spalte B setzen, falls noch nicht vorhanden
  var headerB = sheet.getRange(1, 2).getValue();
  if (!headerB) {
    sheet.getRange(1, 2).setValue("Bereinigt / Unique");
    sheet.getRange(1, 2).setFontWeight("bold");
    sheet.getRange("B:B").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Spalte A und B komplett einlesen (ab Zeile 2 bis lastRow)
  var numRows = lastRow - 1;
  var rangeAB = sheet.getRange(2, 1, numRows, 2);
  var data = rangeAB.getValues();

  // 1. Bereits verarbeitete Unique-Einträge registrieren
  var seenFingerprints = new Set();

  for (var i = 0; i < data.length; i++) {
    var cellA = data[i][0];
    var cellB = data[i][1];

    // Wenn in Spalte B bereits ein echter JSON-Eintrag steht (kein "-" und nicht leer)
    if (cellB && cellB.toString().trim() !== "-" && cellB.toString().trim() !== "") {
      try {
        var existingJson = JSON.parse(cellB);
        var existingKey = generateFingerprint(existingJson);
        seenFingerprints.add(existingKey);
      } catch (e) {
        // Falls kein gültiges JSON in B steht, einfachen Text-Fingerabdruck nutzen
        seenFingerprints.add(cellB.toString().trim());
      }
    }
  }

  // 2. Offene Zeilen prüfen (Spalte A gefüllt, Spalte B leer)
  var updatesMade = false;

  for (var i = 0; i < data.length; i++) {
    var cellA = data[i][0] ? data[i][0].toString().trim() : "";
    var cellB = data[i][1] ? data[i][1].toString().trim() : "";

    // Nur ausführen, wenn A befüllt und B leer ist
    if (cellA !== "" && cellB === "") {
      try {
        var jsonObj = JSON.parse(cellA);
        var fingerprint = generateFingerprint(jsonObj);

        if (seenFingerprints.has(fingerprint)) {
          // Ist ein Duplikat
          data[i][1] = "-";
        } else {
          // Kein Duplikat -> exakte Kopie aus Spalte A
          data[i][1] = cellA;
          seenFingerprints.add(fingerprint);
        }
        updatesMade = true;
      } catch (e) {
        // Falls Zeile A kein valides JSON war
        data[i][1] = "-";
        updatesMade = true;
      }
    }
  }

  // 3. Zurückschreiben, wenn Änderungen vorhanden sind
  if (updatesMade) {
    rangeAB.setValues(data);
  }
}

// Hilfsfunktion: Baut einen eindeutigen Schlüssel aus den relevanten Inhalten
function generateFingerprint(item) {
  var name = (item.gastName || "").trim().toLowerCase();
  var phone = (item.telefonnummer || "").replace(/[^\d]/g, ""); // Nur Ziffern
  var period = (item.zeitraum || "").replace(/\s+/g, "");
  var persons = (item.personen || "").toString().trim();
  
  // Erste 80 Zeichen der Nachricht normalisiert (ohne Leerzeilen/Leerzeichen)
  var msg = (item.nachricht || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .substring(0, 80);

  // Kombinierter Identifikator
  return name + "|" + phone + "|" + period + "|" + persons + "|" + msg;
}