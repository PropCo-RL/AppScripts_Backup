function processMMZDuplicates() {
  var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("MMZ");
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Header für Spalte L setzen
  var headerL = sheet.getRange(1, 12).getValue();
  if (!headerL) {
    sheet.getRange(1, 12).setValue("Bereinigt / Unique");
    sheet.getRange(1, 12).setFontWeight("bold");
    sheet.getRange("L:L").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Spalten K und L einlesen (ab Zeile 2, Spalte 11, 2 Spalten breit)
  var numRows = lastRow - 1;
  var rangeKL = sheet.getRange(2, 11, numRows, 2);
  var data = rangeKL.getValues();

  var TWENTY_FOUR_HOURS_MS = 24 * 60 * 60 * 1000;

  // Speichert: Fingerprint -> Zeitstempel in Millisekunden
  var seenFingerprints = new Map();

  // 1. Bereits vorhandene Uniques registrieren
  for (var i = 0; i < data.length; i++) {
    var cellL = data[i][1];
    if (cellL && cellL.toString().trim() !== "-" && cellL.toString().trim() !== "") {
      try {
        var existingJson = JSON.parse(cellL);
        var fp = generateFingerprintMMZ(existingJson);
        var timeMs = extractTimestampMs(existingJson);
        seenFingerprints.set(fp, timeMs);
      } catch (e) {
        seenFingerprints.set(cellL.toString().trim(), 0);
      }
    }
  }

  // 2. Offene Zeilen prüfen (Spalte K vorhanden, Spalte L noch leer)
  var updatesMade = false;
  for (var i = 0; i < data.length; i++) {
    var cellK = data[i][0] ? data[i][0].toString().trim() : "";
    var cellL = data[i][1] ? data[i][1].toString().trim() : "";

    if (cellK !== "" && cellL === "") {
      try {
        var jsonObj = JSON.parse(cellK);

        // Standard-Floskel in echten Leerstring "" wandeln für spätere Skripte
        if ((jsonObj.nachricht || "").trim().toLowerCase() === "keine nachricht hinterlassen") {
          jsonObj.nachricht = "";
        }

        var fingerprint = generateFingerprintMMZ(jsonObj);
        var currentTimeMs = extractTimestampMs(jsonObj);

        var isDuplicate = false;
        if (seenFingerprints.has(fingerprint)) {
          var lastSeenTimeMs = seenFingerprints.get(fingerprint);
          
          // Duplikat nur, wenn die Zeitdifferenz <= 24 Stunden beträgt
          if (currentTimeMs > 0 && lastSeenTimeMs > 0) {
            if (Math.abs(currentTimeMs - lastSeenTimeMs) <= TWENTY_FOUR_HOURS_MS) {
              isDuplicate = true;
            }
          } else {
            // Fallback falls kein Datum vorhanden: klassischer Duplikat-Ausschluss
            isDuplicate = true;
          }
        }

        if (isDuplicate) {
          // Innerhalb von 24h bereits vorhanden -> Duplikat
          data[i][1] = "-";
        } else {
          // Neu oder älter als 24h -> Als Unique übernehmen
          data[i][1] = JSON.stringify(jsonObj);
          seenFingerprints.set(fingerprint, currentTimeMs);
        }
        updatesMade = true;
      } catch (e) {
        data[i][1] = "-";
        updatesMade = true;
      }
    }
  }

  // 3. Zurückschreiben
  if (updatesMade) {
    rangeKL.setValues(data);
  }
}

// Hilfsfunktion: Zeitstempel gezielt aus anfrageErstelltAm bzw. erfasstAm auslesen
function extractTimestampMs(item) {
  var rawDate = item.anfrageErstelltAm || item.erfasstAm || item.datum || item.date || "";
  if (!rawDate) return 0;

  var parsed = new Date(rawDate).getTime();
  return isNaN(parsed) ? 0 : parsed;
}

// Fingerprint-Funktion angepasst an die Eigenschaften von Mein-Monteurzimmer
function generateFingerprintMMZ(item) {
  // 1. Firmenzusätze in Klammern abschneiden: "Marvin Strubel (L8 Street GmbH)" -> "marvin strubel"
  var name = (item.gastName || "")
    .replace(/\s*\([^)]*\)/g, "")
    .trim().toLowerCase();

  // 2. Nur Ziffern der Telefonnummer (letzte 8 Stellen genügen zur Ländervorwahl-Bereinigung)
  var rawPhone = (item.telefonnummer || "").replace(/[^\d]/g, "");
  var phone = rawPhone.length > 8 ? rawPhone.slice(-8) : rawPhone;

  // 3. E-Mail normalisieren
  var email = (item.email || "").trim().toLowerCase();
  if (email === "nicht gefunden") email = "";

  // 4. Zeitraum & Personen
  var period = (item.zeitraum || "").replace(/\s+/g, "");
  var persons = (item.personen || "").toString().trim();

  // 5. Nachricht: Leere Nachrichten/Standard-Texte normalisieren
  var msg = (item.nachricht || "").trim();
  if (msg.toLowerCase() === "keine nachricht hinterlassen") {
    msg = "";
  } else {
    msg = msg.toLowerCase().replace(/\s+/g, "").substring(0, 80);
  }

  return name + "|" + phone + "|" + email + "|" + period + "|" + persons + "|" + msg;
}