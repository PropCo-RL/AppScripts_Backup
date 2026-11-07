/**
 * Berechnet die Autofahrzeit zwischen zwei Orten.
 * Sucht zuerst im Tab 'Maps Archive'. Falls vorhanden, wird der gespeicherte Wert genommen.
 *
 * @param {string} start Startadresse
 * @param {string} ziel Zieladresse
 * @return {number|string} Fahrzeit als Sheets-Dauer ([h]:mm) oder Fehlermeldung
 * @customfunction
 */
function MAPS_FAHRZEIT(start, ziel) {
  var s = (start == null) ? "" : String(start).trim();
  var z = (ziel == null) ? "" : String(ziel).trim();

  if (s === "") return "ERROR: Start address is empty";
  if (z === "") return "ERROR: Destination address is empty";

  // 1. Zuerst im Tab "Maps Archive" nachsehen (Lesezugriff ist in Custom Functions erlaubt)
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var archiveSheet = ss.getSheetByName("Maps Archive");
    
    if (archiveSheet) {
      var data = archiveSheet.getDataRange().getValues();
      var sLower = s.toLowerCase();
      var zLower = z.toLowerCase();

      for (var i = 1; i < data.length; i++) {
        var archStart = String(data[i][0]).trim().toLowerCase();
        var archZiel = String(data[i][1]).trim().toLowerCase();

        if (archStart === sLower && archZiel === zLower) {
          // Bereits archivierten Wert (in Tagen) zurückgeben
          return data[i][2];
        }
      }
    }
  } catch (readError) {
    console.warn("Archiv konnte nicht gelesen werden: " + readError);
  }

  // 2. Nicht im Archiv gefunden -> Über Google Maps berechnen
  try {
    var directions = Maps.newDirectionFinder()
      .setOrigin(s)
      .setDestination(z)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (!directions) {
      return "ERROR: Google Maps returned no response";
    }

    if (directions.status && directions.status !== "OK") {
      return "ERROR: Google Maps status: " + directions.status;
    }

    if (!directions.routes || directions.routes.length === 0) {
      return "ERROR: No route found for these addresses";
    }

    var route = directions.routes[0];
    if (!route.legs || route.legs.length === 0) {
      return "ERROR: Route contains no legs";
    }

    var leg = route.legs[0];
    if (!leg.duration || typeof leg.duration.value !== "number") {
      return "ERROR: Total duration is missing";
    }

    var seconds = leg.duration.value;
    if (!isFinite(seconds) || seconds < 0) {
      return "ERROR: Invalid duration: " + seconds;
    }

    var minutes = Math.ceil(seconds / 60);

    // Wert in Tagen für das Format [h]:mm
    return minutes / 1440;

  } catch (error) {
    var message = error && error.message ? error.message : String(error);
    console.error("MAPS_FAHRZEIT ERROR: " + message);
    return "ERROR: " + message;
  }
}

/**
 * Automatischer Trigger: Sobald in Zeile 3 Adressen eingegeben/geändert werden,
 * wird das berechnete Ergebnis aus C3 fest in das Archivblatt geschrieben.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  // Name Ihres Eingabe-Sheets
  if (sheet.getName() !== "07.1_Gmail Anfragen") return;

  // Prüfen, ob Zelle A3 oder B3 bearbeitet wurde
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row !== 3 || (col !== 1 && col !== 2)) return;

  var start = String(sheet.getRange("A3").getValue()).trim();
  var ziel = String(sheet.getRange("B3").getValue()).trim();

  if (!start || !ziel) return;

  // Kurz warten, bis die Formel in C3 zu Ende gerechnet hat
  Utilities.sleep(1500);

  var fahrzeitVal = sheet.getRange("C3").getValue();

  // Nur speichern, wenn ein gültiger numerischer Wert vorliegt (kein Fehler-String / leer)
  if (typeof fahrzeitVal !== "number" || isNaN(fahrzeitVal)) return;

  var ss = e.source;
  var archiveSheet = ss.getSheetByName("Maps Archive");
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet("Maps Archive");
    archiveSheet.appendRow(["Start", "Ziel", "Fahrzeit (Tage)", "Datum"]);
    archiveSheet.getRange("C:C").setNumberFormat("[h]:mm");
  }

  // Prüfen, ob es während der Eingabe nicht bereits im Archiv existiert
  var data = archiveSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === start.toLowerCase() &&
        String(data[i][1]).trim().toLowerCase() === ziel.toLowerCase()) {
      return; // Bereits vorhanden
    }
  }

  // Dauerhaft fest im Archiv ablegen
  archiveSheet.appendRow([start, ziel, fahrzeitVal, new Date()]);
}