function updateApartmentCoordinates() {
  // Verbindet sich mit dem externen Apartment-Sheet über die ID aus deinem Link
  var aptSsId = "1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70";
  var aptSheet = SpreadsheetApp.openById(aptSsId).getSheetByName("m");
  
  if (!aptSheet) {
    Logger.log("FEHLER: Tab 'm' im Apartment-Sheet nicht gefunden.");
    return;
  }

  var lastRow = aptSheet.getLastRow();
  if (lastRow < 2) return;

  var colAddress = 6;  // Spalte F (Die Adresse)
  var colLat = 42;     // Spalte AP (Breitengrad)
  var colLng = 43;     // Spalte AQ (Längengrad)

  var dataRange = aptSheet.getRange(2, 1, lastRow - 1, Math.max(colAddress, colLng));
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var rowNum = i + 2;
    var address = values[i][colAddress - 1];
    var lat = values[i][colLat - 1];
    var lng = values[i][colLng - 1];

    // Nur prüfen, wenn Adresse existiert, aber noch keine Koordinaten da sind
    if (address && address.toString().trim() !== "" && (!lat || !lng)) {
      Logger.log("Generiere GPS-Daten für: " + address);
      Utilities.sleep(100); // Kurze Pause, um Google nicht zu überlasten
      
      try {
        var gc = Maps.newGeocoder().geocode(address);
        if (gc.status === "OK" && gc.results.length > 0) {
          var location = gc.results[0].geometry.location;
          aptSheet.getRange(rowNum, colLat).setValue(location.lat);
          aptSheet.getRange(rowNum, colLng).setValue(location.lng);
        }
      } catch(e) {
        Logger.log("Fehler bei: " + address);
      }
    }
  }
  Logger.log("Koordinaten-Update beendet.");
}