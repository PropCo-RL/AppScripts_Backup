function processAllAddressesLive() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return;

  var API_KEY = PropertiesService.getScriptProperties().getProperty('G-Maps-Key');
  if (!API_KEY) {
    Logger.log("FEHLER: API Key fehlt.");
    return;
  }

  // Spalten-Konfiguration (1-basiert)
  var colAddress = 6;  // Spalte F (Startadresse)
  var col10Min   = 37; // Spalte AK
  var col20Min   = 38; // Spalte AL
  var col30Min   = 39; // Spalte AM
  var col40Min   = 40; // Spalte AN
  var colStatus  = 41; // Spalte AO (Status)

  var range = sheet.getRange(2, 1, lastRow - 1, colStatus);
  var values = range.getValues();

  // 1. Alle Zeilen nach identischer Startadresse gruppieren
  var addressGroups = {};

  for (var i = 0; i < values.length; i++) {
    var rowIndex = i + 2;
    var address = values[i][colAddress - 1];
    var status = values[i][colStatus - 1];

    if (address && address.toString().trim() !== "") {
      var addrStr = address.toString().trim();

      if (!addressGroups[addrStr]) {
        addressGroups[addrStr] = {
          rows: [],
          isDone: true // Gehen wir erstmal davon aus, dass sie fertig ist
        };
      }
      
      addressGroups[addrStr].rows.push(rowIndex);
      
      // Sobald mindestens EINE Zeile dieser Adresse nicht fertig ist, muss die Adresse berechnet werden
      if (status !== "Google Maps Done") {
        addressGroups[addrStr].isDone = false;
      }
    }
  }

  // 2. Jede benötigte Adresse genau EINMAL abfragen
  for (var addr in addressGroups) {
    var group = addressGroups[addr];

    // Überspringen, falls schon alle Zeilen dieser Adresse "Google Maps Done" sind
    if (group.isDone) {
      Logger.log("Überspringe (bereits erledigt): " + addr);
      continue; 
    }

    Logger.log("Verarbeite Adresse (schreibe in " + group.rows.length + " Zeilen gleichzeitig): " + addr);

    // --- Start der API-Logik ---
    var match = addr.match(/\b\d{5}\b/);
    var basePlz = match ? parseInt(match[0], 10) : null;

    if (!basePlz) {
      var geocoder = Maps.newGeocoder().geocode(addr);
      if (geocoder.status === "OK" && geocoder.results.length > 0) {
        var components = geocoder.results[0].address_components;
        for (var c = 0; c < components.length; c++) {
          if (components[c].types.indexOf("postal_code") !== -1) {
            basePlz = parseInt(components[c].long_name, 10);
            break;
          }
        }
      }
    }

    if (basePlz === null || isNaN(basePlz)) {
      Logger.log("FEHLER: Konnte keine Start-PLZ ermitteln für: " + addr);
      continue; // Geht zur nächsten Adresse
    }

    // Umkreis-Kandidaten generieren (RADIUS auf 2500 erhöht für Rhein-Neckar & Co.)
    var candidateDestinations = [];
    var startPlz = Math.max(0, basePlz - 2500); 
    var endPlz   = Math.min(99999, basePlz + 2500);

    for (var p = startPlz; p <= endPlz; p++) {
      var plzStr = ("00000" + p).slice(-5);
      candidateDestinations.push(plzStr + ", Germany");
    }

    // Abfahrtszeit: Morgen 05:00 Uhr
    var departureTime = new Date();
    departureTime.setDate(departureTime.getDate() + 1);
    departureTime.setHours(5, 0, 0, 0);

    var results = { min10: [], min20: [], min30: [], min40: [] };

    // Eigene Start-PLZ immer fest als Erstes eintragen
    var startPlzFormatted = ("00000" + basePlz).slice(-5);
    var startFullName = resolveCityName(startPlzFormatted);
    results.min10.push(startFullName);

    // 3. Fahrzeiten in 25er-Blöcken abfragen
    for (var i = 0; i < candidateDestinations.length; i += 25) {
      var chunk = candidateDestinations.slice(i, i + 25);
      
      var payload = {
        "origins": [{ "waypoint": { "address": addr } }],
        "destinations": chunk.map(function(dest) { return { "waypoint": { "address": dest } }; }),
        "travelMode": "DRIVE",
        "routingPreference": "TRAFFIC_AWARE",
        "departureTime": departureTime.toISOString()
      };

      var options = {
        "method": "post",
        "contentType": "application/json",
        "headers": {
          "X-Goog-Api-Key": API_KEY,
          "X-Goog-FieldMask": "destinationIndex,duration"
        },
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      };

      var response = UrlFetchApp.fetch("https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix", options);
      var hasNewResults = false;
      
      if (response.getResponseCode() === 200) {
        var elements = JSON.parse(response.getContentText());
        
        for (var k = 0; k < elements.length; k++) {
          var elem = elements[k];
          if (elem.duration) {
            var durationSec = parseInt(elem.duration.replace("s", ""), 10);
            var durationMin = Math.round(durationSec / 60);
            var rawPlz = chunk[elem.destinationIndex].substring(0, 5);

            // Geändert: >= 0, um 0-Minuten Distanzen zu akzeptieren
            if (durationMin >= 0 && durationMin <= 40) {
              var fullName = resolveCityName(rawPlz);
              
              // Duplikatprüfung
              if (durationMin <= 10) {
                if (results.min10.indexOf(fullName) === -1) results.min10.push(fullName);
              } else if (durationMin <= 20) {
                if (results.min20.indexOf(fullName) === -1) results.min20.push(fullName);
              } else if (durationMin <= 30) {
                if (results.min30.indexOf(fullName) === -1) results.min30.push(fullName);
              } else if (durationMin <= 40) {
                if (results.min40.indexOf(fullName) === -1) results.min40.push(fullName);
              }
              
              hasNewResults = true;
            }
          }
        }
      }

      // --- LIVE-SCHREIBEN INS SHEET ---
      // Nach jedem 25er-Block werden alle zugehörigen Zeilen live geupdatet
      if (hasNewResults) {
        for (var r = 0; r < group.rows.length; r++) {
          var rowNum = group.rows[r];
          if (results.min10.length > 0) sheet.getRange(rowNum, col10Min).setValue(results.min10.join(",\n"));
          if (results.min20.length > 0) sheet.getRange(rowNum, col20Min).setValue(results.min20.join(",\n"));
          if (results.min30.length > 0) sheet.getRange(rowNum, col30Min).setValue(results.min30.join(",\n"));
          if (results.min40.length > 0) sheet.getRange(rowNum, col40Min).setValue(results.min40.join(",\n"));
        }
        SpreadsheetApp.flush(); // Erzwingt die sofortige visuelle Aktualisierung im Sheet
      }
    } // Ende der 25er-Blöcke (für diese eine Adresse)

    // 4. Nach Abschluss der kompletten Adresse den Done-Status setzen
    for (var r = 0; r < group.rows.length; r++) {
      var rowNum = group.rows[r];
      sheet.getRange(rowNum, colStatus).setValue("Google Maps Done");
    }
    SpreadsheetApp.flush();
    Logger.log("Adresse vollständig abgeschlossen: " + addr);
    
  } // Ende der Adress-Schleife

  Logger.log("Gesamter Vorgang abgeschlossen.");
}

// Hilfsfunktion: Wandelt PLZ in "PLZ Ortsname" um
function resolveCityName(plz) {
  Utilities.sleep(50); // Pause gegen Google-Ratelimits
  try {
    var gc = Maps.newGeocoder().geocode(plz + " Deutschland");
    if (gc.status === "OK" && gc.results.length > 0) {
      var components = gc.results[0].address_components;
      for (var i = 0; i < components.length; i++) {
        if (components[i].types.indexOf("locality") !== -1 || components[i].types.indexOf("postal_town") !== -1) {
          return plz + " " + components[i].long_name;
        }
      }
    }
  } catch(e) {}
  return plz;
}