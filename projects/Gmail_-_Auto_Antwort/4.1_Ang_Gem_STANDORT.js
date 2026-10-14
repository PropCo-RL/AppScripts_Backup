function processInquiriesLive() {
  // IDs beider Tabellen (aus deinen Links)
  var aptSsId = "1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70";
  var anfrageSsId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
  
  var aptSheet = SpreadsheetApp.openById(aptSsId).getSheetByName("m");
  var anfrageSheet = SpreadsheetApp.openById(anfrageSsId).getSheetByName("dmz");
  
  var API_KEY = PropertiesService.getScriptProperties().getProperty('G-Maps-Key');
  if (!API_KEY) {
    Logger.log("FEHLER: API Key fehlt in den Script Properties.");
    return;
  }

  // --- 1. Alle Wohnungen & vorberechnete Koordinaten aus Tab 'm' laden ---
  var aptData = aptSheet.getDataRange().getValues();
  var apartments = [];
  var aptLookup = {}; // Schnelles Nachschlagewerk für kostenlose Abfragen
  
  for (var i = 1; i < aptData.length; i++) {
    var aptAddress = aptData[i][5]; // Spalte F (Adresse)
    var aptName = aptData[i][6];    // Spalte G (Bezeichnung, z.B. a/gs1)
    var lat = aptData[i][41];       // Spalte AP (Index 41: Breitengrad)
    var lng = aptData[i][42];       // Spalte AQ (Index 42: Längengrad)
    
    if (aptAddress && aptName && lat && lng) {
      var aptObj = {
        name: aptName.toString().trim(),
        address: aptAddress.toString().trim(),
        lat: lat,
        lng: lng
      };
      apartments.push(aptObj);
      
      // Sowohl Name als auch Adresse für direkten Match abspeichern
      aptLookup[aptObj.name.toLowerCase()] = aptObj;
      aptLookup[aptObj.address.toLowerCase()] = aptObj;
    }
  }

  // --- 2. Anfragen im Tab 'dmz' durchgehen ---
  var lastRowAnfrage = anfrageSheet.getLastRow();
  if (lastRowAnfrage < 2) return;
  
  // Liest bis Spalte Z (26 Spalten)
  var anfrageData = anfrageSheet.getRange(2, 1, lastRowAnfrage - 1, 26).getValues();

  for (var j = 0; j < anfrageData.length; j++) {
    var rowNum = j + 2;
    var searchJ = anfrageData[j][9];   // Spalte J (Index 9)
    var searchM = anfrageData[j][12];  // Spalte M (Index 12)
    var statusS = anfrageData[j][18];  // Spalte S (Index 18)
    var resultZ = anfrageData[j][25];  // Spalte Z (Index 25)

    // CHECK 1: Überspringen, wenn in Spalte S "Dopplung" steht
    if (statusS && statusS.toString().trim().toLowerCase() === "dopplung") {
      continue;
    }

    // CHECK 2: Überspringen, falls in Spalte Z schon etwas steht
    if (resultZ && resultZ.toString().trim() !== "") {
      continue; 
    }

    var targetLat = null;
    var targetLng = null;
    var targetLocationName = "";

    // CHECK 3: Suchort ermitteln (Priorität: M, Fallback: J)
    if (searchM && searchM.toString().trim() !== "") {
      // Fall 1: Kunden-Wunschort in Spalte M -> Geocoding API nutzen
      targetLocationName = searchM.toString().trim();
      try {
        var gc = Maps.newGeocoder().geocode(targetLocationName);
        if (gc.status === "OK" && gc.results.length > 0) {
          targetLat = gc.results[0].geometry.location.lat;
          targetLng = gc.results[0].geometry.location.lng;
        } else {
          anfrageSheet.getRange(rowNum, 26).setValue("Ort in M nicht gefunden");
          continue;
        }
      } catch(e) { continue; }

    } else if (searchJ && searchJ.toString().trim() !== "") {
      // Fall 2: Fallback auf Spalte J -> Prüfe zuerst auf kostenlosen Match
      var cleanJ = searchJ.toString().trim().toLowerCase();
      
      if (aptLookup[cleanJ]) {
        // PERFEKT: Adresse/Name im Apartment-Sheet bekannt -> 0 API-Kosten!
        targetLat = aptLookup[cleanJ].lat;
        targetLng = aptLookup[cleanJ].lng;
      } else {
        // Unbekannte Adresse in J -> Geocoding als Notfall
        try {
          var gcNotfall = Maps.newGeocoder().geocode(searchJ.toString());
          if (gcNotfall.status === "OK" && gcNotfall.results.length > 0) {
            targetLat = gcNotfall.results[0].geometry.location.lat;
            targetLng = gcNotfall.results[0].geometry.location.lng;
          } else {
            anfrageSheet.getRange(rowNum, 26).setValue("Adresse in J nicht gefunden");
            continue;
          }
        } catch(e) { continue; }
      }
    } else {
      // Weder M noch J sind ausgefüllt -> Überspringen
      continue;
    }

    // --- 3. LUFTLINIEN-FILTER (max 40 km Radius, kostenlos) ---
    var candidates = [];
    for (var a = 0; a < apartments.length; a++) {
      var distKm = getDistanceInKm(targetLat, targetLng, apartments[a].lat, apartments[a].lng);
      if (distKm <= 40) { 
        candidates.push(apartments[a]);
      }
    }

    if (candidates.length === 0) {
      anfrageSheet.getRange(rowNum, 26).setValue("Keine Ergebnisse (<40km Luftlinie)");
      continue;
    }

    // --- 4. KOSTEN-SICHERHEITSBREMSE (Max. 40 beste Kandidaten) ---
    if (candidates.length > 40) {
      candidates.sort(function(a, b) {
        return getDistanceInKm(targetLat, targetLng, a.lat, a.lng) - getDistanceInKm(targetLat, targetLng, b.lat, b.lng);
      });
      candidates = candidates.slice(0, 40);
    }

    // --- 5. EXAKTE ROUTES API (Fahrzeiten berechnen) ---
    var validResults = [];
    var departureTime = new Date();
    departureTime.setDate(departureTime.getDate() + 1);
    departureTime.setHours(5, 0, 0, 0); // Typische Berufsverkehrszeit 05:00 Uhr

    for (var c = 0; c < candidates.length; c += 25) {
      var chunk = candidates.slice(c, c + 25);
      var payload = {
        "origins": [{ "waypoint": { "location": { "latLng": { "latitude": targetLat, "longitude": targetLng } } } }],
        "destinations": chunk.map(function(apt) { return { "waypoint": { "address": apt.address } }; }),
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
      if (response.getResponseCode() === 200) {
        var elements = JSON.parse(response.getContentText());
        for (var k = 0; k < elements.length; k++) {
          var elem = elements[k];
          if (elem.duration) {
            var durationSec = parseInt(elem.duration.replace("s", ""), 10);
            var durationMin = Math.round(durationSec / 60);
            var aptInfo = chunk[elem.destinationIndex];

            // Nur Ergebnisse bis maximal 40 Minuten aufnehmen
            if (durationMin >= 0 && durationMin <= 40) {
              validResults.push({
                name: aptInfo.name,
                time: durationMin
              });
            }
          }
        }
      }
    }

    // --- 6. ERGEBNISSE SORTIEREN & IN SPALTE Z SCHREIBEN ---
    if (validResults.length > 0) {
      // Aufsteigend nach Fahrtzeit sortieren (schnellste zuerst)
      validResults.sort(function(a, b) { return a.time - b.time; });

      // Formatieren zu "a/gs1 (17min)"
      var formattedStrings = validResults.map(function(item) {
        return item.name + " (" + item.time + "min)";
      });

      // Mit Komma und Leerzeichen zusammenfügen
      anfrageSheet.getRange(rowNum, 26).setValue(formattedStrings.join(", "));
    } else {
      anfrageSheet.getRange(rowNum, 26).setValue("Keine Ergebnisse (<40 Min Fahrt)");
    }
    
    SpreadsheetApp.flush(); // Zeigt Fortschritt live im Sheet an
  }
}

// Hilfsfunktion: Haversine-Formel zur schnellen Berechnung der Luftlinie (in km)
function getDistanceInKm(lat1, lon1, lat2, lon2) {
  var R = 6371; 
  var dLat = (lat2 - lat1) * Math.PI / 180;  
  var dLon = (lon2 - lon1) * Math.PI / 180; 
  var a = 
    0.5 - Math.cos(dLat)/2 + 
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * 
    (1 - Math.cos(dLon))/2;
  return R * 2 * Math.asin(Math.sqrt(a));
}