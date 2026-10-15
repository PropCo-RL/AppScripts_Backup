function processInquiriesLiveORS() {
  // IDs beider Tabellen
  var aptSsId = "1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70";
  var anfrageSsId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
  
  var aptSheet = SpreadsheetApp.openById(aptSsId).getSheetByName("m");
  var anfrageSheet = SpreadsheetApp.openById(anfrageSsId).getSheetByName("dmz");
  
  // Liest den Schlüssel direkt unter dem Namen aus deinem Screenshot ("HeiGit.org-Key")
  var API_KEY = PropertiesService.getScriptProperties().getProperty('HeiGit.org-Key');
  if (!API_KEY) {
    Logger.log("FEHLER: 'HeiGit.org-Key' fehlt in den Script Properties.");
    return;
  }

  // --- 1. Alle Wohnungen aus Tab 'm' laden ---
  var aptData = aptSheet.getDataRange().getValues();
  var apartments = [];
  var aptLookup = {};
  
  for (var i = 1; i < aptData.length; i++) {
    var aptAddress = aptData[i][5]; // Spalte F (Adresse)
    var aptName = aptData[i][6];    // Spalte G (Bezeichnung, z.B. a/gs1)
    var lat = aptData[i][41];       // Spalte AP (Breitengrad)
    var lng = aptData[i][42];       // Spalte AQ (Längengrad)
    
    if (aptAddress && aptName && lat && lng) {
      var aptObj = {
        name: aptName.toString().trim(),
        address: aptAddress.toString().trim(),
        lat: parseFloat(lat),
        lng: parseFloat(lng)
      };
      apartments.push(aptObj);
      
      aptLookup[aptObj.name.toLowerCase()] = aptObj;
      aptLookup[aptObj.address.toLowerCase()] = aptObj;
    }
  }

  // --- 2. Anfragen im Tab 'dmz' durchgehen ---
  var lastRowAnfrage = anfrageSheet.getLastRow();
  if (lastRowAnfrage < 2) return;
  
  var anfrageData = anfrageSheet.getRange(2, 1, lastRowAnfrage - 1, 26).getValues();

  for (var j = 0; j < anfrageData.length; j++) {
    var rowNum = j + 2;
    var searchJ = anfrageData[j][9];   // Spalte J
    var searchM = anfrageData[j][12];  // Spalte M
    var statusS = anfrageData[j][18];  // Spalte S
    var resultZ = anfrageData[j][25];  // Spalte Z

    // CHECK 1: Überspringen, wenn in Spalte S "Dopplung" steht
    if (statusS && statusS.toString().trim().toLowerCase() === "dopplung") {
      continue;
    }

    // CHECK 2: Überspringen, falls Z bereits belegt ist
    if (resultZ && resultZ.toString().trim() !== "") {
      continue; 
    }

    var targetLat = null;
    var targetLng = null;

    // CHECK 3: Suchort ermitteln (Priorität: M, Fallback: J)
    if (searchM && searchM.toString().trim() !== "") {
      // HeiGIT Geocoding für Wunschort in M
      var searchLocation = searchM.toString().trim();
      var geocodeUrl = "https://api.heigit.org/geocode/search?api_key=" + API_KEY + "&text=" + encodeURIComponent(searchLocation);
      
      try {
        var geoRes = UrlFetchApp.fetch(geocodeUrl, { muteHttpExceptions: true });
        if (geoRes.getResponseCode() === 200) {
          var geoJson = JSON.parse(geoRes.getContentText());
          if (geoJson.features && geoJson.features.length > 0) {
            // HeiGIT/ORS liefert Koordinaten als [lng, lat]
            targetLng = geoJson.features[0].geometry.coordinates[0];
            targetLat = geoJson.features[0].geometry.coordinates[1];
          } else {
            anfrageSheet.getRange(rowNum, 26).setValue("Ort in M nicht gefunden");
            continue;
          }
        }
      } catch(e) { continue; }

    } else if (searchJ && searchJ.toString().trim() !== "") {
      var cleanJ = searchJ.toString().trim().toLowerCase();
      
      if (aptLookup[cleanJ]) {
        // PERFEKT: Direkt kostenlos aus Apartment-Sheet übernehmen!
        targetLat = aptLookup[cleanJ].lat;
        targetLng = aptLookup[cleanJ].lng;
      } else {
        // Geocoding als Notfall für unbekannte Adresse in J
        var geocodeUrlJ = "https://api.heigit.org/geocode/search?api_key=" + API_KEY + "&text=" + encodeURIComponent(searchJ.toString());
        try {
          var geoResJ = UrlFetchApp.fetch(geocodeUrlJ, { muteHttpExceptions: true });
          if (geoResJ.getResponseCode() === 200) {
            var geoJsonJ = JSON.parse(geoResJ.getContentText());
            if (geoJsonJ.features && geoJsonJ.features.length > 0) {
              targetLng = geoJsonJ.features[0].geometry.coordinates[0];
              targetLat = geoJsonJ.features[0].geometry.coordinates[1];
            } else {
              anfrageSheet.getRange(rowNum, 26).setValue("Adresse in J nicht gefunden");
              continue;
            }
          }
        } catch(e) { continue; }
      }
    } else {
      continue;
    }

    // --- 3. LUFTLINIEN-FILTER (max 40 km) ---
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

    // Nimm maximal die 40 besten Kandidaten
    if (candidates.length > 40) {
      candidates.sort(function(a, b) {
        return getDistanceInKm(targetLat, targetLng, a.lat, a.lng) - getDistanceInKm(targetLat, targetLng, b.lat, b.lng);
      });
      candidates = candidates.slice(0, 40);
    }

    // --- 4. HeiGIT MATRIX API (Fahrzeiten berechnen) ---
    var locations = [[targetLng, targetLat]];
    for (var c = 0; c < candidates.length; c++) {
      locations.push([candidates[c].lng, candidates[c].lat]);
    }

    var matrixPayload = {
      "locations": locations,
      "sources": [0],
      "metrics": ["duration"]
    };

    var matrixOptions = {
      "method": "post",
      "contentType": "application/json",
      "headers": {
        "Authorization": API_KEY
      },
      "payload": JSON.stringify(matrixPayload),
      "muteHttpExceptions": true
    };

    var validResults = [];
    try {
      var matrixRes = UrlFetchApp.fetch("https://api.heigit.org/v2/matrix/driving-car", matrixOptions);
      if (matrixRes.getResponseCode() === 200) {
        var matrixJson = JSON.parse(matrixRes.getContentText());
        var durations = matrixJson.durations[0];

        for (var m = 1; m < durations.length; m++) {
          var durationSec = durations[m];
          if (durationSec !== null) {
            var durationMin = Math.round(durationSec / 60);
            var aptInfo = candidates[m - 1];

            if (durationMin >= 0 && durationMin <= 40) {
              validResults.push({
                name: aptInfo.name,
                time: durationMin
              });
            }
          }
        }
      }
    } catch(e) {
      Logger.log("Matrix Fehler: " + e);
    }

    // --- 5. ERGEBNISSE IN SPALTE Z SCHREIBEN ---
    if (validResults.length > 0) {
      // Aufsteigend nach schnellster Fahrzeit sortieren
      validResults.sort(function(a, b) { return a.time - b.time; });

      var formattedStrings = validResults.map(function(item) {
        return item.name + " (" + item.time + "min)";
      });

      anfrageSheet.getRange(rowNum, 26).setValue(formattedStrings.join(", "));
    } else {
      anfrageSheet.getRange(rowNum, 26).setValue("Keine Ergebnisse (<40 Min Fahrt)");
    }
    
    SpreadsheetApp.flush();
  }
}

// Hilfsfunktion: Haversine-Formel für Luftlinie
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

