function findAvailableApartments() {
  const sheet1Id = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const sheet2Id = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';

  const s1 = SpreadsheetApp.openById(sheet1Id).getSheetByName('DMZ');
  const s2_m = SpreadsheetApp.openById(sheet2Id).getSheetByName('M');
  const s2_cal = SpreadsheetApp.openById(sheet2Id).getSheetByName('cal');

  const dmzData = s1.getDataRange().getValues();
  const mData = s2_m.getDataRange().getValues();
  const calData = s2_cal.getDataRange().getValues();

  Logger.log("--- SKRIPT START (Zeitzonen-neutral) ---");

  // 1. Alle gebuchten Zeiträume aus "cal" einlesen
  let bookings = {};
  for (let i = 1; i < calData.length; i++) {
    let aptName = String(calData[i][0]).trim();
    let start = parseIsoDate(calData[i][1]); // Spalte B
    let end = parseIsoDate(calData[i][2]);   // Spalte C

    if (aptName && start !== null && end !== null) {
      if (!bookings[aptName]) bookings[aptName] = [];
      bookings[aptName].push({ start: start, end: end });
    }
  }
  Logger.log("Buchungen für " + Object.keys(bookings).length + " Wohnungen geladen.");

  // 2. Alle Wohnungen aus "M" einlesen
  let apartments = [];
  for (let i = 1; i < mData.length; i++) {
    let aptName = String(mData[i][0]).trim();
    let slug = String(mData[i][6] || "").trim(); // Spalte G
    let regionsString = mData[i][22];            // Spalte W

    let regionsList = [];
    if (regionsString && typeof regionsString === 'string') {
      regionsList = regionsString.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    }

    if (aptName) {
      apartments.push({ name: aptName, slug: slug, regions: regionsList });
    }
  }
  Logger.log(apartments.length + " Wohnungen aus Stammdaten geladen.\n");

  // 3. Anfragen durchgehen und abgleichen
  let outputColO = [];
  outputColO.push(["Verfügbare Links"]); // Header

  for (let i = 1; i < dmzData.length; i++) {
    let dateStr = dmzData[i][4];    // Spalte E
    let addressStr = dmzData[i][9]; // Spalte J

    if (!dateStr || !addressStr) {
      outputColO.push([""]);
      continue;
    }

    // Stadt flexibler extrahieren (4 oder 5 stellige PLZ - wichtig für CH/DE!)
    let cityMatch = String(addressStr).match(/\d{4,5}\s+([a-zA-ZäöüÄÖÜß\-\s]+)/);
    let reqCity = cityMatch ? cityMatch[1].trim().toLowerCase() : "";

    let dates = parseGermanDateRange(dateStr);

    if (!reqCity || !dates) {
      Logger.log("Zeile " + (i + 1) + " -> FEHLER: Stadt oder Datum unlesbar.");
      outputColO.push(["Fehler: Stadt oder Datum unlesbar"]);
      continue;
    }

    // Mindestaufenthalt prüfen (min. 5 Tage)
    let nightCount = Math.round((dates.end - dates.start) / (1000 * 60 * 60 * 24));
    if (nightCount < 5) {
      Logger.log("Zeile " + (i + 1) + " -> ABBRUCH: Zu kurz (" + nightCount + " Nächte).");
      outputColO.push(["Unter Mindestaufenthalt (min. 5 Tage)"]);
      continue;
    }

    Logger.log("Zeile " + (i + 1) + " | Anfrage für '" + reqCity + "' | Nächte: " + nightCount);

    let availableLinks = [];

    for (let apt of apartments) {
      // Prüft ob die Stadt in der Regionen-Liste der Wohnung vorkommt (Teil-Strings erlaubt)
      let matchRegion = apt.regions.some(r => reqCity.includes(r) || r.includes(reqCity));

      if (matchRegion) {
        let isAvailable = true;

        if (bookings[apt.name]) {
          for (let b of bookings[apt.name]) {
            // Überschneidungsformel: Neuer Start muss kleiner als altes Ende sein UND
            // Neues Ende muss größer als alter Start sein (Wechseltag am gleichen Datum erlaubt!)
            if (dates.start < b.end && dates.end > b.start) {
              isAvailable = false;
              break;
            }
          }
        }

        if (isAvailable && apt.slug) {
          let cleanSlug = apt.slug.startsWith('/') ? apt.slug : '/' + apt.slug;
          let link = "https://a.l8street.com" + cleanSlug;
          if (!link.endsWith('/')) link += '/';
          availableLinks.push(link);
          Logger.log("  -> [FREI] " + apt.name);
        } else if (!isAvailable) {
          Logger.log("  -> [BELEGT] " + apt.name);
        } else if (!apt.slug) {
          Logger.log("  -> [KEIN LINK] " + apt.name + " (Spalte G ist leer)");
        }
      }
    }

    if (availableLinks.length > 0) {
      outputColO.push([availableLinks.join(', ')]);
    } else {
      outputColO.push(["Keine verfügbar"]);
    }
  }

  // 4. In Spalte O schreiben
  s1.getRange(1, 15, outputColO.length, 1).setValues(outputColO);
  Logger.log("\n--- SKRIPT ENDE (Daten erfolgreich geschrieben) ---");

  // NEU: Direkt danach den E-Mail-Versand auslösen
  sendOfferEmail();

} // Ende der Funktion findAvailableApartments

// ----------------------------------------------------------------------
// HILFSFUNKTIONEN (Zeitzonen-neutral auf exakt 00:00 Uhr gestellt)
// ----------------------------------------------------------------------

function parseGermanDateRange(dateStr) {
  try {
    let parts = String(dateStr).split('-');
    if (parts.length !== 2) return null;

    let d1 = parts[0].trim().split('.');
    let d2 = parts[1].trim().split('.');

    // Date.UTC erzwingt einen reinen Zahlwert (Timestamp) ohne Zeitzone
    let start = Date.UTC(d1[2], d1[1] - 1, d1[0]);
    let end = Date.UTC(d2[2], d2[1] - 1, d2[0]);

    return { start: start, end: end };
  } catch (e) {
    return null;
  }
}

function parseIsoDate(isoStr) {
  try {
    if (!isoStr) return null;
    
    // Die wichtigste Änderung: Ignoriert Zürich/Schweizer Zeitverschiebungen komplett!
    if (isoStr instanceof Date) {
      return Date.UTC(isoStr.getFullYear(), isoStr.getMonth(), isoStr.getDate());
    }
    
    let parts = String(isoStr).trim().split('-');
    if (parts.length !== 3) return null;
    return Date.UTC(parts[0], parts[1] - 1, parts[2]);
  } catch (e) {
    return null;
  }
}
