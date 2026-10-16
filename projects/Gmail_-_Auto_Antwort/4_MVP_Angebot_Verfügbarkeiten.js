function findAvailableApartments() {
  const sheet1Id = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const sheet2Id = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';
  const s1 = SpreadsheetApp.openById(sheet1Id).getSheetByName('DMZ');
  const s2_m = SpreadsheetApp.openById(sheet2Id).getSheetByName('M');
  const s2_cal = SpreadsheetApp.openById(sheet2Id).getSheetByName('cal');
  const dmzData = s1.getDataRange().getValues();
  const mData = s2_m.getDataRange().getValues();
  const calData = s2_cal.getDataRange().getValues();
  
  Logger.log("--- SKRIPT START ---");

  // 1. Gebuchte Zeiträume aus "cal" einlesen
  let bookings = parseBookings(calData);

  // 2. Stammdaten der Wohnungen aus "M" einlesen
  let apartments = parseApartments(mData);

  let outputColZ = [["Verfügbare Links"]];
  const ONE_DAY_MS = 1000 * 60 * 60 * 24;

  // 3. Anfragen aus "DMZ" verarbeiten
  for (let i = 1; i < dmzData.length; i++) {
    let colA = String(dmzData[i][0] || "").trim(); // Spalte A (Index 0)

    // Abbruch, sobald Spalte A leer ist (keine weiteren Zeilen mehr verarbeiten)
    if (colA === "") {
      break;
    }

    let existingZ = String(dmzData[i][25] || "").trim(); // Spalte Z (Index 25)
    let colS = String(dmzData[i][18] || "").trim();      // Spalte S (Index 18)
    let colT = String(dmzData[i][19] || "").trim();      // Spalte T (Index 19)

    // Überspringen, falls in Spalte Z bereits etwas eingetragen ist...
    // ...oder falls S UND T beide befüllt sind
    if (existingZ !== "" || (colS !== "" && colT !== "")) {
      outputColZ.push([dmzData[i][25] || ""]);
      continue;
    }

    let dateStr = dmzData[i][14] || dmzData[i][4]; // O oder E
    let reqPersons = parseInt(dmzData[i][13]) || parseInt(dmzData[i][5]) || 1; // N oder F
    let addressStr = dmzData[i][9]; // J

    if (!dateStr || !addressStr) {
      outputColZ.push([""]);
      continue;
    }

    let cityMatch = String(addressStr).match(/\d{4,5}\s+([a-zA-ZäöüÄÖÜß\-\s]+)/);
    let reqCity = cityMatch ? cityMatch[1].trim().toLowerCase() : String(addressStr).trim().toLowerCase();
    let dates = parseGermanDateRange(dateStr);

    if (!reqCity || !dates) {
      outputColZ.push([""]);
      continue;
    }

    let nightCount = Math.round((dates.end - dates.start) / ONE_DAY_MS);
    if (nightCount < 5) {
      outputColZ.push([""]);
      continue;
    }

    // Verfügbare Wohnungen in der passenden Region ermitteln (inkl. Terminverschiebungen)
    let availableApts = getAvailableAptsForRequest(apartments, bookings, reqCity, dates, ONE_DAY_MS);

    // 4. Wohnungs-Matching durchführen (Einzeln vs. Kombi)
    let result = matchApartments(availableApts, reqPersons, dates, ONE_DAY_MS);
    outputColZ.push([result.links]);
  }

  // Ergebnisse zurückschreiben (Spalte Z = Index 26) nur für die verarbeiteten Zeilen
  if (outputColZ.length > 0) {
    s1.getRange(1, 26, outputColZ.length, 1).setValues(outputColZ);
  }

  Logger.log("\n--- SKRIPT ENDE ---");
  sendOfferEmail();
}

// ======================================================================
// LOGIK- UND SUCHANFRAGE-UNTERFUNKTIONEN
// ======================================================================

function parseApartments(mData) {
  let apartments = [];
  for (let i = 1; i < mData.length; i++) {
    // Spalte AJ (Index 35): Ausschluss-Flag
    let excludeFlag = String(mData[i][35] || "").trim().toLowerCase();
    if (excludeFlag === "x") continue;

    let aptName = String(mData[i][0]).trim();
    let slug = String(mData[i][6] || "").trim();
    let maxBeds = parseInt(mData[i][15]) || 0; // Spalte P (Index 15): Anzahl Betten
    let regionsString = mData[i][22];
    let minPersons = parseInt(mData[i][34]) || 1; // Spalte AI (Index 34): Minimum Personen

    let regionsList = [];
    if (regionsString && typeof regionsString === 'string') {
      regionsList = regionsString.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    }

    if (aptName) {
      apartments.push({
        name: aptName,
        slug: slug,
        regions: regionsList,
        minPersons: minPersons,
        maxBeds: maxBeds
      });
    }
  }
  return apartments;
}

function getAvailableAptsForRequest(apartments, bookings, reqCity, dates, ONE_DAY_MS) {
  let availableApts = [];
  for (let apt of apartments) {
    let matchRegion = apt.regions.some(r => reqCity.includes(r) || r.includes(reqCity));
    if (!matchRegion || !apt.slug) continue;

    let cleanSlug = apt.slug.startsWith('/') ? apt.slug : '/' + apt.slug;
    let baseLink = "https://a.l8street.com" + cleanSlug;
    if (!baseLink.endsWith('/')) baseLink += '/';

    let bestShiftDays = null;
    if (checkAvailability(dates.start, dates.end, bookings[apt.name])) {
      bestShiftDays = 0;
    } else {
      for (let offset = 1; offset <= 7; offset++) {
        for (let shift of [offset, -offset]) {
          let shiftMs = shift * ONE_DAY_MS;
          if (checkAvailability(dates.start + shiftMs, dates.end + shiftMs, bookings[apt.name])) {
            bestShiftDays = shift;
            break;
          }
        }
        if (bestShiftDays !== null) break;
      }
    }

    if (bestShiftDays !== null) {
      availableApts.push({
        name: apt.name,
        baseLink: baseLink,
        minPersons: apt.minPersons,
        maxBeds: apt.maxBeds,
        dayOffsetAbs: Math.abs(bestShiftDays),
        bestShiftDays: bestShiftDays
      });
    }
  }
  return availableApts;
}

function matchApartments(availableApts, reqPersons, dates, ONE_DAY_MS) {
  // 1. Einzelne passende Objekte suchen
  let singleMatches = availableApts.filter(a => reqPersons >= a.minPersons && reqPersons <= a.maxBeds);

  if (singleMatches.length > 0) {
    let candidateMatches = singleMatches.map(apt => {
      let notes = [];
      if (apt.bestShiftDays !== 0) {
        let altStartMs = dates.start + (apt.bestShiftDays * ONE_DAY_MS);
        notes.push("ab " + formatDateGerman(altStartMs));
      }
      let fullLink = apt.baseLink;
      if (notes.length > 0) fullLink += " (" + notes.join(", ") + ")";

      let unusedBeds = apt.maxBeds - reqPersons;
      return { link: fullLink, dayOffsetAbs: apt.dayOffsetAbs, unusedBeds: unusedBeds };
    });

    candidateMatches.sort((a, b) => {
      if (a.dayOffsetAbs !== b.dayOffsetAbs) return a.dayOffsetAbs - b.dayOffsetAbs;
      return a.unusedBeds - b.unusedBeds;
    });

    return {
      links: candidateMatches.map(item => item.link).join(', '),
      status: "Erfolgreich zugewiesen"
    };
  }

  // 2. Kombinations-Suche
  return findComboApartments(availableApts, reqPersons, dates, ONE_DAY_MS);
}

function findComboApartments(availableApts, reqPersons, dates, ONE_DAY_MS) {
  if (availableApts.length < 2) {
    return { links: "Keine verfügbar", status: "Keine freien Objekte" };
  }

  let comboMatches = [];
  for (let m = 0; m < availableApts.length; m++) {
    for (let n = m + 1; n < availableApts.length; n++) {
      let apt1 = availableApts[m];
      let apt2 = availableApts[n];

      let combinedMinPersons = apt1.minPersons + apt2.minPersons;
      let combinedMaxBeds = apt1.maxBeds + apt2.maxBeds;

      if (reqPersons >= combinedMinPersons && reqPersons <= combinedMaxBeds) {
        let maxOffset = Math.max(apt1.dayOffsetAbs, apt2.dayOffsetAbs);
        let unusedBeds = combinedMaxBeds - reqPersons;
        let notes = ["Kombi für " + combinedMaxBeds + " Pers."];

        if (apt1.bestShiftDays !== 0 || apt2.bestShiftDays !== 0) {
          let shiftUsed = apt1.bestShiftDays !== 0 ? apt1.bestShiftDays : apt2.bestShiftDays;
          let altStartMs = dates.start + (shiftUsed * ONE_DAY_MS);
          notes.push("ab " + formatDateGerman(altStartMs));
        }

        let comboLink = apt1.baseLink + " + " + apt2.baseLink + " (" + notes.join(", ") + ")";
        comboMatches.push({ link: comboLink, dayOffsetAbs: maxOffset, unusedBeds: unusedBeds });
      }
    }
  }

  if (comboMatches.length > 0) {
    comboMatches.sort((a, b) => {
      if (a.dayOffsetAbs !== b.dayOffsetAbs) return a.dayOffsetAbs - b.dayOffsetAbs;
      return a.unusedBeds - b.unusedBeds;
    });

    return {
      links: comboMatches.map(item => item.link).join(', '),
      status: "Kombination gewählt"
    };
  }

  return { links: "Keine verfügbar", status: "Keine passende Kombination" };
}

function parseBookings(calData) {
  let bookings = {};
  for (let i = 1; i < calData.length; i++) {
    let aptName = String(calData[i][0]).trim();
    let start = parseIsoDate(calData[i][1]);
    let end = parseIsoDate(calData[i][2]);

    if (aptName && start !== null && end !== null) {
      if (!bookings[aptName]) bookings[aptName] = [];
      bookings[aptName].push({ start: start, end: end });
    }
  }
  return bookings;
}

// ======================================================================
// DATUMS- UND HILFSFUNKTIONEN
// ======================================================================

function checkAvailability(startMs, endMs, aptBookings) {
  if (!aptBookings || aptBookings.length === 0) return true;
  for (let b of aptBookings) {
    if (startMs < b.end && endMs > b.start) return false;
  }
  return true;
}

function formatDateGerman(timeMs) {
  let d = new Date(timeMs);
  let months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
  return d.getUTCDate() + ". " + months[d.getUTCMonth()];
}

function parseGermanDateRange(dateStr) {
  try {
    let parts = String(dateStr).split('-');
    if (parts.length !== 2) return null;
    let d1 = parts[0].trim().split('.');
    let d2 = parts[1].trim().split('.');
    return { start: Date.UTC(d1[2], d1[1] - 1, d1[0]), end: Date.UTC(d2[2], d2[1] - 1, d2[0]) };
  } catch (e) {
    return null;
  }
}

function parseIsoDate(isoStr) {
  try {
    if (!isoStr) return null;
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