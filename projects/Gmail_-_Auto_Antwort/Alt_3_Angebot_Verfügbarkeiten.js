function findAvailableApartments() {
  const sheet1Id = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const sheet2Id = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';
  const s1 = SpreadsheetApp.openById(sheet1Id).getSheetByName('DMZ');
  const s2_m = SpreadsheetApp.openById(sheet2Id).getSheetByName('M');
  const s2_cal = SpreadsheetApp.openById(sheet2Id).getSheetByName('cal');
  const dmzData = s1.getDataRange().getValues();
  const mData = s2_m.getDataRange().getValues();
  const calData = s2_cal.getDataRange().getValues();
  Logger.log("--- SKRIPT START (Liest Gemini-Spalten -> Schreibt in AB/AC) ---");
  
  // 1. Alle gebuchten Zeiträume aus "cal" einlesen
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
  
  // 2. Alle Wohnungen aus "M" einlesen
  let apartments = [];
  for (let i = 1; i < mData.length; i++) {
    let excludeFlag = String(mData[i][35] || "").trim().toLowerCase(); 
    if (excludeFlag === "x") continue;
    let aptName = String(mData[i][0]).trim();
    let slug = String(mData[i][6] || "").trim(); 
    let regionsString = mData[i][22]; 
    let minPersons = parseInt(mData[i][34]) || 1; 
    let regionsList = [];
    if (regionsString && typeof regionsString === 'string') {
      regionsList = regionsString.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    }
    if (aptName) {
      apartments.push({
        name: aptName,
        slug: slug,
        regions: regionsList,
        minPersons: minPersons
      });
    }
  }
  
  let outputColAB = [["Verfügbare Links"]];
  let outputColAC = [["Status / Hinweis"]];
  const ONE_DAY_MS = 1000 * 60 * 60 * 24;
  
  for (let i = 1; i < dmzData.length; i++) {
    // WICHTIGE ÄNDERUNG: Wenn in Spalte AE (Index 30) schon eine Mail gesendet wurde, Zeile überspringen und alte Werte übernehmen!
    let emailStatus = String(dmzData[i][30] || "").trim(); 
    if (emailStatus !== "") {
      outputColAB.push([dmzData[i][27] || ""]); // AB behalten
      outputColAC.push([dmzData[i][28] || ""]); // AC behalten
      continue;
    }

    let dateStr = dmzData[i][14] || dmzData[i][4]; 
    let reqPersons = parseInt(dmzData[i][13]) || parseInt(dmzData[i][5]) || 1; 
    let addressStr = dmzData[i][12] || dmzData[i][9]; 
    if (!dateStr || !addressStr) {
      outputColAB.push([""]);
      outputColAC.push(["Keine Daten vorhanden"]);
      continue;
    }
    
    let cityMatch = String(addressStr).match(/\d{4,5}\s+([a-zA-ZäöüÄÖÜß\-\s]+)/);
    let reqCity = cityMatch ? cityMatch[1].trim().toLowerCase() : String(addressStr).trim().toLowerCase();
    let dates = parseGermanDateRange(dateStr);
    if (!reqCity || !dates) {
      outputColAB.push([""]);
      outputColAC.push(["Fehler: Stadt oder Datum unlesbar"]);
      continue;
    }
    
    let nightCount = Math.round((dates.end - dates.start) / ONE_DAY_MS);
    if (nightCount < 5) {
      outputColAB.push([""]);
      outputColAC.push(["Unter Mindestaufenthalt (min. 5 Tage)"]);
      continue;
    }
    
    let availableApts = [];
    for (let apt of apartments) {
      let matchRegion = apt.regions.some(r => reqCity.includes(r) || r.includes(reqCity));
      if (matchRegion && apt.slug) {
        let cleanSlug = apt.slug.startsWith('/') ? apt.slug : '/' + apt.slug;
        let baseLink = "https://a.l8street.com" + cleanSlug;
        if (!baseLink.endsWith('/')) baseLink += '/';
        let bestShiftDays = null;
        if (checkAvailability(dates.start, dates.end, bookings[apt.name])) {
          bestShiftDays = 0;
        } else {
          for (let offset = 1; offset <= 7; offset++) {
            let candidateOffsets = [offset, -offset];
            for (let shift of candidateOffsets) {
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
            dayOffsetAbs: Math.abs(bestShiftDays),
            bestShiftDays: bestShiftDays
          });
        }
      }
    }
    
    let exactCapacityMatches = availableApts.filter(a => reqPersons <= a.minPersons);
    if (exactCapacityMatches.length > 0) {
      let candidateMatches = availableApts.map(apt => {
        let personDiff = apt.minPersons - reqPersons;
        let notes = [];
        if (apt.bestShiftDays !== 0) {
          let altStartMs = dates.start + (apt.bestShiftDays * ONE_DAY_MS);
          notes.push("ab " + formatDateGerman(altStartMs));
        }
        if (personDiff > 0) {
          notes.push("ab " + apt.minPersons + " Pers.");
        }
        let fullLink = apt.baseLink;
        if (notes.length > 0) fullLink += " (" + notes.join(", ") + ")";
        return { link: fullLink, dayOffsetAbs: apt.dayOffsetAbs, personDiff: personDiff >= 0 ? personDiff : 99 };
      });
      candidateMatches.sort((a, b) => {
        if (a.dayOffsetAbs !== b.dayOffsetAbs) return a.dayOffsetAbs - b.dayOffsetAbs;
        return a.personDiff - b.personDiff;
      });
      outputColAB.push([candidateMatches.map(item => item.link).join(', ')]);
      outputColAC.push(["Erfolgreich zugewiesen"]);
    } else if (availableApts.length >= 2) {
      let comboMatches = [];
      for (let m = 0; m < availableApts.length; m++) {
        for (let n = m + 1; n < availableApts.length; n++) {
          let apt1 = availableApts[m];
          let apt2 = availableApts[n];
          let totalCapacity = apt1.minPersons + apt2.minPersons;
          if (totalCapacity >= reqPersons) {
            let maxOffset = Math.max(apt1.dayOffsetAbs, apt2.dayOffsetAbs);
            let personDiff = totalCapacity - reqPersons;
            let notes = ["Kombi für " + totalCapacity + " Pers."];
            if (apt1.bestShiftDays !== 0 || apt2.bestShiftDays !== 0) {
              let shiftUsed = apt1.bestShiftDays !== 0 ? apt1.bestShiftDays : apt2.bestShiftDays;
              let altStartMs = dates.start + (shiftUsed * ONE_DAY_MS);
              notes.push("ab " + formatDateGerman(altStartMs));
            }
            let comboLink = apt1.baseLink + " + " + apt2.baseLink + " (" + notes.join(", ") + ")";
            comboMatches.push({ link: comboLink, dayOffsetAbs: maxOffset, personDiff: personDiff });
          }
        }
      }
      if (comboMatches.length > 0) {
        comboMatches.sort((a, b) => {
          if (a.dayOffsetAbs !== b.dayOffsetAbs) return a.dayOffsetAbs - b.dayOffsetAbs;
          return a.personDiff - b.personDiff;
        });
        outputColAB.push([comboMatches.map(item => item.link).join(', ')]);
        outputColAC.push(["Kombination gewählt"]);
      } else {
        outputColAB.push(["Keine verfügbar"]);
        outputColAC.push(["Keine passende Kombination"]);
      }
    } else {
      outputColAB.push(["Keine verfügbar"]);
      outputColAC.push(["Keine freien Objekte"]);
    }
  }
  
  // Schreibvorgänge in Spalte AB (28) und AC (29)
  s1.getRange(1, 28, outputColAB.length, 1).setValues(outputColAB);
  s1.getRange(1, 29, outputColAC.length, 1).setValues(outputColAC);
  Logger.log("\n--- SKRIPT ENDE ---");
  sendOfferEmail();
}

// ----------------------------------------------------------------------
// HILFSFUNKTIONEN FÜR SKRIPT 2
// ----------------------------------------------------------------------
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
  } catch (e) { return null; }
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
  } catch (e) { return null; }
}