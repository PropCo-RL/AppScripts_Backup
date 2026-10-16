/**
 * MVP 2: Verfeinerte Sortierung, Preisberechnung und Zuweisung in Spalte AB
 */
function runMVP2_ApartmentRefinement() {
  const sheet1Id = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const sheet2Id = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';

  const s1 = SpreadsheetApp.openById(sheet1Id).getSheetByName('DMZ');
  const s2_m = SpreadsheetApp.openById(sheet2Id).getSheetByName('M');
  const s2_cal = SpreadsheetApp.openById(sheet2Id).getSheetByName('cal');

  const dmzData = s1.getDataRange().getValues();
  const mData = s2_m.getDataRange().getValues();
  const calData = s2_cal.getDataRange().getValues();

  const ONE_DAY_MS = 1000 * 60 * 60 * 24;
  const bookings = parseBookings(calData);
  const apartmentsMap = parseApartmentsDetailed(mData);

  let outputColAB = [["MVP 2 - Bestes Angebot & Preis"]];

  // DMZ Anfragen durchgehen
  for (let i = 1; i < dmzData.length; i++) {
    let colA = String(dmzData[i][0] || "").trim();
    if (colA === "") break; // Stopp bei erster leerer Zeile in Spalte A

    let existingZ = String(dmzData[i][25] || "").trim(); // Spalte Z (MVP 1 Links)
    let colS = String(dmzData[i][18] || "").trim(); // Spalte S
    let colT = String(dmzData[i][19] || "").trim(); // Spalte T

    // Überspringen, wenn Z leer ist ODER wenn S UND T bereits befüllt sind
    if (existingZ === "" || (colS !== "" && colT !== "")) {
      outputColAB.push([""]);
      continue;
    }

    let dateStr = dmzData[i][14] || dmzData[i][4]; // Spalte O oder E
    let reqPersons = parseInt(dmzData[i][13]) || parseInt(dmzData[i][5]) || 1; // Spalte N/F
    let roomSetupReq = String(dmzData[i][22] || "").trim(); // Spalte W (Zimmeraufteilung)
    let extensionReq = String(dmzData[i][23] || "").trim(); // Spalte X (Mögliche Verlängerung)

    let dates = parseGermanDateRange(dateStr);
    if (!dates) {
      outputColAB.push(["Fehler: Datum ungültig"]);
      continue;
    }

    let nightCount = Math.round((dates.end - dates.start) / ONE_DAY_MS);
    let candidateSlugs = extractSlugsFromLinks(existingZ);
    let evaluatedCandidates = [];

    for (let slug of candidateSlugs) {
      let apt = apartmentsMap[slug];
      if (!apt) continue;

      // 1. Lücken-Score (Anschluss)
      let gapInfo = calculateGapScore(apt.name, dates, bookings[apt.name], ONE_DAY_MS);

      // 2. Zimmer-Match (Spalte W vs. O/P)
      let roomMatchInfo = checkRoomRequirements(apt, reqPersons, roomSetupReq);

      // 3. Verlängerung prüfen (Spalte X)
      let extensionPossible = checkExtensionAvailability(apt.name, dates, extensionReq, bookings[apt.name]);

      // 4. Qualitäts-Grade (Spalte AR: A > B > C)
      let qualityScore = (apt.qualityGrade === "a") ? 1 : (apt.qualityGrade === "b") ? 2 : 3;

      evaluatedCandidates.push({
        apt: apt,
        gapScore: gapInfo.score,
        gapDescription: gapInfo.description,
        roomMatch: roomMatchInfo.isMatch,
        extensionPossible: extensionPossible,
        qualityScore: qualityScore
      });
    }

    if (evaluatedCandidates.length === 0) {
      outputColAB.push(["Keine passenden Optionen"]);
      continue;
    }

    // MULTI-KRITERIEN SORTIERUNG (Prioritäten-Kette)
    evaluatedCandidates.sort((a, b) => {
      // Prio 1: Anschluss / Lücke
      if (a.gapScore !== b.gapScore) return a.gapScore - b.gapScore;
      // Prio 2: Passt zur Zimmeraufteilung
      if (a.roomMatch !== b.roomMatch) return b.roomMatch - a.roomMatch;
      // Prio 3: Optionale Verlängerung frei
      if (a.extensionPossible !== b.extensionPossible) return b.extensionPossible - a.extensionPossible;
      // Prio 4: Qualität Standard (A vor B vor C)
      return a.qualityScore - b.qualityScore;
    });

    // Besten Kandidaten auswählen & Preis berechnen
    let best = evaluatedCandidates[0];
    let pricePerNight = best.apt.netPricePerNight;

    // Preisberechnung: (Nächte * Tagespreis) + (1 * Tagespreis Endreinigung)
    let totalPrice = (nightCount * pricePerNight) + pricePerNight;
    let offerText = best.apt.baseLink + " | " +
                    best.gapDescription + " | " +
                    nightCount + " Nächte x " + pricePerNight + "€ + " + pricePerNight + "€ Endreinigung = Gesamt netto: " + totalPrice + "€";

    outputColAB.push([offerText]);
  }

  // Ergebnisse in Spalte AB (28) schreiben
  s1.getRange(1, 28, outputColAB.length, 1).setValues(outputColAB);
  Logger.log("--- MVP 2 ABGESCHLOSSEN ---");
}

// ======================================================================
// HILFSFUNKTIONEN
// ======================================================================

function parseApartmentsDetailed(mData) {
  let map = {};
  for (let i = 1; i < mData.length; i++) {
    let slug = String(mData[i][6] || "").trim().toLowerCase();
    if (!slug) continue;

    // Führenden Slash sicherstellen
    let cleanSlug = slug.startsWith('/') ? slug : '/' + slug;

    // Nachlaufenden Slash entfernen für einheitlichen Abgleich (z. B. /a/l8)
    if (cleanSlug.endsWith('/') && cleanSlug.length > 1) {
      cleanSlug = cleanSlug.slice(0, -1);
    }

    map[cleanSlug] = {
      name: String(mData[i][0]).trim(),
      slug: cleanSlug,
      baseLink: "https://a.l8street.com" + cleanSlug + "/",
      bedrooms: parseInt(mData[i][14]) || 1, // Spalte O (Index 14)
      maxBeds: parseInt(mData[i][15]) || 0, // Spalte P (Index 15)
      netPricePerNight: parseFloat(mData[i][17]) || 0, // Spalte R (Index 17): Preis/Nacht netto
      qualityGrade: String(mData[i][43] || "").trim().toLowerCase() // Spalte AR (Index 43): Standard A/B/C
    };
  }
  return map;
}

function calculateGapScore(aptName, reqDates, aptBookings, ONE_DAY_MS) {
  if (!aptBookings || aptBookings.length === 0) {
    return { score: 5, description: "Keine Vorbuchung" };
  }

  let minGap = 999;
  for (let b of aptBookings) {
    let gapBefore = Math.abs(reqDates.start - b.end) / ONE_DAY_MS;
    let gapAfter = Math.abs(b.start - reqDates.end) / ONE_DAY_MS;
    let currentMin = Math.min(gapBefore, gapAfter);

    if (currentMin < minGap) {
      minGap = currentMin;
    }
  }

  if (minGap === 0) return { score: 0, description: "Nahtloser Anschluss" };
  if (minGap === 1) return { score: 1, description: "1 Tag Lücke" };
  return { score: Math.round(minGap), description: Math.round(minGap) + " Tage Lücke" };
}

function checkRoomRequirements(apt, reqPersons, roomSetupReq) {
  let requiredBedrooms = Math.ceil(reqPersons / 2);

  if (roomSetupReq !== "") {
    let singleRooms = (roomSetupReq.match(/einzel/gi) || []).length;
    let doubleRooms = (roomSetupReq.match(/doppel/gi) || []).length;
    if (singleRooms > 0 || doubleRooms > 0) {
      requiredBedrooms = singleRooms + doubleRooms;
    }
  }

  return {
    isMatch: (apt.bedrooms >= requiredBedrooms && apt.maxBeds >= reqPersons)
  };
}

function checkExtensionAvailability(aptName, reqDates, extensionReq, aptBookings) {
  if (!extensionReq || !aptBookings) return true;

  let extDates = parseGermanDateRange(extensionReq);
  if (!extDates) return true;

  // Prüfen, ob der Zeitraum von Buchungsende bis Verlängerungsende frei ist
  for (let b of aptBookings) {
    if (reqDates.end < b.end && extDates.end > b.start) {
      return false; // Überschneidung
    }
  }

  return true;
}

function extractSlugsFromLinks(linksStr) {
  // Erfasst Links unabhängig von nachfolgendem Text wie (ab 28. Okt)
  let urls = linksStr.match(/https?:\/\/a\.l8street\.com\/[^\s,\)]+/g);
  if (!urls) return [];

  return urls.map(url => {
    let path = url.replace("https://a.l8street.com", "").toLowerCase().trim();
    // Nachlaufenden Slash entfernen
    if (path.endsWith('/') && path.length > 1) {
      path = path.slice(0, -1);
    }
    return path;
  });
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

function parseGermanDateRange(dateStr) {
  try {
    let parts = String(dateStr).split('-');
    if (parts.length !== 2) return null;

    let d1 = parts[0].trim().split('.');
    let d2 = parts[1].trim().split('.');

    return { 
      start: Date.UTC(d1[2], d1[1] - 1, d1[0]), 
      end: Date.UTC(d2[2], d2[1] - 1, d2[0]) 
    };
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