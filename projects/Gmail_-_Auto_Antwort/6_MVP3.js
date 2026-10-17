/**
 * MVP 3: Verfeinerte Sortierung mit SST-Risiko-Check (RE & VL Tabs) -> Spalte AB
 */
function runMVP3_ApartmentRefinementWithRisk() {
  const sheet1Id = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const sheet2Id = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';
  const sheet3Id = '1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU'; // SST Sheet

  const s1 = SpreadsheetApp.openById(sheet1Id).getSheetByName('DMZ');
  const s2_m = SpreadsheetApp.openById(sheet2Id).getSheetByName('M');
  const s2_cal = SpreadsheetApp.openById(sheet2Id).getSheetByName('cal');
  const s3_re = SpreadsheetApp.openById(sheet3Id).getSheetByName('RE');
  const s3_vl = SpreadsheetApp.openById(sheet3Id).getSheetByName('VL');

  const dmzData = s1.getDataRange().getValues();
  const mData = s2_m.getDataRange().getValues();
  const calData = s2_cal.getDataRange().getValues();
  const reData = s3_re.getDataRange().getValues();
  const vlData = s3_vl.getDataRange().getValues();

  const ONE_DAY_MS = 1000 * 60 * 60 * 24;

  const bookings = parseBookings(calData);
  const apartmentsMap = parseApartmentsDetailed(mData);
  const sstRiskMap = parseSSTRiskData(reData, vlData); // SST-Risikodaten analysieren

  let outputColAB = [["MVP 3 - Bestes Angebot (inkl. Risiko-Check)"]];

  // DMZ Anfragen durchgehen
  for (let i = 1; i < dmzData.length; i++) {
    let colA = String(dmzData[i][0] || "").trim();
    if (colA === "") break; // Stopp bei erster leerer Zeile in Spalte A

    let existingZ = String(dmzData[i][25] || "").trim();  // Spalte Z (MVP 1 Links, Index 25)
    let existingAB = String(dmzData[i][27] || "").trim(); // Spalte AB (MVP 3 Ergebnis, Index 27)
    let colS = String(dmzData[i][18] || "").trim();        // Spalte S (Index 18)
    let colT = String(dmzData[i][19] || "").trim();        // Spalte T (Index 19)

    // 1. NEU: Wenn in Spalte AB bereits etwas steht -> Wert behalten und NICHT neu berechnen
    if (existingAB !== "") {
      outputColAB.push([dmzData[i][27]]);
      continue;
    }

    // 2. Überspringen, wenn Z leer ist ODER wenn S UND T bereits befüllt sind
    if (existingZ === "" || (colS !== "" && colT !== "")) {
      outputColAB.push([""]);
      continue;
    }

    let dateStr = dmzData[i][14] || dmzData[i][4]; // Spalte O oder E (DD.MM.YYYY-DD.MM.YYYY)
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

      // PFLICHT-CHECK: Reicht das Einzel-Apartment für Betten & Zimmeraufteilung?
      let roomMatchInfo = checkRoomRequirements(apt, reqPersons, roomSetupReq);
      if (!roomMatchInfo.isMatch) {
        continue;
      }

      // 0. RISIKO-CHECK (MVP 3 Neu): Ist der aktuelle Vormieter ein Verlängerungs-Kandidat?
      let isRisky = checkSSTRisk(slug, dates, sstRiskMap);

      // 1. Verlängerung prüfen (Spalte X)
      let extensionPossible = checkExtensionAvailability(apt.name, dates, extensionReq, bookings[apt.name]);

      // 2. Lücken-Score (Anschluss)
      let gapInfo = calculateGapScore(apt.name, dates, bookings[apt.name], ONE_DAY_MS);

      // 3. Qualitäts-Grade (Spalte AR: A > B > C)
      let qualityScore = (apt.qualityGrade === "a") ? 1 : (apt.qualityGrade === "b") ? 2 : 3;

      evaluatedCandidates.push({
        apt: apt,
        isRisky: isRisky,
        extensionPossible: extensionPossible,
        gapScore: gapInfo.score,
        gapDescription: gapInfo.description,
        qualityScore: qualityScore
      });
    }

    if (evaluatedCandidates.length === 0) {
      outputColAB.push(["Keine passenden Einzel-Optionen"]);
      continue;
    }

    // MULTI-KRITERIEN SORTIERUNG INKL. RISIKO (MVP 3)
    evaluatedCandidates.sort((a, b) => {
      // Prio 0: Risikofreie Wohnungen bevorzugen (isRisky: false = 0, true = 1)
      if (a.isRisky !== b.isRisky) return (a.isRisky ? 1 : 0) - (b.isRisky ? 1 : 0);

      // Prio 1: Optionale Verlängerung für den neuen Kunden frei
      if (a.extensionPossible !== b.extensionPossible) return b.extensionPossible - a.extensionPossible;

      // Prio 2: Bester Kalender-Anschluss (Nahtlos > Lücke)
      if (a.gapScore !== b.gapScore) return a.gapScore - b.gapScore;

      // Prio 3: Qualität Standard (A vor B vor C)
      return a.qualityScore - b.qualityScore;
    });

    // Besten Einzel-Kandidaten auswählen & Preis berechnen
    let best = evaluatedCandidates[0];
    let pricePerNight = best.apt.netPricePerNight;
    let totalPrice = (nightCount * pricePerNight) + pricePerNight;

    let riskTag = best.isRisky ? " [Achtung: Verlängerungs-Risiko Vormieter]" : "";
    let offerText = best.apt.baseLink + riskTag + " | " + 
                    best.gapDescription + " | " + 
                    nightCount + " Nächte x " + pricePerNight + "€ + " + pricePerNight + "€ Endreinigung = Gesamt netto: " + totalPrice + "€";

    outputColAB.push([offerText]);
  }

  // Ergebnisse in Spalte AB (28) schreiben
  s1.getRange(1, 28, outputColAB.length, 1).setValues(outputColAB);
  Logger.log("--- MVP 3 ABGESCHLOSSEN (Spalte AB befüllt) ---");
}

// ======================================================================
// NEUE SST-RISIKO HILFSFUNKTIONEN (MVP 3)
// ======================================================================

function parseSSTRiskData(reData, vlData) {
  let riskMap = {};
  
  // 1. RE Tab durchgehen (Index 13 = Spalte N für Slug, Index 10/11 = Spalte K/L für Daten, Index 37 = Spalte AL)
  parseSSTRows(reData, riskMap, false);
  
  // 2. VL Tab durchgehen (Spalte N für Slug, Spalte AB/AC für Verlängerungszeitraum, Spalte AL)
  parseSSTRows(vlData, riskMap, true);
  
  return riskMap;
}

function parseSSTRows(data, riskMap, isVLTab) {
  for (let i = 1; i < data.length; i++) {
    let rawSlug = String(data[i][13] || "").trim().toLowerCase(); // Spalte N (Slug)
    if (!rawSlug) continue;

    let cleanSlug = rawSlug.startsWith('/') ? rawSlug : '/' + rawSlug;
    if (cleanSlug.endsWith('/') && cleanSlug.length > 1) {
      cleanSlug = cleanSlug.slice(0, -1);
    }

    // Datumsspalten unterscheiden: bei VL nehmen wir AB/AC (Index 27/28), sonst K/L (Index 10/11)
    let startDateStr = isVLTab ? data[i][27] : data[i][10];
    let endDateStr = isVLTab ? data[i][28] : data[i][11];
    let colAL = String(data[i][37] || "").trim(); // Spalte AL (Definitiver Auszug/Storno)

    let start = parseGermanSingleDate(startDateStr);
    let end = parseGermanSingleDate(endDateStr);

    if (start && end) {
      if (!riskMap[cleanSlug]) riskMap[cleanSlug] = [];
      riskMap[cleanSlug].push({
        start: start,
        end: end,
        hasExtendedBefore: isVLTab, // true, wenn der Eintrag im VL Tab steht
        isDefiniteExit: (colAL !== "") // true, wenn Spalte AL befüllt ist
      });
    }
  }
}

function checkSSTRisk(slug, reqDates, sstRiskMap) {
  let entries = sstRiskMap[slug];
  if (!entries || entries.length === 0) return false;

  // Suche nach Buchungen des Vormieters direkt vor der neuen Anfrage
  for (let b of entries) {
    // Wenn die Buchung des Vormieters im Tab VL liegt UND Spalte AL leer ist -> Risiko!
    if (b.hasExtendedBefore && !b.isDefiniteExit) {
      // Prüfen, ob das Ende dieser Buchung nah am Start der neuen Anfrage liegt (z.B. innerhalb von 14 Tagen)
      let diffDays = (reqDates.start - b.end) / (1000 * 60 * 60 * 24);
      if (diffDays >= -1 && diffDays <= 14) {
        return true; // Risikoreich!
      }
    }
  }
  return false; // Sicher
}

function parseGermanSingleDate(dateVal) {
  try {
    if (!dateVal) return null;
    if (dateVal instanceof Date) {
      return Date.UTC(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate());
    }
    let parts = String(dateVal).trim().split('.');
    if (parts.length !== 3) return null;
    return Date.UTC(parts[2], parts[1] - 1, parts[0]);
  } catch (e) {
    return null;
  }
}

// ======================================================================
// BESTEHENDE HILFSFUNKTIONEN
// ======================================================================

function parseApartmentsDetailed(mData) {
  let map = {};
  for (let i = 1; i < mData.length; i++) {
    let slug = String(mData[i][6] || "").trim().toLowerCase();
    if (!slug) continue;

    let cleanSlug = slug.startsWith('/') ? slug : '/' + slug;
    if (cleanSlug.endsWith('/') && cleanSlug.length > 1) {
      cleanSlug = cleanSlug.slice(0, -1);
    }

    map[cleanSlug] = {
      name: String(mData[i][0]).trim(),
      slug: cleanSlug,
      baseLink: "https://a.l8street.com" + cleanSlug + "/",
      bedrooms: parseInt(mData[i][14]) || 1,
      maxBeds: parseInt(mData[i][15]) || 0,
      netPricePerNight: parseFloat(mData[i][17]) || 0,
      qualityGrade: String(mData[i][43] || "").trim().toLowerCase()
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

  for (let b of aptBookings) {
    if (reqDates.end < b.end && extDates.end > b.start) {
      return false;
    }
  }
  return true;
}

function extractSlugsFromLinks(linksStr) {
  let urls = linksStr.match(/https?:\/\/a\.l8street\.com\/[^\s,\)]+/g);
  if (!urls) return [];

  return urls.map(url => {
    let path = url.replace("https://a.l8street.com", "").toLowerCase().trim();
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