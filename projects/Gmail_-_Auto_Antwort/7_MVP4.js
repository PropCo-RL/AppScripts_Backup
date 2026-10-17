/**
 * MVP 4: Kombi-Angebote (Multi-Apartment Fallback) -> Spalte AC
 */
function runMVP4_ApartmentCombinations() {
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
  const sstRiskMap = parseSSTRiskData(reData, vlData);

  let outputColAC = [["MVP 4 - Kombi-Angebot & Preis"]];

  // DMZ Anfragen durchgehen
  for (let i = 1; i < dmzData.length; i++) {
    let colA = String(dmzData[i][0] || "").trim();
    if (colA === "") break; // Stopp bei erster leerer Zeile in Spalte A

    let existingZ = String(dmzData[i][25] || "").trim();  // Spalte Z (MVP 1 Links, Index 25)
    let existingAC = String(dmzData[i][28] || "").trim(); // Spalte AC (MVP 4 Ergebnis, Index 28)
    let colS = String(dmzData[i][18] || "").trim();        // Spalte S (Index 18)
    let colT = String(dmzData[i][19] || "").trim();        // Spalte T (Index 19)

    // 1. NEU: Wenn in Spalte AC bereits etwas steht -> Wert behalten und NICHT neu berechnen
    if (existingAC !== "") {
      outputColAC.push([dmzData[i][28]]);
      continue;
    }

    // 2. Überspringen, wenn Z leer ist ODER wenn S UND T bereits befüllt sind
    if (existingZ === "" || (colS !== "" && colT !== "")) {
      outputColAC.push([""]);
      continue;
    }

    let dateStr = dmzData[i][14] || dmzData[i][4]; // Spalte O oder E
    let reqPersons = parseInt(dmzData[i][13]) || parseInt(dmzData[i][5]) || 1; // Spalte N/F
    let roomSetupReq = String(dmzData[i][22] || "").trim(); // Spalte W (Zimmeraufteilung)
    let extensionReq = String(dmzData[i][23] || "").trim(); // Spalte X (Mögliche Verlängerung)

    let dates = parseGermanDateRange(dateStr);
    if (!dates) {
      outputColAC.push(["Fehler: Datum ungültig"]);
      continue;
    }

    let nightCount = Math.round((dates.end - dates.start) / ONE_DAY_MS);
    let candidateSlugs = extractSlugsFromLinks(existingZ);

    // Bilde alle möglichen 2er-Kombinationen aus Spalte Z
    let evaluatedCombinations = [];

    for (let j = 0; j < candidateSlugs.length; j++) {
      for (let k = j + 1; k < candidateSlugs.length; k++) {
        let apt1 = apartmentsMap[candidateSlugs[j]];
        let apt2 = apartmentsMap[candidateSlugs[k]];

        if (!apt1 || !apt2) continue;

        // PFLICHT-CHECK für das PAAR: Reichen beide Wohnungen zusammen?
        let combinedBedrooms = apt1.bedrooms + apt2.bedrooms;
        let combinedMaxBeds = apt1.maxBeds + apt2.maxBeds;

        let requiredBedrooms = Math.ceil(reqPersons / 2);
        if (roomSetupReq !== "") {
          let singleRooms = (roomSetupReq.match(/einzel/gi) || []).length;
          let doubleRooms = (roomSetupReq.match(/doppel/gi) || []).length;
          if (singleRooms > 0 || doubleRooms > 0) {
            requiredBedrooms = singleRooms + doubleRooms;
          }
        }

        if (combinedBedrooms < requiredBedrooms || combinedMaxBeds < reqPersons) {
          continue; // Paar reicht zusammen nicht aus
        }

        // Bewerten von Apt 1 & Apt 2
        let risk1 = checkSSTRisk(apt1.slug, dates, sstRiskMap);
        let risk2 = checkSSTRisk(apt2.slug, dates, sstRiskMap);
        let isRiskyCombined = risk1 || risk2; // Wenn eine Wohnung riskant ist, hat das Paar ein Risiko

        let ext1 = checkExtensionAvailability(apt1.name, dates, extensionReq, bookings[apt1.name]);
        let ext2 = checkExtensionAvailability(apt2.name, dates, extensionReq, bookings[apt2.name]);
        let extensionPossibleCombined = ext1 && ext2;

        let gap1 = calculateGapScore(apt1.name, dates, bookings[apt1.name], ONE_DAY_MS);
        let gap2 = calculateGapScore(apt2.name, dates, bookings[apt2.name], ONE_DAY_MS);
        let avgGapScore = (gap1.score + gap2.score) / 2;

        let q1 = (apt1.qualityGrade === "a") ? 1 : (apt1.qualityGrade === "b") ? 2 : 3;
        let q2 = (apt2.qualityGrade === "a") ? 1 : (apt2.qualityGrade === "b") ? 2 : 3;
        let avgQuality = (q1 + q2) / 2;

        evaluatedCombinations.push({
          apt1: apt1,
          apt2: apt2,
          isRisky: isRiskyCombined,
          extensionPossible: extensionPossibleCombined,
          avgGapScore: avgGapScore,
          avgQuality: avgQuality
        });
      }
    }

    if (evaluatedCombinations.length === 0) {
      outputColAC.push(["Keine passenden Kombi-Optionen"]);
      continue;
    }

    // MULTI-KRITERIEN SORTIERUNG FÜR KOMBI-PAARE
    evaluatedCombinations.sort((a, b) => {
      // Prio 0: Risikofreie Paare bevorzugen
      if (a.isRisky !== b.isRisky) return (a.isRisky ? 1 : 0) - (b.isRisky ? 1 : 0);

      // Prio 1: Verlängerung für beide frei
      if (a.extensionPossible !== b.extensionPossible) return b.extensionPossible - a.extensionPossible;

      // Prio 2: Bester durchschnittlicher Kalender-Anschluss
      if (a.avgGapScore !== b.avgGapScore) return a.avgGapScore - b.avgGapScore;

      // Prio 3: Beste durchschnittliche Qualität
      return a.avgQuality - b.avgQuality;
    });

    // Bestes Kombi-Paar auswählen & Preis berechnen
    let best = evaluatedCombinations[0];
    let combinedNightPrice = best.apt1.netPricePerNight + best.apt2.netPricePerNight;
    let combinedCleaningFee = best.apt1.netPricePerNight + best.apt2.netPricePerNight; // 2x Endreinigung (je 1 Tagespreis)
    let totalPrice = (nightCount * combinedNightPrice) + combinedCleaningFee;

    let riskTag = best.isRisky ? " [Achtung: Vormieter-Risiko]" : "";
    let offerText = "KOMBI: " + best.apt1.baseLink + " + " + best.apt2.baseLink + riskTag + " | " +
                    nightCount + " Nächte x (" + best.apt1.netPricePerNight + "€ + " + best.apt2.netPricePerNight + "€) + " +
                    combinedCleaningFee + "€ Endreinigung (2x) = Gesamt netto: " + totalPrice + "€";

    outputColAC.push([offerText]);
  }

  // Ergebnisse in Spalte AC (Spalte 29) schreiben
  s1.getRange(1, 29, outputColAC.length, 1).setValues(outputColAC);
  Logger.log("--- MVP 4 ABGESCHLOSSEN (Spalte AC befüllt) ---");
}

// ======================================================================
// HILFSFUNKTIONEN (Gleich wie bei MVP 3)
// ======================================================================

function parseSSTRiskData(reData, vlData) {
  let riskMap = {};
  parseSSTRows(reData, riskMap, false);
  parseSSTRows(vlData, riskMap, true);
  return riskMap;
}

function parseSSTRows(data, riskMap, isVLTab) {
  for (let i = 1; i < data.length; i++) {
    let rawSlug = String(data[i][13] || "").trim().toLowerCase();
    if (!rawSlug) continue;

    let cleanSlug = rawSlug.startsWith('/') ? rawSlug : '/' + rawSlug;
    if (cleanSlug.endsWith('/') && cleanSlug.length > 1) {
      cleanSlug = cleanSlug.slice(0, -1);
    }

    let startDateStr = isVLTab ? data[i][27] : data[i][10];
    let endDateStr = isVLTab ? data[i][28] : data[i][11];
    let colAL = String(data[i][37] || "").trim();

    let start = parseGermanSingleDate(startDateStr);
    let end = parseGermanSingleDate(endDateStr);

    if (start && end) {
      if (!riskMap[cleanSlug]) riskMap[cleanSlug] = [];
      riskMap[cleanSlug].push({
        start: start,
        end: end,
        hasExtendedBefore: isVLTab,
        isDefiniteExit: (colAL !== "")
      });
    }
  }
}

function checkSSTRisk(slug, reqDates, sstRiskMap) {
  let entries = sstRiskMap[slug];
  if (!entries || entries.length === 0) return false;

  for (let b of entries) {
    if (b.hasExtendedBefore && !b.isDefiniteExit) {
      let diffDays = (reqDates.start - b.end) / (1000 * 60 * 60 * 24);
      if (diffDays >= -1 && diffDays <= 14) {
        return true;
      }
    }
  }
  return false;
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