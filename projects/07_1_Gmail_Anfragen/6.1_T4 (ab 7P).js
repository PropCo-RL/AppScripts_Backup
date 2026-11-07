// =================================================================
// STUFE 2b: ENGINE B - MULTI-APARTMENT FINDER (Liest O, schreibt in Q)
// Für Großanfragen ab 7 Personen (BIG_SHORT & BIG_LONG)
// Strategie Lücken-Priorisierung:
// - Echte Anschlüsse werden nur bis max. 14 Tage Leerstand gewertet
// - Anreise <= 5 Tage: Vollkommen freie Einheiten (Score -1) bevorzugen
// - Anreise > 5 Tage:  Nahtlose Anschlüsse bevorzugen, freie Einheiten zurückstellen (Score 999)
// =================================================================
function runApartmentMatchingEngineB() {
  const SPREADSHEET_LEADS_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const SPREADSHEET_INVENTORY_ID = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';

  Logger.log("=== START ENGINE B MULTI-MATCHING (BIG_SHORT & BIG_LONG) ===");
  let leadsSheet, invSheetM, invSheetCal;
  try {
    const leadsSS = SpreadsheetApp.openById(SPREADSHEET_LEADS_ID);
    leadsSheet = leadsSS.getSheetByName('DMZ2');
    const invSS = SpreadsheetApp.openById(SPREADSHEET_INVENTORY_ID);
    invSheetM = invSS.getSheetByName('M');
    invSheetCal = invSS.getSheetByName('cal');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabellen: " + e.toString());
    return;
  }

  if (!leadsSheet || !invSheetM || !invSheetCal) {
    Logger.log("FEHLER: Ein Tabellenblatt fehlt (DMZ2, M oder cal)!");
    return;
  }

  const lastRowLeads = leadsSheet.getLastRow();
  if (lastRowLeads <= 1) return;

  // Header für Spalte Q (Spalte 17) sicherstellen
  const headerCell = leadsSheet.getRange(1, 17);
  if (!headerCell.getValue()) {
    headerCell.setValue("Optionen Kombi (>6P)");
    headerCell.setFontWeight("bold");
    leadsSheet.getRange("Q:Q").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // 1. Apartments aus Tab 'M' einlesen
  const mData = invSheetM.getDataRange().getValues();
  const apartments = [];
  for (let r = 1; r < mData.length; r++) {
    const name = String(mData[r][0] || "").trim(); // Spalte A: Name (für cal)
    if (!name) continue;

    const rawShortCode = String(mData[r][6] || "").trim(); // Spalte G: Kürzel
    const shortCode = cleanShortCode(rawShortCode) || name;

    apartments.push({
      name: name,
      code: shortCode,
      schlafzimmer: Number(mData[r][14]) || 0, // Spalte O
      maxPersonen: Number(mData[r][15]) || 0,  // Spalte P
      regionen: String(mData[r][22] || "").toLowerCase(), // Spalte W
      dauerhaftBelegt: String(mData[r][44] || "").trim().toLowerCase() === "x" // Spalte AS
    });
  }
  Logger.log(`Geladene Apartments aus Tab M: ${apartments.length}`);

  // 2. Buchungen aus Tab 'cal' einlesen
  const calData = invSheetCal.getDataRange().getValues();
  const bookingsMap = {};
  for (let c = 1; c < calData.length; c++) {
    const propName = String(calData[c][0] || "").trim().toLowerCase();
    const startStr = parseIsoDateString(calData[c][1]);
    const endStr = parseIsoDateString(calData[c][2]);
    if (!propName || !startStr || !endStr) continue;

    if (!bookingsMap[propName]) bookingsMap[propName] = [];
    bookingsMap[propName].push({
      startStr: startStr,
      endStr: endStr,
      startMs: new Date(startStr + "T00:00:00Z").getTime(),
      endMs: new Date(endStr + "T00:00:00Z").getTime()
    });
  }

  // 3. Leads aus DMZ2 einlesen (Spalten A bis Q = 17 Spalten)
  const numRows = lastRowLeads - 1;
  const leadsRange = leadsSheet.getRange(2, 1, numRows, 17);
  const leadsData = leadsRange.getValues();
  let updatesMade = false;

  // Heutiges Datum auf 00:00:00 UTC normiert
  const jetzt = new Date();
  const heuteMs = new Date(Date.UTC(jetzt.getFullYear(), jetzt.getMonth(), jetzt.getDate())).getTime();

  for (let i = 0; i < leadsData.length; i++) {
    const rowNum = i + 2;
    const routingStatus = String(leadsData[i][14] || "").trim(); // Spalte O (Index 14)
    const cellQ = String(leadsData[i][16] || "").trim();         // Spalte Q (Index 16)

    // Nur Zeilen mit Status "BIG_SHORT" oder "BIG_LONG" bearbeiten
    if ((routingStatus !== "BIG_SHORT" && routingStatus !== "BIG_LONG") || cellQ !== "") {
      continue;
    }

    // JSON-Daten auslesen: E=Ort, F=Pers, G=Dauer
    let ortData = {}, persData = {}, dauerData = {};
    try {
      ortData = JSON.parse(leadsData[i][4] || "{}");
      persData = JSON.parse(leadsData[i][5] || "{}");
      dauerData = JSON.parse(leadsData[i][6] || "{}");
    } catch (e) {
      Logger.log(`Zeile ${rowNum}: JSON Parse Fehler in Spalten E/F/G`);
      continue;
    }

    const anfragePers = Number(persData.Pers) || 0;
    const anfrageStart = String(dauerData.start || "").trim();
    const anfrageEnde = String(dauerData.ende || "").trim();

    // Freitext-Zielort vor Portalort
    const primarStadt = String(ortData.suchortText || "").toLowerCase().trim();
    const basisStadt = String(ortData.Ort || "").toLowerCase().trim();

    if ((!primarStadt && !basisStadt) || !anfrageStart || !anfrageEnde || anfragePers <= 6) {
      continue;
    }

    const anfrageStartMs = new Date(anfrageStart + "T00:00:00Z").getTime();
    const anfrageEndeMs = new Date(anfrageEnde + "T00:00:00Z").getTime();
    const tageBisAnreise = Math.round((anfrageStartMs - heuteMs) / 86400000);

    // Schritt A: Alle freien Apartments in der Region sammeln
    const verfuegbareApts = [];

    for (const apt of apartments) {
      let regionMatch = false;
      if (primarStadt && apt.regionen.includes(primarStadt)) {
        regionMatch = true;
      } else if (basisStadt && apt.regionen.includes(basisStadt)) {
        regionMatch = true;
      }
      if (!regionMatch) continue;
      if (apt.dauerhaftBelegt) continue;

      const aptKey = apt.name.toLowerCase();
      const existingBookings = bookingsMap[aptKey] || [];
      let isBlocked = false;
      let letzteBuchungEndeMs = null;

      for (const booking of existingBookings) {
        if (booking.startMs < anfrageEndeMs && booking.endMs > anfrageStartMs) {
          isBlocked = true;
          break;
        }
        if (booking.endMs <= anfrageStartMs) {
          if (letzteBuchungEndeMs === null || booking.endMs > letzteBuchungEndeMs) {
            letzteBuchungEndeMs = booking.endMs;
          }
        }
      }

      if (isBlocked) continue;

      // 14-Tage-Deckel für echte Lücken
      let leerTageVorher = null;
      if (letzteBuchungEndeMs !== null) {
        const tageDiff = Math.max(0, Math.round((anfrageStartMs - letzteBuchungEndeMs) / 86400000));
        if (tageDiff <= 14) {
          leerTageVorher = tageDiff;
        }
      }

      // Lücken-Score abhängig von Vorlaufzeit
      let lueckeScore;
      if (tageBisAnreise <= 5) {
        lueckeScore = (leerTageVorher === null) ? -1 : leerTageVorher;
      } else {
        lueckeScore = (leerTageVorher === null) ? 999 : leerTageVorher;
      }

      verfuegbareApts.push({
        id: apt.code,
        maxP: apt.maxPersonen,
        sz: apt.schlafzimmer,
        score: lueckeScore
      });
    }

    Logger.log(`Zeile ${rowNum} (${routingStatus}, ${anfragePers}P, in ${tageBisAnreise}d): ${verfuegbareApts.length} Einheiten grundsätzlich frei.`);

    // Schritt B: Kombinationen berechnen (2 bis 4 Apartments)
    const gefundeneKombis = findeApartmentKombinationen(verfuegbareApts, anfragePers);

    // Schritt C: Nach Wirtschaftlichkeit sortieren
    gefundeneKombis.sort((a, b) => {
      // 1. Wenigste Restbetten verschwenden
      if (a.restB !== b.restB) return a.restB - b.restB;
      // 2. Weniger Wohnungen bevorzugen (weniger Verwaltungsaufwand)
      if (a.anzahl !== b.anzahl) return a.anzahl - b.anzahl;
      // 3. Beste kombinierte Lücken-Bewertung (niedrigster Gesamt-Score)
      return a.summeScore - b.summeScore;
    });

    if (gefundeneKombis.length > 0) {
      // Beschränke auf die Top 3 Vorschläge
      const topKombis = gefundeneKombis.slice(0, 3).map(k => ({
        apts: k.apts,
        betten: k.gesamtBetten,
        restB: k.restB,
        anzahl: k.anzahl
      }));
      leadsData[i][16] = JSON.stringify(topKombis);
      updatesMade = true;
      Logger.log(`Zeile ${rowNum}: ${topKombis.length} Kombi-Option(en) in Spalte Q eingetragen.`);
    } else {
      leadsData[i][16] = "[]";
      updatesMade = true;
      Logger.log(`Zeile ${rowNum}: Keine passende Apartment-Kombination gefunden.`);
    }
  }

  if (updatesMade) {
    leadsRange.setValues(leadsData);
    for (let r = 2; r <= lastRowLeads; r++) {
      leadsSheet.setRowHeight(r, 21);
    }
    Logger.log("=== ENGINE B ERFOLGREICH IN SPALTE Q GESCHRIEBEN ===");
  }
}

// -----------------------------------------------------------------
// KOMBINATIONS-ALGORITHMUS (Branch & Bound bis max. 4 Einheiten)
// -----------------------------------------------------------------
function findeApartmentKombinationen(apts, zielPersonen) {
  const resultate = [];
  const maxKombiGroesse = 4; // Maximal 4 Einheiten bündeln

  function backtrack(startIndex, aktuelleKombi, summeBetten, summeScore) {
    if (summeBetten >= zielPersonen) {
      resultate.push({
        apts: aktuelleKombi.map(a => a.id),
        gesamtBetten: summeBetten,
        restB: summeBetten - zielPersonen,
        anzahl: aktuelleKombi.length,
        summeScore: summeScore
      });
      return;
    }

    if (aktuelleKombi.length >= maxKombiGroesse) return;

    for (let j = startIndex; j < apts.length; j++) {
      aktuelleKombi.push(apts[j]);
      backtrack(
        j + 1,
        aktuelleKombi,
        summeBetten + apts[j].maxP,
        summeScore + apts[j].score
      );
      aktuelleKombi.pop();
    }
  }

  backtrack(0, [], 0, 0);
  return resultate;
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN
// -----------------------------------------------------------------
function cleanShortCode(val) {
  if (!val) return "";
  return String(val)
    .replace(/^\/?a\//i, "")
    .replace(/^\//, "")
    .trim();
}

function parseIsoDateString(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT", "yyyy-MM-dd");
  }
  const str = String(val).trim();
  const m = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) {
    const y = m[1];
    const month = m[2].padStart(2, '0');
    const day = m[3].padStart(2, '0');
    return `${y}-${month}-${day}`;
  }
  return "";
}