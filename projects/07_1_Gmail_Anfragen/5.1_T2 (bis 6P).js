// =================================================================
// STUFE 2: ENGINE A - SINGLE APARTMENT FINDER (Liest O, schreibt in P)
// Strategie Lücken-Priorisierung:
// - Echte Anschlüsse werden nur bis max. 14 Tage Leerstand gewertet.
// - Anreise <= 5 Tage: Vollkommen frei (null / >14 Tage) gewinnt (Score -1)
// - Anreise > 5 Tage:  Nahtloser Anschluss (0 Tage) gewinnt, vollkommen frei wird zurückgestellt (Score 999)
// =================================================================
function runApartmentMatchingEngineA() {
  const SPREADSHEET_LEADS_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const SPREADSHEET_INVENTORY_ID = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';

  Logger.log("=== START ENGINE A MATCHING (STD_SHORT & STD_LONG) ===");
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

  // Header für Spalte P (Spalte 16) sicherstellen
  const headerCell = leadsSheet.getRange(1, 16);
  if (!headerCell.getValue()) {
    headerCell.setValue("Optionen (<=6P)");
    headerCell.setFontWeight("bold");
    leadsSheet.getRange("P:P").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // 1. Apartments aus Tab 'M' einlesen
  const mData = invSheetM.getDataRange().getValues();
  const apartments = [];
  for (let r = 1; r < mData.length; r++) {
    const name = String(mData[r][0] || "").trim(); // Spalte A: Name (für cal-Abgleich)
    if (!name) continue;

    // Kürzel aus Spalte G (Index 6) holen und "/a/" bereinigen
    const rawShortCode = String(mData[r][6] || "").trim();
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

  // 3. Leads aus DMZ2 einlesen (Spalten A bis P = 16 Spalten)
  const numRows = lastRowLeads - 1;
  const leadsRange = leadsSheet.getRange(2, 1, numRows, 16);
  const leadsData = leadsRange.getValues();
  let updatesMade = false;

  // Heutiges Datum auf 00:00:00 UTC genormt (für Vorlaufzeit-Berechnung)
  const jetzt = new Date();
  const heuteMs = new Date(Date.UTC(jetzt.getFullYear(), jetzt.getMonth(), jetzt.getDate())).getTime();

  for (let i = 0; i < leadsData.length; i++) {
    const rowNum = i + 2;
    const routingStatus = String(leadsData[i][14] || "").trim(); // Spalte O (Index 14)
    const cellP = String(leadsData[i][15] || "").trim();         // Spalte P (Index 15)

    // Nur Zeilen mit Status "STD_SHORT" oder "STD_LONG" bearbeiten, falls Spalte P leer ist
    if ((routingStatus !== "STD_SHORT" && routingStatus !== "STD_LONG") || cellP !== "") {
      continue;
    }

    // JSON-Daten parsen: E=Ort, F=Pers, G=Dauer
    let ortData = {}, persData = {}, dauerData = {};
    try {
      ortData = JSON.parse(leadsData[i][4] || "{}"); // Spalte E
      persData = JSON.parse(leadsData[i][5] || "{}"); // Spalte F
      dauerData = JSON.parse(leadsData[i][6] || "{}"); // Spalte G
    } catch (e) {
      Logger.log(`Zeile ${rowNum}: JSON Parse Fehler in Spalten E/F/G`);
      continue;
    }

    const anfragePers = Number(persData.Pers) || 0;
    const suchStadt = String(ortData.Ort || "").toLowerCase().trim();
    const anfrageStart = String(dauerData.start || "").trim();
    const anfrageEnde = String(dauerData.ende || "").trim();

    if (!suchStadt || !anfrageStart || !anfrageEnde) {
      continue;
    }

    const anfrageStartMs = new Date(anfrageStart + "T00:00:00Z").getTime();
    const anfrageEndeMs = new Date(anfrageEnde + "T00:00:00Z").getTime();

    // Vorlaufzeit in Tagen berechnen
    const tageBisAnreise = Math.round((anfrageStartMs - heuteMs) / 86400000);

    const freieApartments = [];

    for (const apt of apartments) {
      // 1. Regionen-Abgleich gegen Spalte W
      if (!apt.regionen.includes(suchStadt)) continue;

      // 2. Ausschluss über Spalte AS (Dauerhaft belegt)
      if (apt.dauerhaftBelegt) continue;

      // 3. Kapazität (Spalte P)
      if (apt.maxPersonen < anfragePers) continue;

      // 4. Kalenderabgleich gegen Tab 'cal'
      const aptKey = apt.name.toLowerCase();
      const existingBookings = bookingsMap[aptKey] || [];
      let isBlocked = false;
      let letzteBuchungEndeMs = null;

      for (const booking of existingBookings) {
        // Kollision mit bestehender Buchung
        if (booking.startMs < anfrageEndeMs && booking.endMs > anfrageStartMs) {
          isBlocked = true;
          break;
        }

        // Vorherige Buchungsabreise ermitteln
        if (booking.endMs <= anfrageStartMs) {
          if (letzteBuchungEndeMs === null || booking.endMs > letzteBuchungEndeMs) {
            letzteBuchungEndeMs = booking.endMs;
          }
        }
      }

      if (isBlocked) continue;

      const restBetten = apt.maxPersonen - anfragePers;

      // Leerstandstage berechnen & Altbuchungen deckeln:
      // Liegt der letzte Auszug mehr als 14 Tage zurück, gilt das Apartment als komplett frei (null)
      let leerTageVorher = null;
      if (letzteBuchungEndeMs !== null) {
        const tageDiff = Math.max(0, Math.round((anfrageStartMs - letzteBuchungEndeMs) / 86400000));
        if (tageDiff <= 14) {
          leerTageVorher = tageDiff;
        }
      }

      // Lücken-Score abhängig von Vorlaufzeit:
      // Niedrigerer Score = Bessere Platzierung
      let lueckeScore;
      if (tageBisAnreise <= 5) {
        // KURZFRISTIG (<= 5 Tage): Vollkommen freies Apartment gewinnt (Score -1),
        // danach direkte Anschlüsse (0, 1, 2...)
        lueckeScore = (leerTageVorher === null) ? -1 : leerTageVorher;
      } else {
        // ZUKUNFT (> 5 Tage): Nahtloser Anschluss (0 Tage) gewinnt,
        // komplett freie Einheiten werden nach hinten gestellt (Score 999), um sie für Sofortbuchungen frei zu halten
        lueckeScore = (leerTageVorher === null) ? 999 : leerTageVorher;
      }

      freieApartments.push({
        id: apt.code,
        restB: restBetten,
        lueckeScore: lueckeScore,
        lueckeDisplay: leerTageVorher
      });
    }

    // Sortierung:
    // 1. Passgenauigkeit (Wenigste ungenutzte Betten zuerst)
    // 2. Lücken-Score (Abhängig von Vorlaufzeit)
    freieApartments.sort((a, b) => {
      if (a.restB !== b.restB) {
        return a.restB - b.restB;
      }
      return a.lueckeScore - b.lueckeScore;
    });

    if (freieApartments.length > 0) {
      const exportListe = freieApartments.map(item => ({
        id: item.id,
        restB: item.restB,
        lueckeT: item.lueckeDisplay
      }));
      leadsData[i][15] = JSON.stringify(exportListe);
      updatesMade = true;
      Logger.log(`Zeile ${rowNum} (${routingStatus}, Anreise in ${tageBisAnreise}d): ${exportListe.length} Apartment(s) gematcht.`);
    } else {
      leadsData[i][15] = "[]";
      updatesMade = true;
      Logger.log(`Zeile ${rowNum} (${routingStatus}): Keine freien Apartments gefunden.`);
    }
  }

  if (updatesMade) {
    leadsRange.setValues(leadsData);
    for (let r = 2; r <= lastRowLeads; r++) {
      leadsSheet.setRowHeight(r, 21);
    }
    Logger.log("=== MATCHING ERFOLGREICH IN SPALTE P GESCHRIEBEN ===");
  }
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