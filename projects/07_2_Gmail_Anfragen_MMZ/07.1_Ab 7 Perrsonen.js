// =================================================================
// STUFE 2b: ENGINE B - MULTI-APARTMENT FINDER FÜR TAB "MMZ"
// Liest Spalte D (Daten), Spalte E (Routing) & Spalte F (Engine A)
// Schreibt Kombinations-Optionen in Spalte G
// =================================================================
function runApartmentMatchingEngineB_MMZ() {
  const SPREADSHEET_INVENTORY_ID = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';
  const SHEET_NAME = 'MMZ';

  Logger.log("=== START ENGINE B MULTI-MATCHING FÜR MMZ ===");

  const ssLeads = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ssLeads.getSheetByName(SHEET_NAME);
  let invSheetM, invSheetCal;

  try {
    const invSS = SpreadsheetApp.openById(SPREADSHEET_INVENTORY_ID);
    invSheetM = invSS.getSheetByName('M');
    invSheetCal = invSS.getSheetByName('cal');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen des Inventory-Spreadsheets: " + e.toString());
    return;
  }

  if (!leadsSheet || !invSheetM || !invSheetCal) {
    Logger.log("FEHLER: Ein Tabellenblatt fehlt (MMZ, M oder cal)!");
    return;
  }

  const lastRowLeads = leadsSheet.getLastRow();
  if (lastRowLeads <= 1) return;

  // Header für Spalte G sicherstellen
  ensureEngineBHeaderMMZ(leadsSheet);

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
      schlafzimmer: Number(mData[r][14]) || 1, // Spalte O (Fallback mind. 1)
      maxPersonen: Number(mData[r][15]) || 2, // Spalte P (Fallback mind. 2)
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

  // 3. Leads aus MMZ einlesen (Spalten A bis G = 7 Spalten)
  // Index: A=0, B=1, C=2, D=3 (Gesamt), E=4 (Routing), F=5 (Single), G=6 (Kombi)
  const numRows = lastRowLeads - 1;
  const leadsRange = leadsSheet.getRange(2, 1, numRows, 7);
  const leadsData = leadsRange.getValues();
  let updatesMade = false;

  const jetzt = new Date();
  const heuteMs = new Date(Date.UTC(jetzt.getFullYear(), jetzt.getMonth(), jetzt.getDate())).getTime();

  for (let i = 0; i < leadsData.length; i++) {
    const rowNum = i + 2;
    const cellB = String(leadsData[i][1] || "").trim(); // Spalte B: Unique
    const cellD = String(leadsData[i][3] || "").trim(); // Spalte D: Gesamt-JSON
    const cellE = String(leadsData[i][4] || "").trim(); // Spalte E: Routing-JSON
    const cellF = String(leadsData[i][5] || "").trim(); // Spalte F: Single-Ergebnis (Engine A)
    const cellG = String(leadsData[i][6] || "").trim(); // Spalte G: Kombi-Ausgabe

    // Überspringen wenn: Duplikat "-", Vorstufen fehlen
    if (!cellB || cellB === "-" || !cellD || !cellE || cellE.startsWith("FEHLER")) {
      continue;
    }

    // Bereits befüllt (außer bei Neudurchläufen mit leerem '[]')
    if (cellG !== "" && cellG !== "[]") {
      continue;
    }

    // Routing-Status & Kennzahlen aus Spalte E lesen
    let routingStatus = "";
    let kennzahlen = {};
    try {
      const routingObj = JSON.parse(cellE);
      routingStatus = routingObj.status || "";
      kennzahlen = routingObj.kennzahlen || {};
    } catch (e) {
      continue;
    }

    // Kaskaden-Prüfung: Wann muss Engine B suchen?
    const istGrossGruppe = (routingStatus === "BIG_SHORT" || routingStatus === "BIG_LONG");
    const istStandardOhneSingleTreffer = (routingStatus === "STD_SHORT" || routingStatus === "STD_LONG") && (cellF === "[]" || cellF === "");

    if (!istGrossGruppe && !istStandardOhneSingleTreffer) {
      continue;
    }

    // Daten aus dem Gesamt-JSON in Spalte D auslesen
    let valData = {};
    try {
      valData = JSON.parse(cellD);
    } catch (e) {
      Logger.log(`Zeile ${rowNum}: JSON Parse Fehler in Spalte D`);
      continue;
    }

    const ortData = valData.ort || {};
    const persData = valData.personen || {};
    const dauerData = valData.dauer || {};

    const anfragePers = Number(persData.Pers) || Number(kennzahlen.pers) || 0;

    // Bevorzugt validierte Datumsangaben aus Stufe 1 verwenden
    const anfrageStart = String(kennzahlen.start || dauerData.start || "").trim();
    let anfrageEnde = String(kennzahlen.ende || dauerData.ende || "").trim();

    // Fallback: Falls noch immer kein Ende vorliegt, 6 Tage Mindestzeitraum erzwingen
    if (anfrageStart && !anfrageEnde) {
      const dStart = new Date(anfrageStart + "T00:00:00Z");
      const dEnd = new Date(dStart.getTime() + 6 * 86400000);
      anfrageEnde = dEnd.toISOString().split("T")[0];
    }

    const zielZimmer = Number(kennzahlen.zimmer) || Math.ceil(anfragePers / 2);

    // Reinen Stadtnamen isolieren
    const rohOrt = String(ortData.suchortText || ortData.Ort || "").trim();
    const suchStadt = extractPureCityName(rohOrt);

    if (!suchStadt || !anfrageStart || !anfrageEnde || anfragePers <= 1) {
      continue;
    }

    const anfrageStartMs = new Date(anfrageStart + "T00:00:00Z").getTime();
    const anfrageEndeMs = new Date(anfrageEnde + "T00:00:00Z").getTime();
    const tageBisAnreise = Math.round((anfrageStartMs - heuteMs) / 86400000);

    // Schritt A: Alle freien Einheiten in der Region sammeln
    const verfuegbareApts = [];
    for (const apt of apartments) {
      const aptRegion = apt.regionen;
      const match = aptRegion.includes(suchStadt) || suchStadt.includes(aptRegion);
      if (!match) continue;

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

      // 14-Tage-Deckel für Lücken
      let leerTageVorher = null;
      if (letzteBuchungEndeMs !== null) {
        const tageDiff = Math.max(0, Math.round((anfrageStartMs - letzteBuchungEndeMs) / 86400000));
        if (tageDiff <= 14) {
          leerTageVorher = tageDiff;
        }
      }

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

    Logger.log(`Zeile ${rowNum} (${routingStatus}, ${anfragePers}P / ${zielZimmer}Z, Zeitraum: ${anfrageStart} bis ${anfrageEnde}): ${verfuegbareApts.length} Einheiten in ${suchStadt} frei.`);

    // Schritt B: Kombinationen berechnen (inklusive Maximalkapazitäts-Fallback)
    const gefundeneKombis = findeApartmentKombinationen(verfuegbareApts, anfragePers, zielZimmer);

    // Schritt C: Nach Wirtschaftlichkeit sortieren
    gefundeneKombis.sort((a, b) => {
      // 1. Vollangebote immer vor Teilangeboten
      if (a.teilangebot !== b.teilangebot) {
        return a.teilangebot ? 1 : -1;
      }
      // 2. Bei Teilangeboten: Maximale Personen zuerst
      if (a.teilangebot && b.teilangebot && a.gesamtBetten !== b.gesamtBetten) {
        return b.gesamtBetten - a.gesamtBetten;
      }
      // 3. Wenigste Rest-Zimmer verschwenden
      if (a.restZimmer !== b.restZimmer) return a.restZimmer - b.restZimmer;
      // 4. Wenigste Rest-Betten verschwenden
      if (a.restB !== b.restB) return a.restB - b.restB;
      // 5. Weniger Wohnungen bevorzugen
      if (a.anzahl !== b.anzahl) return a.anzahl - b.anzahl;
      // 6. Niedrigster kombinierter Lücken-Score
      return a.summeScore - b.summeScore;
    });

    if (gefundeneKombis.length > 0) {
      const topKombis = gefundeneKombis.slice(0, 3).map(k => ({
        apts: k.apts,
        betten: k.gesamtBetten,
        schlafzimmer: k.gesamtZimmer,
        restB: k.restB,
        restZ: k.restZimmer,
        anzahl: k.anzahl,
        teilangebot: Boolean(k.teilangebot),
        fehlendePersonen: k.fehlendePersonen || 0
      }));

      leadsData[i][6] = JSON.stringify(topKombis);
      updatesMade = true;
      Logger.log(`Zeile ${rowNum}: ${topKombis.length} Kombi-Option(en) (Teilangebot: ${topKombis[0].teilangebot}) eingetragen.`);
    } else {
      leadsData[i][6] = "[]";
      updatesMade = true;
      Logger.log(`Zeile ${rowNum}: Keine passende Apartment-Kombination gefunden.`);
    }
  }

  // Zurückschreiben in Spalten A:G
  if (updatesMade) {
    leadsRange.setValues(leadsData);
    for (let r = 2; r <= lastRowLeads; r++) {
      leadsSheet.setRowHeight(r, 21);
    }
    Logger.log("=== ENGINE B ERFOLGREICH IN SPALTE G GESCHRIEBEN ===");
  }
}

// -----------------------------------------------------------------
// KOMBINATIONS-ALGORITHMUS (Vollabdeckung + maximales Teilangebot)
// -----------------------------------------------------------------
function findeApartmentKombinationen(apts, zielPersonen, zielZimmer) {
  const resultateVoll = [];
  let besteTeilKombi = null;
  const maxKombiGroesse = 4;
  const minZimmer = Number(zielZimmer) || 0;

  function backtrack(startIndex, aktuelleKombi, summeBetten, summeZimmer, summeScore) {
    const bettenOk = summeBetten >= zielPersonen;
    const zimmerOk = minZimmer === 0 || summeZimmer >= minZimmer;

    // Voll-Treffer gefunden
    if (bettenOk && zimmerOk) {
      resultateVoll.push({
        apts: aktuelleKombi.map(a => a.id),
        gesamtBetten: summeBetten,
        gesamtZimmer: summeZimmer,
        restB: summeBetten - zielPersonen,
        restZimmer: minZimmer > 0 ? (summeZimmer - minZimmer) : 0,
        anzahl: aktuelleKombi.length,
        summeScore: summeScore,
        teilangebot: false,
        fehlendePersonen: 0
      });
      return;
    }

    // Teilangebot erfassen (maximale Betten/Zimmer ohne Vollabdeckung)
    if (aktuelleKombi.length > 0) {
      if (!besteTeilKombi ||
          summeBetten > besteTeilKombi.gesamtBetten ||
          (summeBetten === besteTeilKombi.gesamtBetten && summeZimmer > besteTeilKombi.gesamtZimmer) ||
          (summeBetten === besteTeilKombi.gesamtBetten && summeZimmer === besteTeilKombi.gesamtZimmer && summeScore < besteTeilKombi.summeScore)) {
        besteTeilKombi = {
          apts: aktuelleKombi.map(a => a.id),
          gesamtBetten: summeBetten,
          gesamtZimmer: summeZimmer,
          restB: 0,
          restZimmer: 0,
          anzahl: aktuelleKombi.length,
          summeScore: summeScore,
          teilangebot: true,
          fehlendePersonen: Math.max(0, zielPersonen - summeBetten)
        };
      }
    }

    if (aktuelleKombi.length >= maxKombiGroesse) return;

    for (let j = startIndex; j < apts.length; j++) {
      aktuelleKombi.push(apts[j]);
      backtrack(
        j + 1,
        aktuelleKombi,
        summeBetten + apts[j].maxP,
        summeZimmer + apts[j].sz,
        summeScore + apts[j].score
      );
      aktuelleKombi.pop();
    }
  }

  backtrack(0, [], 0, 0, 0);

  // Wenn Voll-Treffer existieren, werden diese zurückgegeben
  if (resultateVoll.length > 0) {
    return resultateVoll;
  }

  // Fallback: Wenn keine Vollkombination möglich ist, aber Betten frei sind (mind. 2 Betten)
  if (besteTeilKombi && besteTeilKombi.gesamtBetten >= 2) {
    return [besteTeilKombi];
  }

  return [];
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN
// -----------------------------------------------------------------
function ensureEngineBHeaderMMZ(sheet) {
  const cell = sheet.getRange(1, 7); // Spalte G
  if (!cell.getValue()) {
    cell.setValue("Optionen Kombi (>6P)");
    cell.setFontWeight("bold");
    sheet.getRange("G:G").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}