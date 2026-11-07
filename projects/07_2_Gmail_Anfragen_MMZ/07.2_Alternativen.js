// =================================================================
// STUFE 2b-Alt: NEXT-AVAILABLE KOMBI-FINDER FÜR TAB "MMZ"
// Liest Spalte D (Daten), Spalte E (Routing), Spalte G (Engine B)
// Schreibt Alternativ-Kombis AUSSCHLIESSLICH in Spalte I (Spalte 9)
// =================================================================
function runApartmentMatchingEngineB_Alternative_MMZ() {
  const SPREADSHEET_INVENTORY_ID = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';
  const SHEET_NAME = 'MMZ';

  const MAX_VERSCHIEBUNG_TAGE = 14;
  const GEWICHT_TAG_VERSCHIEBUNG = 15;
  const GEWICHT_LEERBETT = 2;
  const GEWICHT_LEERZIMMER = 3;
  const GEWICHT_APARTMENT_ANZAHL = 5;

  Logger.log("=== START ENGINE B-ALT (NEXT AVAILABLE KOMBI) ===");
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

  ensureEngineBAltHeaderMMZ(leadsSheet);

  const mData = invSheetM.getDataRange().getValues();
  const apartments = [];
  for (let r = 1; r < mData.length; r++) {
    const name = String(mData[r][0] || "").trim();
    if (!name) continue;
    const rawShortCode = String(mData[r][6] || "").trim();
    const shortCode = cleanShortCode(rawShortCode) || name;
    apartments.push({
      name: name,
      code: shortCode,
      schlafzimmer: Number(mData[r][14]) || 1,
      maxPersonen: Number(mData[r][15]) || 2,
      regionen: String(mData[r][22] || "").toLowerCase(),
      dauerhaftBelegt: String(mData[r][44] || "").trim().toLowerCase() === "x"
    });
  }

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

  for (const k in bookingsMap) {
    bookingsMap[k].sort((a, b) => a.startMs - b.startMs);
  }

  const numRows = lastRowLeads - 1;
  const leadsRange = leadsSheet.getRange(2, 1, numRows, 9);
  const leadsData = leadsRange.getValues();
  let updatesMade = false;

  for (let i = 0; i < numRows; i++) {
    const rowNum = i + 2;
    const cellB = String(leadsData[i][1] || "").trim();
    const cellD = String(leadsData[i][3] || "").trim();
    const cellE = String(leadsData[i][4] || "").trim();
    const cellG = String(leadsData[i][6] || "").trim(); // Engine B
    const cellI = String(leadsData[i][8] || "").trim(); // Zielspalte I

    if (!cellB || cellB === "-" || !cellD || !cellE || cellE.startsWith("FEHLER")) continue;
    if (cellG !== "[]") continue;
    if (cellI !== "" && cellI !== "[]") continue;

    let routingStatus = "";
    let kennzahlen = {};
    try {
      const routingObj = JSON.parse(cellE);
      routingStatus = routingObj.status || "";
      kennzahlen = routingObj.kennzahlen || {};
    } catch (e) {
      continue;
    }

    const istGrossGruppe = (routingStatus === "BIG_SHORT" || routingStatus === "BIG_LONG");
    const istStandard = (routingStatus === "STD_SHORT" || routingStatus === "STD_LONG");
    if (!istGrossGruppe && !istStandard) continue;

    let valData = {};
    try {
      valData = JSON.parse(cellD);
    } catch (e) {
      continue;
    }

    const persData = valData.personen || {};
    const dauerData = valData.dauer || {};
    const ortData = valData.ort || {};
    const anfragePers = Number(persData.Pers) || Number(kennzahlen.pers) || 0;
    const anfrageStart = String(dauerData.start || "").trim();
    let anfrageEnde = String(dauerData.ende || "").trim();

    if (anfrageStart && !anfrageEnde) {
      const dStart = new Date(anfrageStart + "T00:00:00Z");
      anfrageEnde = new Date(dStart.getTime() + 6 * 86400000).toISOString().split("T")[0];
    }

    const zielZimmer = Number(kennzahlen.zimmer) || Math.ceil(anfragePers / 2);
    const rohOrt = String(ortData.suchortText || ortData.Ort || "").trim();
    const suchStadt = extractPureCityName(rohOrt);
    if (!suchStadt || !anfrageStart || !anfrageEnde || anfragePers <= 1) continue;

    const wunschStartMs = new Date(anfrageStart + "T00:00:00Z").getTime();
    const wunschEndeMs = new Date(anfrageEnde + "T00:00:00Z").getTime();
    const aufenthaltDauerMs = wunschEndeMs - wunschStartMs;
    if (aufenthaltDauerMs <= 0) continue;

    const regionApts = apartments.filter(apt => {
      if (apt.dauerhaftBelegt) return false;
      return apt.regionen.includes(suchStadt) || suchStadt.includes(apt.regionen);
    });

    if (regionApts.length === 0) {
      leadsData[i][8] = "[]";
      updatesMade = true;
      continue;
    }

    const gefundeneKombiAlternativen = [];
    for (let shift = 1; shift <= MAX_VERSCHIEBUNG_TAGE; shift++) {
      const testStartMs = wunschStartMs + (shift * 86400000);
      const testEndeMs = testStartMs + aufenthaltDauerMs;

      const freieInDiesemSlot = [];
      for (const apt of regionApts) {
        const aptKey = apt.name.toLowerCase();
        const bookings = bookingsMap[aptKey] || [];
        const hatKollision = bookings.some(b => b.startMs < testEndeMs && b.endMs > testStartMs);
        if (!hatKollision) {
          freieInDiesemSlot.push({
            id: apt.code,
            maxP: apt.maxPersonen,
            sz: apt.schlafzimmer
          });
        }
      }

      const summeMoeglicheBetten = freieInDiesemSlot.reduce((sum, a) => sum + a.maxP, 0);
      if (summeMoeglicheBetten < anfragePers) continue;

      const kombis = findeApartmentKombinationenAlt(freieInDiesemSlot, anfragePers, zielZimmer);
      for (const k of kombis) {
        if (k.teilangebot) continue;
        const penaltyScore = (shift * GEWICHT_TAG_VERSCHIEBUNG) +
          (k.restB * GEWICHT_LEERBETT) +
          (k.restZimmer * GEWICHT_LEERZIMMER) +
          (k.anzahl * GEWICHT_APARTMENT_ANZAHL);

        gefundeneKombiAlternativen.push({
          apts: k.apts,
          betten: k.gesamtBetten,
          schlafzimmer: k.gesamtZimmer,
          restB: k.restB,
          restZ: k.restZimmer,
          anzahl: k.anzahl,
          istAlternative: true,
          tageVerschoben: shift,
          altStart: new Date(testStartMs).toISOString().split("T")[0],
          altEnde: new Date(testEndeMs).toISOString().split("T")[0],
          penaltyScore: penaltyScore
        });
      }

      if (gefundeneKombiAlternativen.length >= 4) break;
    }

    gefundeneKombiAlternativen.sort((a, b) => a.penaltyScore - b.penaltyScore);

    if (gefundeneKombiAlternativen.length > 0) {
      const exportListe = gefundeneKombiAlternativen.slice(0, 3);
      leadsData[i][8] = JSON.stringify(exportListe);
      updatesMade = true;
      Logger.log(`Zeile ${rowNum}: ${exportListe.length} Kombi-Alternative(n) in Spalte I hinterlegt (+${exportListe[0].tageVerschoben}d, Apts: ${exportListe[0].apts.join("+")}).`);
    } else {
      leadsData[i][8] = "[]";
      updatesMade = true;
      Logger.log(`Zeile ${rowNum}: Keine Kombi-Alternative innerhalb von ${MAX_VERSCHIEBUNG_TAGE} Tagen gefunden.`);
    }
  }

  if (updatesMade) {
    const colIValues = leadsData.map(row => [row[8]]);
    leadsSheet.getRange(2, 9, numRows, 1).setValues(colIValues);
    for (let r = 2; r <= lastRowLeads; r++) {
      leadsSheet.setRowHeight(r, 21);
    }
    Logger.log("=== KOMBI-ALTERNATIVEN ERFOLGREICH IN SPALTE I GESPEICHERT ===");
  }
}

function ensureEngineBAltHeaderMMZ(sheet) {
  const cell = sheet.getRange(1, 9); // Spalte I
  if (!cell.getValue() || cell.getValue() !== "Alternative Kombi (>6P)") {
    cell.setValue("Alternative Kombi (>6P)");
    cell.setFontWeight("bold");
    sheet.getRange("I:I").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN FÜR ENGINE B-ALT
// -----------------------------------------------------------------
function findeApartmentKombinationenAlt(apts, zielPersonen, zielZimmer) {
  const resultateVoll = [];
  const maxKombiGroesse = 4;
  const minZimmer = Number(zielZimmer) || 0;

  function backtrack(startIndex, aktuelleKombi, summeBetten, summeZimmer) {
    const bettenOk = summeBetten >= zielPersonen;
    const zimmerOk = minZimmer === 0 || summeZimmer >= minZimmer;

    if (bettenOk && zimmerOk) {
      resultateVoll.push({
        apts: aktuelleKombi.map(a => a.id),
        gesamtBetten: summeBetten,
        gesamtZimmer: summeZimmer,
        restB: summeBetten - zielPersonen,
        restZimmer: minZimmer > 0 ? (summeZimmer - minZimmer) : 0,
        anzahl: aktuelleKombi.length,
        teilangebot: false
      });
      return;
    }

    if (aktuelleKombi.length >= maxKombiGroesse) return;

    for (let j = startIndex; j < apts.length; j++) {
      aktuelleKombi.push(apts[j]);
      backtrack(j + 1, aktuelleKombi, summeBetten + apts[j].maxP, summeZimmer + apts[j].sz);
      aktuelleKombi.pop();
    }
  }

  backtrack(0, [], 0, 0);
  return resultateVoll;
}

function cleanShortCode(rawCode) {
  if (!rawCode) return "";
  return String(rawCode).trim().replace(/[\r\n\t]/g, "");
}

function parseIsoDateString(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "UTC", "yyyy-MM-dd");
  }
  const str = String(val).trim();
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(str)) {
    const parts = str.split(".");
    return `${parts[2]}-${parts[1].padStart(2, "0")}-${parts[0].padStart(2, "0")}`;
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.substring(0, 10);
  }
  return "";
}

function extractPureCityName(text) {
  if (!text) return "";
  let clean = String(text).toLowerCase().trim();
  clean = clean.replace(/\b\d{5}\b/g, "").trim();
  clean = clean.replace(/\b(nähe|raum|umkreis|zentrum|city|hbf|bahnhof|stadt)\b/gi, "").trim();
  if (clean.includes(",")) {
    clean = clean.split(",")[0].trim();
  }
  return clean.replace(/\s+/g, " ");
}