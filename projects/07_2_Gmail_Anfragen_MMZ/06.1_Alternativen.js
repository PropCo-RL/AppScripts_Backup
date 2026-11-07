// =================================================================
// STUFE 2a-Alt: NEXT-AVAILABLE ALTERNATIVE FINDER FÜR TAB "MMZ"
// Liest Spalte D (Daten), Spalte E (Routing) & Spalte F (Engine A)
// Schreibt Alternativ-Optionen AUSSCHLIESSLICH in Spalte H (Spalte 8)
// =================================================================
function runApartmentMatchingEngineA_Alternative_MMZ() {
  const SPREADSHEET_INVENTORY_ID = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';
  const SHEET_NAME = 'MMZ';

  const MAX_VERSCHIEBUNG_TAGE = 14;
  const MAX_LEERBETTEN = 3;
  const GEWICHT_TAG_VERSCHIEBUNG = 10;
  const GEWICHT_LEERBETT = 3;
  const GEWICHT_LEERZIMMER = 4;

  Logger.log("=== START ENGINE A-ALT (NEXT AVAILABLE APARTMENT) ===");
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

  ensureEngineAAltHeaderMMZ(leadsSheet);

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
  const leadsRange = leadsSheet.getRange(2, 1, numRows, 8);
  const leadsData = leadsRange.getValues();
  let updatesMade = false;

  for (let i = 0; i < numRows; i++) {
    const rowNum = i + 2;
    const cellB = String(leadsData[i][1] || "").trim();
    const cellD = String(leadsData[i][3] || "").trim();
    const cellE = String(leadsData[i][4] || "").trim();
    const cellF = String(leadsData[i][5] || "").trim(); // Engine A
    const cellH = String(leadsData[i][7] || "").trim(); // Zielspalte H

    if (!cellB || cellB === "-" || !cellD || !cellE || cellE.startsWith("FEHLER")) continue;
    if (cellF !== "[]") continue;
    if (cellH !== "" && cellH !== "[]") continue;

    let routingStatus = "";
    let kennzahlen = {};
    try {
      const routingObj = JSON.parse(cellE);
      routingStatus = routingObj.status || "";
      kennzahlen = routingObj.kennzahlen || {};
    } catch (e) {
      continue;
    }

    if (routingStatus !== "STD_SHORT" && routingStatus !== "STD_LONG") continue;

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

    const kandidaten = [];
    for (const apt of apartments) {
      if (apt.dauerhaftBelegt) continue;

      const match = apt.regionen.includes(suchStadt) || suchStadt.includes(apt.regionen);
      if (!match) continue;

      if (apt.maxPersonen < anfragePers) continue;
      if (apt.schlafzimmer < zielZimmer) continue;

      const restBetten = apt.maxPersonen - anfragePers;
      if (restBetten > MAX_LEERBETTEN) continue;

      const aptKey = apt.name.toLowerCase();
      const bookings = bookingsMap[aptKey] || [];

      let slotGefunden = false;
      let tageVerschobenGefunden = 0;
      let gefundenerStartMs = 0;

      for (let shift = 1; shift <= MAX_VERSCHIEBUNG_TAGE; shift++) {
        const testStartMs = wunschStartMs + (shift * 86400000);
        const testEndeMs = testStartMs + aufenthaltDauerMs;

        const hatKollision = bookings.some(b => b.startMs < testEndeMs && b.endMs > testStartMs);
        if (!hatKollision) {
          slotGefunden = true;
          tageVerschobenGefunden = shift;
          gefundenerStartMs = testStartMs;
          break;
        }
      }

      if (!slotGefunden) continue;

      const gefundenerEndeMs = gefundenerStartMs + aufenthaltDauerMs;
      const altStartStr = new Date(gefundenerStartMs).toISOString().split("T")[0];
      const altEndeStr = new Date(gefundenerEndeMs).toISOString().split("T")[0];
      const restZimmer = Math.max(0, apt.schlafzimmer - zielZimmer);

      const score = (tageVerschobenGefunden * GEWICHT_TAG_VERSCHIEBUNG) +
        (restBetten * GEWICHT_LEERBETT) +
        (restZimmer * GEWICHT_LEERZIMMER);

      kandidaten.push({
        id: apt.code,
        restB: restBetten,
        restZ: restZimmer,
        sz: apt.schlafzimmer,
        istAlternative: true,
        tageVerschoben: tageVerschobenGefunden,
        altStart: altStartStr,
        altEnde: altEndeStr,
        penaltyScore: score
      });
    }

    kandidaten.sort((a, b) => a.penaltyScore - b.penaltyScore);

    if (kandidaten.length > 0) {
      const exportListe = kandidaten.slice(0, 2);
      leadsData[i][7] = JSON.stringify(exportListe);
      updatesMade = true;
      Logger.log(`Zeile ${rowNum}: Alternative in Spalte H eingetragen: ${exportListe[0].id} ab ${exportListe[0].altStart} (+${exportListe[0].tageVerschoben}d).`);
    } else {
      leadsData[i][7] = "[]";
      updatesMade = true;
      Logger.log(`Zeile ${rowNum}: Keine Alternative innerhalb von ${MAX_VERSCHIEBUNG_TAGE} Tagen gefunden.`);
    }
  }

  if (updatesMade) {
    const colHValues = leadsData.map(row => [row[7]]);
    leadsSheet.getRange(2, 8, numRows, 1).setValues(colHValues);
    for (let r = 2; r <= lastRowLeads; r++) {
      leadsSheet.setRowHeight(r, 21);
    }
    Logger.log("=== ALTERNATIVEN ERFOLGREICH IN SPALTE H GESPEICHERT ===");
  }
}

function ensureEngineAAltHeaderMMZ(sheet) {
  const cell = sheet.getRange(1, 8); // Spalte H
  if (!cell.getValue() || cell.getValue() !== "Alternative (<6P)") {
    cell.setValue("Alternative (<6P)");
    cell.setFontWeight("bold");
    sheet.getRange("H:H").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN FÜR ENGINE A-ALT
// -----------------------------------------------------------------
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