// =================================================================
// STUFE 2a: ENGINE A-ALTERNATIVE (NEXT-AVAILABLE FÜR AUDIO-CALLS)
// Liest Spalte K (Daten), Spalte L (Routing) & Spalte M (Engine A)
// Schreibt alternative Optionen gebündelt in Spalte O
// =================================================================
const KONFIGURATION_ENGINE_A_ALT = {
  INVENTORY_SPREADSHEET_ID: '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70',
  TABELLE_INVENTAR: 'M',
  TABELLE_KALENDER: 'cal',
  MAX_VERSCHIEBUNG_TAGE: 14,
  MAX_LEERBETTEN: 3,
  GEWICHT_TAG_VERSCHIEBUNG: 10,
  GEWICHT_LEERBETT: 3,
  GEWICHT_LEERZIMMER: 4,
  ZEILENHOEHE_PIXEL: 21,
  MS_PRO_TAG: 86400000
};

// =================================================================
// ORCHESTRATOR / TRIGGER
// =================================================================
function runApartmentMatchingEngineA_Alternative_Audio() {
  Logger.log("=== START ENGINE A-ALT (NEXT AVAILABLE AUDIO) ===");
  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const letzteZeile = tabellenBlatt.getLastRow();
  if (letzteZeile <= 1) return;

  const inventarMappe = ladeInventarSpreadsheet(KONFIGURATION_ENGINE_A_ALT.INVENTORY_SPREADSHEET_ID);
  if (!inventarMappe) return;

  const invSheetM = inventarMappe.getSheetByName(KONFIGURATION_ENGINE_A_ALT.TABELLE_INVENTAR);
  const invSheetCal = inventarMappe.getSheetByName(KONFIGURATION_ENGINE_A_ALT.TABELLE_KALENDER);
  if (!invSheetM || !invSheetCal) {
    Logger.log("FEHLER: Ein Tabellenblatt fehlt (M oder cal)!");
    return;
  }

  stelleEngineAAltHeaderSicher(tabellenBlatt);

  const wohnungen = ladeWohnungenFuerAlternative(invSheetM);
  const buchungenKarte = ladeBuchungenSortiert(invSheetCal);

  verarbeiteAlternativenBatch(tabellenBlatt, letzteZeile, wohnungen, buchungenKarte);
  Logger.log("=== ENGINE A-ALT ABGESCHLOSSEN ===");
}

// =================================================================
// MODUL: BATCH-VERARBEITUNG LEADS
// =================================================================
function verarbeiteAlternativenBatch(tabellenBlatt, letzteZeile, wohnungen, buchungenKarte) {
  const anzahlDatenZeilen = letzteZeile - 1;
  // Liest Spalten K (11), L (12), M (13), N (14), O (15) = 5 Spalten
  const datenBereich = tabellenBlatt.getRange(2, 11, anzahlDatenZeilen, 5);
  const zeilenWerte = datenBereich.getValues();
  let aenderungenVorhanden = false;

  for (let i = 0; i < zeilenWerte.length; i++) {
    const rawK = String(zeilenWerte[i][0] || "").trim(); // Spalte K
    const rawL = String(zeilenWerte[i][1] || "").trim(); // Spalte L
    const rawM = String(zeilenWerte[i][2] || "").trim(); // Spalte M (Engine A)
    const rawO = String(zeilenWerte[i][4] || "").trim(); // Spalte O (Ziel)

    if (!istZeileFuerAlternativeBereit(rawK, rawL, rawM, rawO)) continue;

    const routingObj = parseJsonSicher(rawL);
    if (!routingObj || (routingObj.status !== "STD_SHORT" && routingObj.status !== "STD_LONG")) continue;

    const datenObj = parseJsonSicher(rawK);
    if (!datenObj) continue;

    const leadParam = extrahiereAlternativeParameter(datenObj, routingObj.kennzahlen || {});
    if (!leadParam) continue;

    const kandidaten = findeAlternativeSlots(leadParam, wohnungen, buchungenKarte);
    zeilenWerte[i][4] = kandidaten.length > 0 ? JSON.stringify(kandidaten.slice(0, 2)) : "[]";
    aenderungenVorhanden = true;
  }

  if (aenderungenVorhanden) {
    datenBereich.setValues(zeilenWerte);
    tabellenBlatt.setRowHeights(2, anzahlDatenZeilen, KONFIGURATION_ENGINE_A_ALT.ZEILENHOEHE_PIXEL);
    Logger.log("Alternativen erfolgreich in Spalte O gespeichert.");
  }
}

function istZeileFuerAlternativeBereit(rawK, rawL, rawM, rawO) {
  if (!rawK || rawK.startsWith("Fehler") || !rawL || rawL.startsWith("FEHLER")) return false;
  if (rawM !== "[]") return false; // Nur ausführen, wenn Sofort-Matching leer war
  return rawO === "" || rawO === "[]";
}

// =================================================================
// MODUL: SLOT-SUCHE & BEWERTUNG
// =================================================================
function extrahiereAlternativeParameter(datenObj, kennzahlen) {
  const anfragePers = Number(datenObj.personen?.Pers) || Number(kennzahlen.pers) || 0;
  const startStr = String(datenObj.dauer?.start || kennzahlen.start || "").trim();
  let endeStr = String(datenObj.dauer?.ende || kennzahlen.ende || "").trim();

  if (startStr && !endeStr) {
    const basisStart = new Date(startStr + "T00:00:00Z");
    endeStr = new Date(basisStart.getTime() + 6 * KONFIGURATION_ENGINE_A_ALT.MS_PRO_TAG).toISOString().split("T")[0];
  }

  const suchStadt = extractPureCityName(datenObj.ort?.suchortText || datenObj.ort?.Ort);
  if (!suchStadt || !startStr || !endeStr || anfragePers <= 1) return null;

  const wunschStartMs = new Date(startStr + "T00:00:00Z").getTime();
  const wunschEndeMs = new Date(endeStr + "T00:00:00Z").getTime();
  const dauerMs = wunschEndeMs - wunschStartMs;
  if (dauerMs <= 0) return null;

  return {
    personen: anfragePers,
    zielZimmer: Number(kennzahlen.zimmer) || Math.ceil(anfragePers / 2),
    stadt: suchStadt,
    startMs: wunschStartMs,
    dauerMs: dauerMs
  };
}

function findeAlternativeSlots(leadParam, wohnungen, buchungenKarte) {
  const kandidaten = [];

  for (const apt of wohnungen) {
    if (!pruefeAptBasisKriterien(apt, leadParam)) continue;

    const restBetten = apt.maxPersonen - leadParam.personen;
    if (restBetten > KONFIGURATION_ENGINE_A_ALT.MAX_LEERBETTEN) continue;

    const bookings = buchungenKarte[apt.name.toLowerCase()] || [];
    const slot = pruefeVerschiebung(bookings, leadParam.startMs, leadParam.dauerMs);
    if (!slot) continue;

    const restZimmer = Math.max(0, apt.schlafzimmer - leadParam.zielZimmer);
    const score = berechneAlternativeScore(slot.tageVerschoben, restBetten, restZimmer);

    kandidaten.push({
      id: apt.code,
      restB: restBetten,
      restZ: restZimmer,
      sz: apt.schlafzimmer,
      istAlternative: true,
      tageVerschoben: slot.tageVerschoben,
      altStart: new Date(slot.startMs).toISOString().split("T")[0],
      altEnde: new Date(slot.startMs + leadParam.dauerMs).toISOString().split("T")[0],
      penaltyScore: score
    });
  }

  return kandidaten.sort((a, b) => a.penaltyScore - b.penaltyScore);
}

function pruefeVerschiebung(bookings, startMs, dauerMs) {
  for (let shift = 1; shift <= KONFIGURATION_ENGINE_A_ALT.MAX_VERSCHIEBUNG_TAGE; shift++) {
    const testStartMs = startMs + (shift * KONFIGURATION_ENGINE_A_ALT.MS_PRO_TAG);
    const testEndeMs = testStartMs + dauerMs;

    const hatKollision = bookings.some(b => b.startMs < testEndeMs && b.endMs > testStartMs);
    if (!hatKollision) {
      return { tageVerschoben: shift, startMs: testStartMs };
    }
  }
  return null;
}

function berechneAlternativeScore(verschiebung, restBetten, restZimmer) {
  return (verschiebung * KONFIGURATION_ENGINE_A_ALT.GEWICHT_TAG_VERSCHIEBUNG) +
         (restBetten * KONFIGURATION_ENGINE_A_ALT.GEWICHT_LEERBETT) +
         (restZimmer * KONFIGURATION_ENGINE_A_ALT.GEWICHT_LEERZIMMER);
}

function pruefeAptBasisKriterien(apt, leadParam) {
  if (apt.dauerhaftBelegt) return false;
  if (!apt.regionen.includes(leadParam.stadt) && !leadParam.stadt.includes(apt.regionen)) return false;
  if (apt.maxPersonen < leadParam.personen || apt.schlafzimmer < leadParam.zielZimmer) return false;
  return true;
}

// =================================================================
// MODUL: INVENTAR & KALENDER DATENBANK
// =================================================================
function ladeWohnungenFuerAlternative(invSheetM) {
  const mData = invSheetM.getDataRange().getValues();
  const apartments = [];

  for (let r = 1; r < mData.length; r++) {
    const name = String(mData[r][0] || "").trim();
    if (!name) continue;

    const istAusgeschlossen = String(mData[r][35] || "").trim().toLowerCase() === "x";
    if (istAusgeschlossen) continue;

    const shortCode = cleanShortCode(mData[r][6]) || name;

    apartments.push({
      name: name,
      code: shortCode,
      schlafzimmer: Number(mData[r][14]) || 1,
      maxPersonen: Number(mData[r][15]) || 2,
      regionen: String(mData[r][22] || "").toLowerCase(),
      dauerhaftBelegt: String(mData[r][44] || "").trim().toLowerCase() === "x"
    });
  }
  return apartments;
}

function ladeBuchungenSortiert(invSheetCal) {
  const calData = invSheetCal.getDataRange().getValues();
  const bookingsMap = {};

  for (let c = 1; c < calData.length; c++) {
    const propName = String(calData[c][0] || "").trim().toLowerCase();
    const startStr = parseIsoDateString(calData[c][1]);
    const endStr = parseIsoDateString(calData[c][2]);
    if (!propName || !startStr || !endStr) continue;

    if (!bookingsMap[propName]) bookingsMap[propName] = [];
    bookingsMap[propName].push({
      startMs: new Date(startStr + "T00:00:00Z").getTime(),
      endMs: new Date(endStr + "T00:00:00Z").getTime()
    });
  }

  for (const k in bookingsMap) {
    bookingsMap[k].sort((a, b) => a.startMs - b.startMs);
  }
  return bookingsMap;
}

// =================================================================
// MODUL: HILFSFUNKTIONEN
// =================================================================
function cleanShortCode(rawCode) {
  if (!rawCode) return "";
  return String(rawCode).trim().replace(/[\r\n\t]/g, "");
}

function parseIsoDateString(eingabeDatum) {
  if (!eingabeDatum) return "";
  if (eingabeDatum instanceof Date) return Utilities.formatDate(eingabeDatum, "UTC", "yyyy-MM-dd");
  const datumString = String(eingabeDatum).trim();
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(datumString)) {
    const teile = datumString.split(".");
    return `${teile[2]}-${teile[1].padStart(2, "0")}-${teile[0].padStart(2, "0")}`;
  }
  const isoTeil = datumString.split(/[T\s]/)[0];
  return /^\d{4}-\d{2}-\d{2}$/.test(isoTeil) ? isoTeil : "";
}

function extractPureCityName(text) {
  if (!text) return "";
  let clean = String(text).toLowerCase().trim();
  clean = clean.replace(/\b\d{5}\b/g, "").trim();
  clean = clean.replace(/\b(nähe|raum|umkreis|zentrum|city|hbf|bahnhof|stadt)\b/gi, "").trim();
  if (clean.includes(",")) clean = clean.split(",")[0].trim();
  return clean.replace(/\s+/g, " ");
}

function parseJsonSicher(text) {
  try {
    return JSON.parse(text);
  } catch (e) {
    return null;
  }
}

function ladeInventarSpreadsheet(id) {
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    Logger.log("FEHLER beim Öffnen des Inventory-Sheets: " + e.toString());
    return null;
  }
}

function stelleEngineAAltHeaderSicher(blatt) {
  const headerO = blatt.getRange(1, 15);
  if (!headerO.getValue() || headerO.getValue() !== "Alternative (<6P)") {
    headerO.setValue("Alternative (<6P)");
    headerO.setFontWeight("bold");
    blatt.getRange("O:O").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}