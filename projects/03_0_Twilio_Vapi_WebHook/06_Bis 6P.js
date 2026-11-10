// =================================================================
// STUFE 2: ENGINE A - SINGLE APARTMENT FINDER (AUDIO-CALLS)
// Liest Spalte K (Daten) & Spalte L (Routing), schreibt in Spalte M
// =================================================================
const KONFIGURATION_ENGINE_A = {
  INVENTORY_SPREADSHEET_ID: '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70',
  TABELLE_INVENTAR: 'M',
  TABELLE_KALENDER: 'cal',
  ZEILENHOEHE_PIXEL: 21,
  MILLISEKUNDEN_PRO_TAG: 86400000
};

// =================================================================
// ORCHESTRATOR / TRIGGER
// =================================================================
function runApartmentMatchingEngineA_Audio() {
  Logger.log("=== START ENGINE A MATCHING FÜR AUDIO-CALLS ===");
  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const letztesZeilenEnde = tabellenBlatt.getLastRow();
  if (letztesZeilenEnde <= 1) return;

  const inventarMappe = ladeInventarSpreadsheet(KONFIGURATION_ENGINE_A.INVENTORY_SPREADSHEET_ID);
  if (!inventarMappe) return;

  const inventarTabelle = inventarMappe.getSheetByName(KONFIGURATION_ENGINE_A.TABELLE_INVENTAR);
  const kalenderTabelle = inventarMappe.getSheetByName(KONFIGURATION_ENGINE_A.TABELLE_KALENDER);
  if (!inventarTabelle || !kalenderTabelle) {
    Logger.log("FEHLER: Tabellenblatt M oder cal fehlt im Inventory-Spreadsheet!");
    return;
  }

  stelleEngineAHeaderSicher(tabellenBlatt);

  const wohnungsListe = ladeWohnungenAusInventar(inventarTabelle);
  const buchungenKarte = ladeBuchungenAusKalender(kalenderTabelle);

  verarbeiteAudioLeadsBatch(tabellenBlatt, letztesZeilenEnde, wohnungsListe, buchungenKarte);
  Logger.log("=== ENGINE A ABGESCHLOSSEN ===");
}

// =================================================================
// MODUL: BATCH-VERARBEITUNG LEADS
// =================================================================
function verarbeiteAudioLeadsBatch(tabellenBlatt, zeilenAnzahl, wohnungen, buchungen) {
  const anzahlDatenZeilen = zeilenAnzahl - 1;
  // Liest Spalte K (11), L (12) und M (13)
  const datenBereich = tabellenBlatt.getRange(2, 11, anzahlDatenZeilen, 3);
  const zeilenWerte = datenBereich.getValues();
  let aenderungenVorhanden = false;

  const heuteUtcMs = ermittleHeuteUtcMs();

  for (let i = 0; i < zeilenWerte.length; i++) {
    const rawJsonK = String(zeilenWerte[i][0] || "").trim(); // Spalte K
    const rawRoutingL = String(zeilenWerte[i][1] || "").trim(); // Spalte L
    const optionsZelleM = String(zeilenWerte[i][2] || "").trim(); // Spalte M

    if (!istZeileFuerEngineAGueltig(rawJsonK, rawRoutingL, optionsZelleM)) continue;

    const routingObjekt = parseJsonSicher(rawRoutingL);
    if (!routingObjekt || (routingObjekt.status !== "STD_SHORT" && routingObjekt.status !== "STD_LONG")) {
      continue;
    }

    const datenObjekt = parseJsonSicher(rawJsonK);
    if (!datenObjekt) continue;

    const anfrageParameter = extrahiereAnfrageParameter(datenObjekt, routingObjekt.kennzahlen || {}, heuteUtcMs);
    if (!anfrageParameter) continue;

    const gefundeneOptionen = findeFreieWohnungen(anfrageParameter, wohnungen, buchungen);
    zeilenWerte[i][2] = JSON.stringify(erstelleExportOptionen(gefundeneOptionen));
    aenderungenVorhanden = true;
  }

  if (aenderungenVorhanden) {
    datenBereich.setValues(zeilenWerte);
    tabellenBlatt.setRowHeights(2, anzahlDatenZeilen, KONFIGURATION_ENGINE_A.ZEILENHOEHE_PIXEL);
    Logger.log("Ergebnisse erfolgreich in Spalte M geschrieben.");
  }
}

function istZeileFuerEngineAGueltig(rawK, rawL, rawM) {
  if (!rawK || rawK.startsWith("Fehler") || !rawL || rawL.startsWith("FEHLER")) {
    return false;
  }
  return rawM === "";
}

// =================================================================
// MODUL: PARAMETER-AUFBEREITUNG & LOGIK
// =================================================================
function extrahiereAnfrageParameter(datenObjekt, kennzahlen, heuteUtcMs) {
  const ortDaten = datenObjekt.ort || {};
  const personenDaten = datenObjekt.personen || {};
  const dauerDaten = datenObjekt.dauer || {};

  const anzahlPersonen = Number(personenDaten.Pers) || Number(kennzahlen.pers) || 0;
  const startDatum = String(dauerDaten.start || kennzahlen.start || "").trim();
  let endDatum = String(dauerDaten.ende || kennzahlen.ende || "").trim();

  // 6-Tage-Fallback für Anrufe, falls nur Startdatum genannt wurde
  if (startDatum && !endDatum) {
    const basisStart = new Date(startDatum + "T00:00:00Z");
    endDatum = new Date(basisStart.getTime() + 6 * KONFIGURATION_ENGINE_A.MILLISEKUNDEN_PRO_TAG).toISOString().split("T")[0];
  }

  const bereinigteStadt = extrahiereStadtName(ortDaten.suchortText || ortDaten.Ort);
  if (!bereinigteStadt || !startDatum || !endDatum || anzahlPersonen <= 1) {
    return null;
  }

  const anfrageStartMs = new Date(startDatum + "T00:00:00Z").getTime();
  const anfrageEndeMs = new Date(endDatum + "T00:00:00Z").getTime();
  if (anfrageEndeMs <= anfrageStartMs) return null;

  return {
    suchStadt: bereinigteStadt,
    personen: anzahlPersonen,
    zielZimmer: Number(kennzahlen.zimmer) || Math.ceil(anzahlPersonen / 2),
    startMs: anfrageStartMs,
    endeMs: anfrageEndeMs,
    tageBisAnreise: Math.round((anfrageStartMs - heuteUtcMs) / KONFIGURATION_ENGINE_A.MILLISEKUNDEN_PRO_TAG)
  };
}

function findeFreieWohnungen(anfrage, wohnungen, buchungen) {
  const treffer = [];

  for (const wohnung of wohnungen) {
    if (!stimmtRegionUeberein(wohnung.zulaessigeRegionen, anfrage.suchStadt)) continue;
    if (wohnung.istDauerhaftBelegt) continue;
    if (wohnung.maximalePersonen < anfrage.personen || wohnung.anzahlSchlafzimmer < anfrage.zielZimmer) continue;

    const kollision = pruefeKalenderKollision(wohnung.name, anfrage.startMs, anfrage.endeMs, buchungen);
    if (kollision.istBelegt) continue;

    const lueckenBewertung = berechneLueckenScore(kollision.letzteAbreiseMs, anfrage.startMs, anfrage.tageBisAnreise);

    treffer.push({
      id: wohnung.kuerzel,
      restB: wohnung.maximalePersonen - anfrage.personen,
      restZ: wohnung.anzahlSchlafzimmer - anfrage.zielZimmer,
      sz: wohnung.anzahlSchlafzimmer,
      lueckeScore: lueckenBewertung.score,
      lueckeT: lueckenBewertung.anzeigeTage
    });
  }

  return sortiereTreffer(treffer);
}

function sortiereTreffer(optionen) {
  return optionen.sort((a, b) => {
    if (a.restZ !== b.restZ) return a.restZ - b.restZ;
    if (a.restB !== b.restB) return a.restB - b.restB;
    return a.lueckeScore - b.lueckeScore;
  });
}

function erstelleExportOptionen(sortierteOptionen) {
  return sortierteOptionen.map(item => ({
    id: item.id,
    restB: item.restB,
    restZ: item.restZ,
    sz: item.sz,
    lueckeT: item.lueckeT
  }));
}

// =================================================================
// MODUL: KALENDER & INVENTAR DATENBANK
// =================================================================
function ladeWohnungenAusInventar(tabelle) {
  const daten = tabelle.getDataRange().getValues();
  const wohnungen = [];

  for (let i = 1; i < daten.length; i++) {
    const name = String(daten[i][0] || "").trim();
    if (!name) continue;

    const istAusgeschlossen = String(daten[i][35] || "").trim().toLowerCase() === "x"; // Spalte AJ
    if (istAusgeschlossen) continue;

    const kuerzel = reinigeKuerzel(daten[i][6]) || name; // Spalte G

    wohnungen.push({
      name: name,
      kuerzel: kuerzel,
      anzahlSchlafzimmer: Number(daten[i][14]) || 1, // Spalte O
      maximalePersonen: Number(daten[i][15]) || 2,   // Spalte P
      zulaessigeRegionen: String(daten[i][22] || "").toLowerCase(), // Spalte W
      istDauerhaftBelegt: String(daten[i][44] || "").trim().toLowerCase() === "x" // Spalte AS
    });
  }
  return wohnungen;
}

function ladeBuchungenAusKalender(tabelle) {
  const daten = tabelle.getDataRange().getValues();
  const buchungen = {};

  for (let i = 1; i < daten.length; i++) {
    const unterkunft = String(daten[i][0] || "").trim().toLowerCase();
    const startStr = wandleInIsoDatum(daten[i][1]);
    const endeStr = wandleInIsoDatum(daten[i][2]);

    if (!unterkunft || !startStr || !endeStr) continue;

    if (!buchungen[unterkunft]) {
      buchungen[unterkunft] = [];
    }

    buchungen[unterkunft].push({
      startMs: new Date(startStr + "T00:00:00Z").getTime(),
      endeMs: new Date(endeStr + "T00:00:00Z").getTime()
    });
  }
  return buchungen;
}

function pruefeKalenderKollision(name, startMs, endeMs, buchungen) {
  const liste = buchungen[name.toLowerCase()] || [];
  let istBelegt = false;
  let letzteAbreiseMs = null;

  for (const b of liste) {
    if (b.startMs < endeMs && b.endeMs > startMs) {
      istBelegt = true;
      break;
    }
    if (b.endeMs <= startMs) {
      if (letzteAbreiseMs === null || b.endeMs > letzteAbreiseMs) {
        letzteAbreiseMs = b.endeMs;
      }
    }
  }
  return { istBelegt: istBelegt, letzteAbreiseMs: letzteAbreiseMs };
}

// =================================================================
// MODUL: HILFSFUNKTIONEN
// =================================================================
function berechneLueckenScore(letzteAbreiseMs, anfrageStartMs, tageBisAnreise) {
  let anzeigeTage = null;
  if (letzteAbreiseMs !== null) {
    const differenz = Math.max(0, Math.round((anfrageStartMs - letzteAbreiseMs) / KONFIGURATION_ENGINE_A.MILLISEKUNDEN_PRO_TAG));
    if (differenz <= 14) anzeigeTage = differenz;
  }

  const score = anzeigeTage === null
    ? (tageBisAnreise <= 5 ? -1 : 999)
    : anzeigeTage;

  return { score: score, anzeigeTage: anzeigeTage };
}

function stimmtRegionUeberein(regionen, stadt) {
  return regionen.includes(stadt) || stadt.includes(regionen);
}

function extrahiereStadtName(roherText) {
  if (!roherText) return "";
  let stadt = String(roherText).toLowerCase();
  stadt = stadt.replace(/\b\d{4,5}\b/g, "");
  stadt = stadt.replace(/\b(nähe|raum|bereich|stadt|umgebung|zentrum|mitte|ot|ortsteil)\b/gi, "");
  stadt = stadt.replace(/[,\.\-\/\(\)]/g, " ");
  return stadt.replace(/\s+/g, " ").trim();
}

function wandleInIsoDatum(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "UTC", "yyyy-MM-dd");
  const str = String(val).trim();
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(str)) {
    const teile = str.split(".");
    return `${teile[2]}-${teile[1].padStart(2, "0")}-${teile[0].padStart(2, "0")}`;
  }
  const isoTeil = str.split(/[T\s]/)[0];
  return /^\d{4}-\d{2}-\d{2}$/.test(isoTeil) ? isoTeil : "";
}

function reinigeKuerzel(text) {
  if (!text) return "";
  return String(text).replace(/[^a-zA-Z0-9_-]/g, "").trim();
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

function ermittleHeuteUtcMs() {
  const d = new Date();
  return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())).getTime();
}

function stelleEngineAHeaderSicher(blatt) {
  const zelle = blatt.getRange(1, 13);
  if (!zelle.getValue()) {
    zelle.setValue("Optionen (<=6P)");
    zelle.setFontWeight("bold");
    blatt.getRange("M:M").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}