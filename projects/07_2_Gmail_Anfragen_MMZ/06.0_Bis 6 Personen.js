// =================================================================
// STUFE 2: ENGINE A - SINGLE APARTMENT FINDER FÜR TAB "MMZ"
// Liest Spalte D (Daten) & Spalte E (Routing), schreibt in Spalte F
// Prüft ausschließlich NEUE Zeilen (Spalte F komplett leer)
// =================================================================
const KONFIGURATION = {
  INVENTORY_SPREADSHEET_ID: '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70',
  TABELLE_LEADS: 'MMZ',
  TABELLE_INVENTAR: 'M',
  TABELLE_KALENDER: 'cal',
  ZEILENHOEHE_PIXEL: 21,
  MILLISEKUNDEN_PRO_TAG: 86400000
};

/**
 * Zentrale Orchestrator-Funktion (Trigger-Einstiegspunkt)
 */
function runApartmentMatchingEngineA_MMZ() {
  Logger.log("=== START ENGINE A MATCHING FÜR MMZ (STD_SHORT & STD_LONG) ===");
  const aktiveArbeitsmappe = SpreadsheetApp.getActiveSpreadsheet();
  const leadsTabelle = aktiveArbeitsmappe.getSheetByName(KONFIGURATION.TABELLE_LEADS);
  if (!leadsTabelle) {
    Logger.log("FEHLER: Tabellenblatt MMZ nicht gefunden!");
    return;
  }

  const letztesZeilenEnde = leadsTabelle.getLastRow();
  if (letztesZeilenEnde <= 1) return;

  const inventarMappe = ladeInventarSpreadsheet(KONFIGURATION.INVENTORY_SPREADSHEET_ID);
  if (!inventarMappe) return;

  const inventarTabelle = inventarMappe.getSheetByName(KONFIGURATION.TABELLE_INVENTAR);
  const kalenderTabelle = inventarMappe.getSheetByName(KONFIGURATION.TABELLE_KALENDER);
  if (!inventarTabelle || !kalenderTabelle) {
    Logger.log("FEHLER: Ein Tabellenblatt fehlt (M oder cal)!");
    return;
  }

  stelleEngineAHeaderSicher(leadsTabelle);
  const wohnungsListe = ladeWohnungenAusInventar(inventarTabelle);
  const buchungenKarte = ladeBuchungenAusKalender(kalenderTabelle);
  verarbeiteLeadsUndSchreibeErgebnisse(leadsTabelle, letztesZeilenEnde, wohnungsListe, buchungenKarte);
  Logger.log("=== ENGINE A ABGESCHLOSSEN ===");
}

/**
 * Öffnet das externe Inventar-Spreadsheet isoliert mit Fehlerabfang
 */
function ladeInventarSpreadsheet(tabellenId) {
  try {
    return SpreadsheetApp.openById(tabellenId);
  } catch (fehler) {
    Logger.log("FEHLER beim Öffnen des Inventory-Spreadsheets: " + fehler.toString());
    return null;
  }
}

/**
 * Liest alle Apartments aus Tab 'M' ein
 * Ignoriert Zeilen, wenn in Spalte AJ ("ausschluss") ein "x" steht
 */
function ladeWohnungenAusInventar(tabelle) {
  const rohWerte = tabelle.getDataRange().getValues();
  const wohnungen = [];

  for (let zeile = 1; zeile < rohWerte.length; zeile++) {
    const wohnungsName = String(rohWerte[zeile][0] || "").trim();
    if (!wohnungsName) continue;

    // Spalte AJ (Index 35): Prüfung auf Ausschluss "x"
    const istAusgeschlossen = String(rohWerte[zeile][35] || "").trim().toLowerCase() === "x";
    if (istAusgeschlossen) continue;

    const rohesKuerzel = String(rohWerte[zeile][6] || "").trim();
    const bereinigtesKuerzel = cleanShortCode(rohesKuerzel) || wohnungsName;

    wohnungen.push({
      name: wohnungsName,
      kuerzel: bereinigtesKuerzel,
      anzahlSchlafzimmer: Number(rohWerte[zeile][14]) || 1,
      maximalePersonen: Number(rohWerte[zeile][15]) || 2,
      zulaessigeRegionen: String(rohWerte[zeile][22] || "").toLowerCase(),
      istDauerhaftBelegt: String(rohWerte[zeile][44] || "").trim().toLowerCase() === "x"
    });
  }

  Logger.log(`Geladene Apartments aus Tab M (nach Ausschluss Spalte AJ): ${wohnungen.length}`);
  return wohnungen;
}

/**
 * Liest Buchungen aus dem Kalenderblatt 'cal' ein
 */
function ladeBuchungenAusKalender(tabelle) {
  const daten = tabelle.getDataRange().getValues();
  const buchungenProWohnung = {};

  for (let zeile = 1; zeile < daten.length; zeile++) {
    const unterkunftsName = String(daten[zeile][0] || "").trim().toLowerCase();
    const startDatumText = parseIsoDateString(daten[zeile][1]);
    const endDatumText = parseIsoDateString(daten[zeile][2]);

    if (!unterkunftsName || !startDatumText || !endDatumText) continue;

    if (!buchungenProWohnung[unterkunftsName]) {
      buchungenProWohnung[unterkunftsName] = [];
    }
    buchungenProWohnung[unterkunftsName].push({
      startMillisekunden: new Date(startDatumText + "T00:00:00Z").getTime(),
      endeMillisekunden: new Date(endDatumText + "T00:00:00Z").getTime()
    });
  }
  return buchungenProWohnung;
}

/**
 * Iteriert über Leads, matcht passende Wohnungen und speichert gesammelt
 */
function verarbeiteLeadsUndSchreibeErgebnisse(leadsTabelle, zeilenAnzahl, wohnungen, buchungen) {
  const anzahlDatenZeilen = zeilenAnzahl - 1;
  const leadsBereich = leadsTabelle.getRange(2, 1, anzahlDatenZeilen, 6);
  const leadsDaten = leadsBereich.getValues();
  let aenderungenVorhanden = false;

  const heute = new Date();
  const heuteUtcMs = new Date(Date.UTC(heute.getFullYear(), heute.getMonth(), heute.getDate())).getTime();

  for (let index = 0; index < leadsDaten.length; index++) {
    const aktuelleZeile = index + 2;
    const leadSpalten = leadsDaten[index];

    // Nur völlig unberührte Zeilen bearbeiten
    if (!pruefeLeadGueltigkeit(leadSpalten)) continue;

    const routingObjekt = parseJsonSicher(leadSpalten[4]);
    if (!routingObjekt || (routingObjekt.status !== "STD_SHORT" && routingObjekt.status !== "STD_LONG")) {
      continue;
    }

    const datenObjekt = parseJsonSicher(leadSpalten[3]);
    if (!datenObjekt) {
      Logger.log(`Zeile ${aktuelleZeile}: JSON Parse Fehler in Spalte D`);
      continue;
    }

    const leadParameter = extrahiereLeadParameter(datenObjekt, routingObjekt.kennzahlen || {}, aktuelleZeile, heuteUtcMs);
    if (!leadParameter) continue;

    const passendeWohnungen = findeFreieWohnungen(leadParameter, wohnungen, buchungen);
    protokolliereUndSetzeTreffer(leadsDaten, index, passendeWohnungen, leadParameter, aktuelleZeile);
    aenderungenVorhanden = true;
  }

  if (aenderungenVorhanden) {
    leadsBereich.setValues(leadsDaten);
    leadsTabelle.setRowHeights(2, anzahlDatenZeilen, KONFIGURATION.ZEILENHOEHE_PIXEL);
    Logger.log("=== MATCHING ERFOLGREICH IN SPALTE F GESCHRIEBEN ===");
  } else {
    Logger.log("Keine neuen Zeilen zur Bearbeitung vorhanden.");
  }
}

/**
 * Prüft Zeilengültigkeit: Verarbeitet NUR Zeilen, die in Spalte F noch komplett LEER sind
 */
function pruefeLeadGueltigkeit(zeile) {
  const eindeutigeId = String(zeile[1] || "").trim();
  const datenJson = String(zeile[3] || "").trim();
  const routingJson = String(zeile[4] || "").trim();
  const optionsZelle = String(zeile[5] || "").trim();

  if (!eindeutigeId || eindeutigeId === "-" || !datenJson || !routingJson || routingJson.startsWith("FEHLER")) {
    return false;
  }
  return optionsZelle === "";
}

/**
 * Bereitet Anfrageparameter auf, berechnet Vorlaufzeit und setzt Fallback bei fehlendem Enddatum
 */
function extrahiereLeadParameter(datenObjekt, kennzahlen, zeilennummer, heuteUtcMs) {
  const ortDaten = datenObjekt.ort || {};
  const personenDaten = datenObjekt.personen || {};
  const dauerDaten = datenObjekt.dauer || {};

  const anzahlPersonen = Number(personenDaten.Pers) || Number(kennzahlen.pers) || 0;
  const startDatum = String(dauerDaten.start || "").trim();
  let endDatum = String(dauerDaten.ende || "").trim();

  // Einheitlicher 6-Tage-Fallback, falls nur Startdatum vorhanden ist
  if (startDatum && !endDatum) {
    const basisStart = new Date(startDatum + "T00:00:00Z");
    endDatum = new Date(basisStart.getTime() + 6 * KONFIGURATION.MILLISEKUNDEN_PRO_TAG).toISOString().split("T")[0];
  }

  const zielZimmer = Number(kennzahlen.zimmer) || Math.ceil(anzahlPersonen / 2);
  const roherOrtText = String(ortDaten.suchortText || ortDaten.Ort || "").trim();
  const bereinigteStadt = extractPureCityName(roherOrtText);

  if (!bereinigteStadt || !startDatum || !endDatum || anzahlPersonen <= 1) {
    Logger.log(`Zeile ${zeilennummer}: Fehlende Pflichtdaten oder ungültige Personenanzahl (Stadt: "${bereinigteStadt}", Start: "${startDatum}", Ende: "${endDatum}", Pers: ${anzahlPersonen})`);
    return null;
  }

  const anfrageStartMs = new Date(startDatum + "T00:00:00Z").getTime();
  const anfrageEndeMs = new Date(endDatum + "T00:00:00Z").getTime();

  if (anfrageEndeMs <= anfrageStartMs) {
    Logger.log(`Zeile ${zeilennummer}: Ungültige Reisedauer (Ende vor Start).`);
    return null;
  }

  const tageBisAnreise = Math.round((anfrageStartMs - heuteUtcMs) / KONFIGURATION.MILLISEKUNDEN_PRO_TAG);

  return {
    suchStadt: bereinigteStadt,
    personen: anzahlPersonen,
    zielZimmer: zielZimmer,
    startMs: anfrageStartMs,
    endeMs: anfrageEndeMs,
    tageBisAnreise: tageBisAnreise
  };
}

/**
 * Sucht und bewertet alle freien Wohnungen für eine Anfrage
 */
function findeFreieWohnungen(anfrage, wohnungen, buchungen) {
  const gefundeneOptionen = [];

  for (const wohnung of wohnungen) {
    if (!stimmtRegionUeberein(wohnung.zulaessigeRegionen, anfrage.suchStadt)) continue;
    if (wohnung.istDauerhaftBelegt) continue;
    if (wohnung.maximalePersonen < anfrage.personen || wohnung.anzahlSchlafzimmer < anfrage.zielZimmer) continue;

    const kollision = pruefeKalenderVerfuegbarkeit(wohnung.name, anfrage.startMs, anfrage.endeMs, buchungen);
    if (kollision.istBelegt) continue;

    const lueckenBewertung = berechneLueckenScore(kollision.letzteAbreiseMs, anfrage.startMs, anfrage.tageBisAnreise);

    gefundeneOptionen.push({
      id: wohnung.kuerzel,
      restB: wohnung.maximalePersonen - anfrage.personen,
      restZ: wohnung.anzahlSchlafzimmer - anfrage.zielZimmer,
      sz: wohnung.anzahlSchlafzimmer,
      lueckeScore: lueckenBewertung.score,
      lueckeT: lueckenBewertung.anzeigeTage
    });
  }
  return gefundeneOptionen;
}

/**
 * Protokolliert Trefferquote und setzt Ergebniswert für Spalte F
 */
function protokolliereUndSetzeTreffer(leadsDaten, zeilenIndex, gefundeneOptionen, anfrage, zeilennummer) {
  if (gefundeneOptionen.length > 0) {
    gefundeneOptionen.sort((a, b) => {
      if (a.restZ !== b.restZ) return a.restZ - b.restZ;
      if (a.restB !== b.restB) return a.restB - b.restB;
      return a.lueckeScore - b.lueckeScore;
    });

    const exportListe = gefundeneOptionen.map(item => ({
      id: item.id,
      restB: item.restB,
      restZ: item.restZ,
      sz: item.sz,
      lueckeT: item.lueckeT
    }));

    leadsDaten[zeilenIndex][5] = JSON.stringify(exportListe);
    Logger.log(`Zeile ${zeilennummer} (${anfrage.suchStadt}, Anreise in ${anfrage.tageBisAnreise}d, ${anfrage.personen}P / ${anfrage.zielZimmer}Z): ${exportListe.length} freie(s) Apartment(s) gematcht.`);
  } else {
    leadsDaten[zeilenIndex][5] = "[]";
    Logger.log(`Zeile ${zeilennummer} (${anfrage.suchStadt}, ${anfrage.personen}P / ${anfrage.zielZimmer}Z): Keine passenden Single-Apartments gefunden.`);
  }
}

/**
 * Prüft Übereinstimmung zwischen Regionsangabe und Suchstadt
 */
function stimmtRegionUeberein(wohnungsRegionen, gesuchteStadt) {
  return wohnungsRegionen.includes(gesuchteStadt) || gesuchteStadt.includes(wohnungsRegionen);
}

/**
 * Prüft Kalenderkollisionen und ermittelt die vorherige Abreise
 */
function pruefeKalenderVerfuegbarkeit(wohnungsName, anfrageStartMs, anfrageEndeMs, buchungen) {
  const existierendeBuchungen = buchungen[wohnungsName.toLowerCase()] || [];
  let istBelegt = false;
  let letzteAbreiseMs = null;

  for (const buchung of existierendeBuchungen) {
    if (buchung.startMillisekunden < anfrageEndeMs && buchung.endeMillisekunden > anfrageStartMs) {
      istBelegt = true;
      break;
    }
    if (buchung.endeMillisekunden <= anfrageStartMs) {
      if (letzteAbreiseMs === null || buchung.endeMillisekunden > letzteAbreiseMs) {
        letzteAbreiseMs = buchung.endeMillisekunden;
      }
    }
  }
  return { istBelegt: istBelegt, letzteAbreiseMs: letzteAbreiseMs };
}

/**
 * Berechnet Lücken-Score anhand der Leerstandstage und Vorlaufzeit
 */
function berechneLueckenScore(letzteAbreiseMs, anfrageStartMs, tageBisAnreise) {
  let anzeigeTage = null;
  if (letzteAbreiseMs !== null) {
    const differenzTage = Math.max(0, Math.round((anfrageStartMs - letzteAbreiseMs) / KONFIGURATION.MILLISEKUNDEN_PRO_TAG));
    if (differenzTage <= 14) {
      anzeigeTage = differenzTage;
    }
  }

  let bewertungsScore = anzeigeTage === null
    ? (tageBisAnreise <= 5 ? -1 : 999)
    : anzeigeTage;

  return { score: bewertungsScore, anzeigeTage: anzeigeTage };
}

/**
 * Kapselt JSON.parse() sicher in einem try/catch Block
 */
function parseJsonSicher(jsonZeichenkette) {
  try {
    return JSON.parse(jsonZeichenkette);
  } catch (fehler) {
    return null;
  }
}

/**
 * Setzt den Spalten-Header für Spalte F
 */
function stelleEngineAHeaderSicher(tabelle) {
  const kopfZelle = tabelle.getRange(1, 6);
  if (!kopfZelle.getValue()) {
    kopfZelle.setValue("Optionen (<=6P)");
    kopfZelle.setFontWeight("bold");
    tabelle.getRange("F:F").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

/**
 * Bereinigt Kürzel von Sonderzeichen
 */
function cleanShortCode(kuerzelText) {
  if (!kuerzelText) return "";
  return String(kuerzelText).replace(/[^a-zA-Z0-9_-]/g, "").trim();
}

/**
 * Wandelt Datumsangaben (Date-Objekt, DD.MM.YYYY oder YYYY-MM-DD) in ein einheitliches ISO-Format (YYYY-MM-DD) um
 */
function parseIsoDateString(eingabeDatum) {
  if (!eingabeDatum) return "";

  if (eingabeDatum instanceof Date) {
    return Utilities.formatDate(eingabeDatum, "UTC", "yyyy-MM-dd");
  }

  const datumString = String(eingabeDatum).trim();

  // Deutsches Format DD.MM.YYYY erkennen und in YYYY-MM-DD umwandeln
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(datumString)) {
    const teile = datumString.split(".");
    return `${teile[2]}-${teile[1].padStart(2, "0")}-${teile[0].padStart(2, "0")}`;
  }

  // ISO-Format YYYY-MM-DD (auch wenn Zeitanteile mit T oder Leerzeichen folgen)
  const isoTeil = datumString.split(/[T\s]/)[0];
  return /^\d{4}-\d{2}-\d{2}$/.test(isoTeil) ? isoTeil : "";
}

/**
 * Filtert Füllwörter, Postleitzahlen und Satzzeichen aus Stadtnamen
 */
function extractPureCityName(roherOrt) {
  if (!roherOrt) return "";
  let stadt = String(roherOrt).toLowerCase();
  stadt = stadt.replace(/\b\d{4,5}\b/g, "");
  stadt = stadt.replace(/\b(nähe|raum|bereich|stadt|umgebung|zentrum|mitte|ot|ortsteil)\b/gi, "");
  stadt = stadt.replace(/[,\.\-\/\(\)]/g, " ");
  return stadt.replace(/\s+/g, " ").trim();
}