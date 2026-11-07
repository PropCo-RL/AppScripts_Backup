// =================================================================
// STUFE 2b-Alt: NEXT-AVAILABLE KOMBI-FINDER FÜR TAB "MMZ"
// Liest Spalte D (Daten), Spalte E (Routing), Spalte G (Engine B)
// Schreibt Alternativ-Kombis AUSSCHLIESSLICH in Spalte I (Spalte 9)
// =================================================================

const KONFIGURATION_B_ALT = {
  INVENTORY_SPREADSHEET_ID: '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70',
  TABELLE_LEADS: 'MMZ',
  TABELLE_INVENTAR: 'M',
  TABELLE_KALENDER: 'cal',
  MAX_VERSCHIEBUNG_TAGE: 14,
  MAX_KOMBI_GROESSE: 4,
  MAX_KANDIDATEN_SLOT: 20,
  GEWICHT_TAG_VERSCHIEBUNG: 15,
  GEWICHT_LEERBETT: 2,
  GEWICHT_LEERZIMMER: 3,
  GEWICHT_APARTMENT_ANZAHL: 5,
  ZEILENHOEHE_PIXEL: 21,
  MILLISEKUNDEN_PRO_TAG: 86400000
};

/**
 * Zentrale Orchestrator-Funktion (Trigger-Einstiegspunkt)
 */
function runApartmentMatchingEngineB_Alternative_MMZ() {
  Logger.log("=== START ENGINE B-ALT (NEXT AVAILABLE KOMBI) ===");
  const aktiveArbeitsmappe = SpreadsheetApp.getActiveSpreadsheet();
  const leadsTabelle = aktiveArbeitsmappe.getSheetByName(KONFIGURATION_B_ALT.TABELLE_LEADS);
  if (!leadsTabelle) {
    Logger.log("FEHLER: Tabellenblatt MMZ nicht gefunden!");
    return;
  }

  const anzahlZeilen = leadsTabelle.getLastRow();
  if (anzahlZeilen <= 1) return;

  const inventarMappe = ladeInventarSpreadsheetBAlt(KONFIGURATION_B_ALT.INVENTORY_SPREADSHEET_ID);
  if (!inventarMappe) return;

  const inventarTabelle = inventarMappe.getSheetByName(KONFIGURATION_B_ALT.TABELLE_INVENTAR);
  const kalenderTabelle = inventarMappe.getSheetByName(KONFIGURATION_B_ALT.TABELLE_KALENDER);
  if (!inventarTabelle || !kalenderTabelle) {
    Logger.log("FEHLER: Ein Tabellenblatt fehlt (M oder cal)!");
    return;
  }

  stelleEngineBAltHeaderSicher(leadsTabelle);
  const wohnungsListe = ladeWohnungenAusInventarBAlt(inventarTabelle);
  const buchungenKarte = ladeBuchungenAusKalenderBAlt(kalenderTabelle);

  verarbeiteKombiAlternativenUndSchreibeErgebnisse(leadsTabelle, anzahlZeilen, wohnungsListe, buchungenKarte);
  Logger.log("=== ENGINE B-ALT ABGESCHLOSSEN ===");
}

/**
 * Öffnet das externe Inventar-Spreadsheet isoliert mit Fehlerabfang
 */
function ladeInventarSpreadsheetBAlt(tabellenId) {
  try {
    return SpreadsheetApp.openById(tabellenId);
  } catch (fehler) {
    Logger.log("FEHLER beim Öffnen des Inventory-Spreadsheets: " + fehler.toString());
    return null;
  }
}

/**
 * Liest alle Apartments aus Tab 'M' ein
 * Ignoriert Zeilen mit "x" in Spalte AJ (Index 35, Ausschluss)
 */
function ladeWohnungenAusInventarBAlt(tabelle) {
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
 * Liest Buchungen aus dem Kalenderblatt 'cal' ein und sortiert chronologisch
 */
function ladeBuchungenAusKalenderBAlt(tabelle) {
  const rohDaten = tabelle.getDataRange().getValues();
  const buchungenProWohnung = {};

  for (let zeile = 1; zeile < rohDaten.length; zeile++) {
    const unterkunftsName = String(rohDaten[zeile][0] || "").trim().toLowerCase();
    const startDatumText = parseIsoDateString(rohDaten[zeile][1]);
    const endDatumText = parseIsoDateString(rohDaten[zeile][2]);

    if (!unterkunftsName || !startDatumText || !endDatumText) continue;

    if (!buchungenProWohnung[unterkunftsName]) {
      buchungenProWohnung[unterkunftsName] = [];
    }

    buchungenProWohnung[unterkunftsName].push({
      startMillisekunden: new Date(startDatumText + "T00:00:00Z").getTime(),
      endeMillisekunden: new Date(endDatumText + "T00:00:00Z").getTime()
    });
  }

  for (const name in buchungenProWohnung) {
    buchungenProWohnung[name].sort((a, b) => a.startMillisekunden - b.startMillisekunden);
  }

  return buchungenProWohnung;
}

/**
 * Iteriert über Leads, ermittelt Alternativ-Kombinationen und schreibt isoliert nach Spalte I
 */
function verarbeiteKombiAlternativenUndSchreibeErgebnisse(leadsTabelle, zeilenAnzahl, wohnungen, buchungen) {
  const anzahlDatenZeilen = zeilenAnzahl - 1;
  const leadsBereich = leadsTabelle.getRange(2, 1, anzahlDatenZeilen, 9);
  const leadsDaten = leadsBereich.getValues();
  let aenderungenVorhanden = false;

  for (let index = 0; index < anzahlDatenZeilen; index++) {
    const aktuelleZeile = index + 2;
    const leadSpalten = leadsDaten[index];

    if (!pruefeLeadGueltigkeitBAlt(leadSpalten)) continue;

    const routingObjekt = parseJsonSicher(leadSpalten[4]);
    if (!routingObjekt || !istZulaessigerRoutingStatus(routingObjekt.status)) continue;

    const datenObjekt = parseJsonSicher(leadSpalten[3]);
    if (!datenObjekt) continue;

    const anfrage = extrahiereAnfrageParameterBAlt(datenObjekt, routingObjekt.kennzahlen || {});
    if (!anfrage) continue;

    const regionWohnungen = filtereWohnungenFuerRegion(wohnungen, anfrage.suchStadt);
    if (regionWohnungen.length === 0) {
      leadsDaten[index][8] = "[]";
      aenderungenVorhanden = true;
      continue;
    }

    const kombiAlternativen = sucheKombiAlternativenImZeitfenster(regionWohnungen, anfrage, buchungen);
    protokolliereUndSetzeKombiAlternative(leadsDaten, index, kombiAlternativen, aktuelleZeile);
    aenderungenVorhanden = true;
  }

  if (aenderungenVorhanden) {
    schreibeErgebnisseNachSpalteI(leadsTabelle, leadsDaten, anzahlDatenZeilen);
  } else {
    Logger.log("Keine Zeilen für Kombi-Alternativ-Suche vorhanden.");
  }
}

/**
 * Validiert die Zeile: Verarbeitet nur, wenn Engine B leer gematcht hat (Spalte G === "[]")
 * und Spalte I noch nicht befüllt ist.
 */
function pruefeLeadGueltigkeitBAlt(zeile) {
  const eindeutigeId = String(zeile[1] || "").trim();
  const datenJson = String(zeile[3] || "").trim();
  const routingJson = String(zeile[4] || "").trim();
  const engineBErgebnis = String(zeile[6] || "").trim(); // Spalte G (Index 6)
  const spalteIInhalt = String(zeile[8] || "").trim();    // Spalte I (Index 8)

  if (!eindeutigeId || eindeutigeId === "-" || !datenJson || !routingJson || routingJson.startsWith("FEHLER")) {
    return false;
  }
  if (engineBErgebnis !== "[]") return false;
  return spalteIInhalt === "" || spalteIInhalt === "[]";
}

/**
 * Prüft den Routing-Status auf Großgruppe oder Standardlead
 */
function istZulaessigerRoutingStatus(status) {
  return status === "BIG_SHORT" || status === "BIG_LONG" || status === "STD_SHORT" || status === "STD_LONG";
}

/**
 * Bereitet Anfrageparameter für den Zeitversatz auf.
 * Bevorzugt validierte Datumsangaben aus kennzahlen vor dauerDaten (mit 6-Tage-Fallback).
 */
function extrahiereAnfrageParameterBAlt(datenObjekt, kennzahlen) {
  const personenDaten = datenObjekt.personen || {};
  const dauerDaten = datenObjekt.dauer || {};
  const ortDaten = datenObjekt.ort || {};

  const anzahlPersonen = Number(personenDaten.Pers) || Number(kennzahlen.pers) || 0;
  
  const startDatumText = String(kennzahlen.start || dauerDaten.start || "").trim();
  let endDatumText = String(kennzahlen.ende || dauerDaten.ende || "").trim();

  // 6-Tage-Fallback bei fehlendem Enddatum
  if (startDatumText && !endDatumText) {
    const basisStart = new Date(startDatumText + "T00:00:00Z");
    endDatumText = new Date(basisStart.getTime() + 6 * KONFIGURATION_B_ALT.MILLISEKUNDEN_PRO_TAG).toISOString().split("T")[0];
  }

  const zielZimmer = Number(kennzahlen.zimmer) || Math.ceil(anzahlPersonen / 2);
  const suchStadt = extractPureCityName(String(ortDaten.suchortText || ortDaten.Ort || ""));

  if (!suchStadt || !startDatumText || !endDatumText || anzahlPersonen <= 1) {
    return null;
  }

  const wunschStartMs = new Date(startDatumText + "T00:00:00Z").getTime();
  const wunschEndeMs = new Date(endDatumText + "T00:00:00Z").getTime();
  const aufenthaltsdauerMs = wunschEndeMs - wunschStartMs;

  if (aufenthaltsdauerMs <= 0) return null;

  return {
    suchStadt: suchStadt,
    personen: anzahlPersonen,
    zielZimmer: zielZimmer,
    wunschStartMs: wunschStartMs,
    aufenthaltsdauerMs: aufenthaltsdauerMs
  };
}

/**
 * Filtert aktive Wohnungen für die gesuchte Region
 */
function filtereWohnungenFuerRegion(wohnungen, suchStadt) {
  return wohnungen.filter(wohnung => {
    if (wohnung.istDauerhaftBelegt) return false;
    return wohnung.zulaessigeRegionen.includes(suchStadt) || suchStadt.includes(wohnung.zulaessigeRegionen);
  });
}

/**
 * Prüft Zeitversatz (+1 bis +14 Tage) und generiert Kombinationskandidaten
 * Enthält Timeout-Schutz durch Begrenzung des Pools vor dem Backtracking.
 */
function sucheKombiAlternativenImZeitfenster(regionWohnungen, anfrage, buchungen) {
  const gefundeneAlternativen = [];

  for (let tagVerschub = 1; tagVerschub <= KONFIGURATION_B_ALT.MAX_VERSCHIEBUNG_TAGE; tagVerschub++) {
    const testStartMs = anfrage.wunschStartMs + (tagVerschub * KONFIGURATION_B_ALT.MILLISEKUNDEN_PRO_TAG);
    const testEndeMs = testStartMs + anfrage.aufenthaltsdauerMs;

    const freieInSlot = ermittleFreieWohnungenInSlot(regionWohnungen, testStartMs, testEndeMs, buchungen);
    const summeMoeglicheBetten = freieInSlot.reduce((summe, w) => summe + w.maxP, 0);

    if (summeMoeglicheBetten < anfrage.personen) continue;

    // Timeout-Schutz: Vor dem Backtracking auf maximal konfigurierte Einheiten begrenzen
    const pool = freieInSlot.length > KONFIGURATION_B_ALT.MAX_KANDIDATEN_SLOT
      ? freieInSlot.slice(0, KONFIGURATION_B_ALT.MAX_KANDIDATEN_SLOT)
      : freieInSlot;

    const kombis = findeApartmentKombinationenBAlt(pool, anfrage.personen, anfrage.zielZimmer);
    for (const kombi of kombis) {
      if (kombi.teilangebot) continue;

      const penaltyScore = berechneKombiPenaltyScore(tagVerschub, kombi.restB, kombi.restZimmer, kombi.anzahl);
      gefundeneAlternativen.push({
        apts: kombi.apts,
        betten: kombi.gesamtBetten,
        schlafzimmer: kombi.gesamtZimmer,
        restB: kombi.restB,
        restZ: kombi.restZimmer,
        anzahl: kombi.anzahl,
        istAlternative: true,
        tageVerschoben: tagVerschub,
        altStart: new Date(testStartMs).toISOString().split("T")[0],
        altEnde: new Date(testEndeMs).toISOString().split("T")[0],
        penaltyScore: penaltyScore
      });
    }

    if (gefundeneAlternativen.length >= 4) break;
  }

  gefundeneAlternativen.sort((a, b) => a.penaltyScore - b.penaltyScore);
  return gefundeneAlternativen;
}

/**
 * Ermittelt kollisionsfreie Wohnungen in einem definierten Testfenster
 */
function ermittleFreieWohnungenInSlot(regionWohnungen, testStartMs, testEndeMs, buchungen) {
  const freie = [];

  for (const wohnung of regionWohnungen) {
    const existierendeBuchungen = buchungen[wohnung.name.toLowerCase()] || [];
    const hatKollision = existierendeBuchungen.some(b => b.startMillisekunden < testEndeMs && b.endeMillisekunden > testStartMs);

    if (!hatKollision) {
      freie.push({
        id: wohnung.kuerzel,
        maxP: wohnung.maximalePersonen,
        sz: wohnung.anzahlSchlafzimmer
      });
    }
  }

  return freie;
}

/**
 * Berechnet Penalty-Score für zeitversetzte Kombinationen
 */
function berechneKombiPenaltyScore(tageVerschoben, restBetten, restZimmer, anzahlWohnungen) {
  return (tageVerschoben * KONFIGURATION_B_ALT.GEWICHT_TAG_VERSCHIEBUNG) +
         (restBetten * KONFIGURATION_B_ALT.GEWICHT_LEERBETT) +
         (restZimmer * KONFIGURATION_B_ALT.GEWICHT_LEERZIMMER) +
         (anzahlWohnungen * KONFIGURATION_B_ALT.GEWICHT_APARTMENT_ANZAHL);
}

/**
 * Backtracking-Algorithmus zur Ermittlung valider Vollkombinationen
 */
function findeApartmentKombinationenBAlt(freieWohnungen, zielPersonen, zielZimmer) {
  const resultateVoll = [];
  const minZimmer = Number(zielZimmer) || 0;

  function backtrack(startIndex, aktuelleKombi, summeBetten, summeZimmer) {
    const bettenOk = summeBetten >= zielPersonen;
    const zimmerOk = minZimmer === 0 || summeZimmer >= minZimmer;

    if (bettenOk && zimmerOk) {
      resultateVoll.push({
        apts: aktuelleKombi.map(w => w.id),
        gesamtBetten: summeBetten,
        gesamtZimmer: summeZimmer,
        restB: summeBetten - zielPersonen,
        restZimmer: minZimmer > 0 ? (summeZimmer - minZimmer) : 0,
        anzahl: aktuelleKombi.length,
        teilangebot: false
      });
      return;
    }

    if (aktuelleKombi.length >= KONFIGURATION_B_ALT.MAX_KOMBI_GROESSE) return;

    for (let i = startIndex; i < freieWohnungen.length; i++) {
      aktuelleKombi.push(freieWohnungen[i]);
      backtrack(i + 1, aktuelleKombi, summeBetten + freieWohnungen[i].maxP, summeZimmer + freieWohnungen[i].sz);
      aktuelleKombi.pop();
    }
  }

  backtrack(0, [], 0, 0);
  return resultateVoll;
}

/**
 * Schreibt gefundene Alternativen ins Array und loggt das Ergebnis
 */
function protokolliereUndSetzeKombiAlternative(leadsDaten, zeilenIndex, alternativen, zeilennummer) {
  if (alternativen.length > 0) {
    const topAlternativen = alternativen.slice(0, 3);
    leadsDaten[zeilenIndex][8] = JSON.stringify(topAlternativen);
    Logger.log(`Zeile ${zeilennummer}: ${topAlternativen.length} Kombi-Alternative(n) in Spalte I hinterlegt (+${topAlternativen[0].tageVerschoben}d, Apts: ${topAlternativen[0].apts.join("+")}).`);
  } else {
    leadsDaten[zeilenIndex][8] = "[]";
    Logger.log(`Zeile ${zeilennummer}: Keine Kombi-Alternative innerhalb von ${KONFIGURATION_B_ALT.MAX_VERSCHIEBUNG_TAGE} Tagen gefunden.`);
  }
}

/**
 * Schreibt isoliert NUR Spalte I zurück und setzt die Zeilenhöhe
 * Behandelt den Edge-Case anzahlDatenZeilen === 1 vs. > 1
 */
function schreibeErgebnisseNachSpalteI(leadsTabelle, leadsDaten, anzahlDatenZeilen) {
  const spalteIWerte = leadsDaten.map(zeile => [zeile[8]]);
  leadsTabelle.getRange(2, 9, anzahlDatenZeilen, 1).setValues(spalteIWerte);

  if (anzahlDatenZeilen === 1) {
    leadsTabelle.setRowHeight(2, KONFIGURATION_B_ALT.ZEILENHOEHE_PIXEL);
  } else if (anzahlDatenZeilen > 1) {
    leadsTabelle.setRowHeights(2, anzahlDatenZeilen, KONFIGURATION_B_ALT.ZEILENHOEHE_PIXEL);
  }

  Logger.log("=== KOMBI-ALTERNATIVEN ERFOLGREICH IN SPALTE I GESPEICHERT ===");
}

/**
 * Header für Spalte I sicherstellen
 */
function stelleEngineBAltHeaderSicher(tabelle) {
  const kopfZelle = tabelle.getRange(1, 9);
  if (!kopfZelle.getValue() || kopfZelle.getValue() !== "Alternative Kombi (>6P)") {
    kopfZelle.setValue("Alternative Kombi (>6P)");
    kopfZelle.setFontWeight("bold");
    tabelle.getRange("I:I").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

/**
 * Bereinigt Kürzel von Sonderzeichen (identisch zu Engine A, A-Alt & B)
 */
function cleanShortCode(kuerzelText) {
  if (!kuerzelText) return "";
  return String(kuerzelText).replace(/[^a-zA-Z0-9_-]/g, "").trim();
}

/**
 * Wandelt Datumsangaben (Date-Objekt, DD.MM.YYYY oder YYYY-MM-DD) in ein einheitliches ISO-Format um
 */
function parseIsoDateString(eingabeDatum) {
  if (!eingabeDatum) return "";

  if (eingabeDatum instanceof Date) {
    return Utilities.formatDate(eingabeDatum, "UTC", "yyyy-MM-dd");
  }

  const datumString = String(eingabeDatum).trim();

  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(datumString)) {
    const teile = datumString.split(".");
    return `${teile[2]}-${teile[1].padStart(2, "0")}-${teile[0].padStart(2, "0")}`;
  }

  const isoTeil = datumString.split(/[T\s]/)[0];
  return /^\d{4}-\d{2}-\d{2}$/.test(isoTeil) ? isoTeil : "";
}

/**
 * Filtert Füllwörter, Postleitzahlen und Satzzeichen aus Stadtnamen
 */
function extractPureCityName(roherOrt) {
  if (!roherOrt) return "";
  let stadt = String(roherOrt).toLowerCase().trim();
  stadt = stadt.replace(/\b\d{4,5}\b/g, "").trim();
  stadt = stadt.replace(/\b(nähe|raum|bereich|stadt|umgebung|zentrum|mitte|ot|ortsteil|city|hbf|bahnhof)\b/gi, "").trim();
  if (stadt.includes(",")) {
    stadt = stadt.split(",")[0].trim();
  }
  stadt = stadt.replace(/[.\-\/\(\)]/g, " ");
  return stadt.replace(/\s+/g, " ").trim();
}

/**
 * Kapselt JSON.parse() sicher in try/catch
 */
function parseJsonSicher(jsonZeichenkette) {
  try {
    return JSON.parse(jsonZeichenkette);
  } catch (fehler) {
    return null;
  }
}