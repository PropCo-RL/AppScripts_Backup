// =================================================================
// STUFE 2b: ENGINE B - MULTI-APARTMENT FINDER FÜR TAB "MMZ"
// Liest Spalte D (Daten), Spalte E (Routing) & Spalte F (Engine A)
// Schreibt Kombinations-Optionen AUSSCHLIESSLICH in Spalte G (Spalte 7)
// =================================================================

const KONFIGURATION_B = {
  INVENTORY_SPREADSHEET_ID: '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70',
  TABELLE_LEADS: 'MMZ',
  TABELLE_INVENTAR: 'M',
  TABELLE_KALENDER: 'cal',
  MAX_KOMBI_GROESSE: 4,
  MAX_KANDIDATEN_KOMBI: 20,
  MAX_LUECKE_TAGE: 14,
  ZEILENHOEHE_PIXEL: 21,
  MILLISEKUNDEN_PRO_TAG: 86400000
};

/**
 * Zentrale Orchestrator-Funktion (Trigger-Einstiegspunkt)
 */
function runApartmentMatchingEngineB_MMZ() {
  Logger.log("=== START ENGINE B MULTI-MATCHING FÜR MMZ ===");
  const aktiveArbeitsmappe = SpreadsheetApp.getActiveSpreadsheet();
  const leadsTabelle = aktiveArbeitsmappe.getSheetByName(KONFIGURATION_B.TABELLE_LEADS);
  if (!leadsTabelle) {
    Logger.log("FEHLER: Tabellenblatt MMZ nicht gefunden!");
    return;
  }

  const anzahlZeilen = leadsTabelle.getLastRow();
  if (anzahlZeilen <= 1) return;

  const inventarMappe = ladeInventarSpreadsheetB(KONFIGURATION_B.INVENTORY_SPREADSHEET_ID);
  if (!inventarMappe) return;

  const inventarTabelle = inventarMappe.getSheetByName(KONFIGURATION_B.TABELLE_INVENTAR);
  const kalenderTabelle = inventarMappe.getSheetByName(KONFIGURATION_B.TABELLE_KALENDER);
  if (!inventarTabelle || !kalenderTabelle) {
    Logger.log("FEHLER: Ein Tabellenblatt fehlt (M oder cal)!");
    return;
  }

  stelleEngineBHeaderSicher(leadsTabelle);
  const wohnungsListe = ladeWohnungenAusInventarB(inventarTabelle);
  const buchungenKarte = ladeBuchungenAusKalenderB(kalenderTabelle);

  verarbeiteKombiLeadsUndSchreibeErgebnisse(leadsTabelle, anzahlZeilen, wohnungsListe, buchungenKarte);
  Logger.log("=== ENGINE B ABGESCHLOSSEN ===");
}

/**
 * Öffnet das externe Inventar-Spreadsheet isoliert mit Fehlerabfang
 */
function ladeInventarSpreadsheetB(tabellenId) {
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
function ladeWohnungenAusInventarB(tabelle) {
  const rohWerte = tabelle.getDataRange().getValues();
  const wohnungen = [];

  for (let zeile = 1; zeile < rohWerte.length; zeile++) {
    const wohnungsName = String(rohWerte[zeile][0] || "").trim();
    if (!wohnungsName) continue;

    // Spalte AJ (Index 35): Ausschluss-Prüfung
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
function ladeBuchungenAusKalenderB(tabelle) {
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
 * Iteriert über Leads, prüft Kaskaden-Bedingung und schreibt Kombinationen isoliert in Spalte G
 */
function verarbeiteKombiLeadsUndSchreibeErgebnisse(leadsTabelle, zeilenAnzahl, wohnungen, buchungen) {
  const anzahlDatenZeilen = zeilenAnzahl - 1;
  const leadsBereich = leadsTabelle.getRange(2, 1, anzahlDatenZeilen, 7);
  const leadsDaten = leadsBereich.getValues();
  let aenderungenVorhanden = false;

  const heute = new Date();
  const heuteUtcMs = new Date(Date.UTC(heute.getFullYear(), heute.getMonth(), heute.getDate())).getTime();

  for (let index = 0; index < anzahlDatenZeilen; index++) {
    const aktuelleZeile = index + 2;
    const leadSpalten = leadsDaten[index];

    if (!pruefeLeadGueltigkeitB(leadSpalten)) continue;

    const routingObjekt = parseJsonSicher(leadSpalten[4]);
    if (!routingObjekt || !pruefeRoutingKaskade(routingObjekt.status, leadSpalten[5])) continue;

    const datenObjekt = parseJsonSicher(leadSpalten[3]);
    if (!datenObjekt) {
      Logger.log(`Zeile ${aktuelleZeile}: JSON Parse Fehler in Spalte D`);
      continue;
    }

    const anfrage = extrahiereAnfrageParameterB(datenObjekt, routingObjekt.kennzahlen || {}, heuteUtcMs);
    if (!anfrage) continue;

    const regionWohnungen = wohnungEnsemble(wohnungen, anfrage.suchStadt);
    const freieWohnungen = filtereFreieWohnungenInRegion(regionWohnungen, anfrage, buchungen);
    Logger.log(`Zeile ${aktuelleZeile} (${routingObjekt.status}, ${anfrage.personen}P / ${anfrage.zielZimmer}Z): ${freieWohnungen.length} freie Einheiten gefunden.`);

    // Performance-Deckelung auf beste Kandidaten vor dem Backtracking
    const priorisierteWohnungen = begrenzKandidatenFuerKombinatorik(freieWohnungen);

    const gefundeneKombis = findeApartmentKombinationen(priorisierteWohnungen, anfrage.personen, anfrage.zielZimmer);
    sortiereApartmentKombinationen(gefundeneKombis);

    protokolliereUndSetzeKombiTreffer(leadsDaten, index, gefundeneKombis, aktuelleZeile);
    aenderungenVorhanden = true;
  }

  if (aenderungenVorhanden) {
    schreibeErgebnisseNachSpalteG(leadsTabelle, leadsDaten, anzahlDatenZeilen);
  } else {
    Logger.log("Keine Zeilen zur Bearbeitung für Engine B vorhanden.");
  }
}

/**
 * Schreibt gesammelt NUR Spalte G zurück und korrigiert die Zeilenhöhe
 */
function schreibeErgebnisseNachSpalteG(leadsTabelle, leadsDaten, anzahlDatenZeilen) {
  const spalteGWerte = leadsDaten.map(zeile => [zeile[6]]);
  leadsTabelle.getRange(2, 7, anzahlDatenZeilen, 1).setValues(spalteGWerte);
  leadsTabelle.setRowHeights(2, anzahlDatenZeilen, KONFIGURATION_B.ZEILENHOEHE_PIXEL);
  Logger.log("=== ENGINE B ERFOLGREICH IN SPALTE G GESCHRIEBEN ===");
}

/**
 * Begrenzt Kandidaten für Backtracking auf die wirtschaftlichsten Einheiten
 */
function begrenzKandidatenFuerKombinatorik(wohnungen) {
  if (wohnungen.length <= KONFIGURATION_B.MAX_KANDIDATEN_KOMBI) {
    return wohnungen;
  }
  // Bevorzuge Einheiten mit bestem Lücken-Score (wenig Leerstand)
  const kopie = wohnungen.slice();
  kopie.sort((a, b) => a.score - b.score);
  return kopie.slice(0, KONFIGURATION_B.MAX_KANDIDATEN_KOMBI);
}

/**
 * Validiert die Zeilengrundlage für Stufe 2b
 */
function pruefeLeadGueltigkeitB(zeile) {
  const eindeutigeId = String(zeile[1] || "").trim();
  const datenJson = String(zeile[3] || "").trim();
  const routingJson = String(zeile[4] || "").trim();
  const spalteGInhalt = String(zeile[6] || "").trim();

  if (!eindeutigeId || eindeutigeId === "-" || !datenJson || !routingJson || routingJson.startsWith("FEHLER")) {
    return false;
  }
  return spalteGInhalt === "" || spalteGInhalt === "[]";
}

/**
 * Prüft Kaskaden-Bedingung: Nur Großgruppen oder Standardleads ohne Single-Apartment-Treffer
 */
function pruefeRoutingKaskade(routingStatus, engineAErgebnis) {
  const istGrossGruppe = (routingStatus === "BIG_SHORT" || routingStatus === "BIG_LONG");
  const singleErgebnisText = String(engineAErgebnis || "").trim();
  const istStandardOhneSingleTreffer = (routingStatus === "STD_SHORT" || routingStatus === "STD_LONG") &&
                                       (singleErgebnisText === "[]" || singleErgebnisText === "");
  return istGrossGruppe || istStandardOhneSingleTreffer;
}

/**
 * Bereitet Anfrageparameter auf, inklusive 6-Tage-Fallback und Vorlaufzeitberechnung
 */
function extrahiereAnfrageParameterB(datenObjekt, kennzahlen, heuteUtcMs) {
  const personenDaten = datenObjekt.personen || {};
  const dauerDaten = datenObjekt.dauer || {};
  const ortDaten = datenObjekt.ort || {};

  const anzahlPersonen = Number(personenDaten.Pers) || Number(kennzahlen.pers) || 0;
  const startDatumText = String(kennzahlen.start || dauerDaten.start || "").trim();
  let endDatumText = String(kennzahlen.ende || dauerDaten.ende || "").trim();

  // 6-Tage-Fallback bei fehlendem Enddatum
  if (startDatumText && !endDatumText) {
    const basisStart = new Date(startDatumText + "T00:00:00Z");
    endDatumText = new Date(basisStart.getTime() + 6 * KONFIGURATION_B.MILLISEKUNDEN_PRO_TAG).toISOString().split("T")[0];
  }

  const zielZimmer = Number(kennzahlen.zimmer) || Math.ceil(anzahlPersonen / 2);
  const suchStadt = extractPureCityName(String(ortDaten.suchortText || ortDaten.Ort || ""));

  if (!suchStadt || !startDatumText || !endDatumText || anzahlPersonen <= 1) {
    return null;
  }

  const anfrageStartMs = new Date(startDatumText + "T00:00:00Z").getTime();
  const anfrageEndeMs = new Date(endDatumText + "T00:00:00Z").getTime();

  if (anfrageEndeMs <= anfrageStartMs) return null;

  const tageBisAnreise = Math.round((anfrageStartMs - heuteUtcMs) / KONFIGURATION_B.MILLISEKUNDEN_PRO_TAG);

  return {
    suchStadt: suchStadt,
    personen: anzahlPersonen,
    zielZimmer: zielZimmer,
    startMs: anfrageStartMs,
    endeMs: anfrageEndeMs,
    tageBisAnreise: tageBisAnreise
  };
}

/**
 * Filtert Apartments vorab nach Region
 */
function wohnungEnsemble(wohnungen, suchStadt) {
  return wohnungen.filter(wohnung => {
    if (wohnung.istDauerhaftBelegt) return false;
    return wohnung.zulaessigeRegionen.includes(suchStadt) || suchStadt.includes(wohnung.zulaessigeRegionen);
  });
}

/**
 * Ermittelt freie Einheiten und berechnet deren Belegungs-/Lücken-Score
 */
function filtereFreieWohnungenInRegion(regionWohnungen, anfrage, buchungen) {
  const verfuegbareEinheiten = [];

  for (const wohnung of regionWohnungen) {
    const kollision = pruefeKollisionUndLuecke(wohnung.name, anfrage.startMs, anfrage.endeMs, buchungen);
    if (kollision.istBelegt) continue;

    const lueckenScore = ermittleKombiLueckenScore(kollision.letzteAbreiseMs, anfrage.startMs, anfrage.tageBisAnreise);

    verfuegbareEinheiten.push({
      id: wohnung.kuerzel,
      maxP: wohnung.maximalePersonen,
      sz: wohnung.anzahlSchlafzimmer,
      score: lueckenScore
    });
  }

  return verfuegbareEinheiten;
}

/**
 * Prüft Buchungskollisionen und sucht letzte Abreise
 */
function pruefeKollisionUndLuecke(wohnungsName, startMs, endeMs, buchungen) {
  const existierendeBuchungen = buchungen[wohnungsName.toLowerCase()] || [];
  let istBelegt = false;
  let letzteAbreiseMs = null;

  for (const buchung of existierendeBuchungen) {
    if (buchung.startMillisekunden < endeMs && buchung.endeMillisekunden > startMs) {
      istBelegt = true;
      break;
    }
    if (buchung.endeMillisekunden <= startMs) {
      if (letzteAbreiseMs === null || buchung.endeMillisekunden > letzteAbreiseMs) {
        letzteAbreiseMs = buchung.endeMillisekunden;
      }
    }
  }

  return { istBelegt: istBelegt, letzteAbreiseMs: letzteAbreiseMs };
}

/**
 * Berechnet Lücken-Score für eine Wohnung
 */
function ermittleKombiLueckenScore(letzteAbreiseMs, anfrageStartMs, tageBisAnreise) {
  let leerTageVorher = null;
  if (letzteAbreiseMs !== null) {
    const diffTage = Math.max(0, Math.round((anfrageStartMs - letzteAbreiseMs) / KONFIGURATION_B.MILLISEKUNDEN_PRO_TAG));
    if (diffTage <= KONFIGURATION_B.MAX_LUECKE_TAGE) {
      leerTageVorher = diffTage;
    }
  }

  if (tageBisAnreise <= 5) {
    return (leerTageVorher === null) ? -1 : leerTageVorher;
  }
  return (leerTageVorher === null) ? 999 : leerTageVorher;
}

/**
 * Backtracking-Algorithmus für Kombinationen (Vollabdeckung + bestes Teilangebot)
 */
function findeApartmentKombinationen(verfuegbareApts, zielPersonen, zielZimmer) {
  const resultateVoll = [];
  let besteTeilKombi = null;
  const minZimmer = Number(zielZimmer) || 0;

  function backtrack(startIndex, aktuelleKombi, summeBetten, summeZimmer, summeScore) {
    const bettenErreicht = summeBetten >= zielPersonen;
    const zimmerErreicht = minZimmer === 0 || summeZimmer >= minZimmer;

    if (bettenErreicht && zimmerErreicht) {
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

    if (aktuelleKombi.length > 0) {
      besteTeilKombi = aktualisiereBestesTeilangebot(besteTeilKombi, aktuelleKombi, summeBetten, summeZimmer, summeScore, zielPersonen);
    }

    if (aktuelleKombi.length >= KONFIGURATION_B.MAX_KOMBI_GROESSE) return;

    for (let i = startIndex; i < verfuegbareApts.length; i++) {
      aktuelleKombi.push(verfuegbareApts[i]);
      backtrack(
        i + 1,
        aktuelleKombi,
        summeBetten + verfuegbareApts[i].maxP,
        summeZimmer + verfuegbareApts[i].sz,
        summeScore + verfuegbareApts[i].score
      );
      aktuelleKombi.pop();
    }
  }

  backtrack(0, [], 0, 0, 0);

  if (resultateVoll.length > 0) return resultateVoll;
  if (besteTeilKombi && besteTeilKombi.gesamtBetten >= 2) return [besteTeilKombi];

  return [];
}

/**
 * Ermittelt das qualitativ beste Teilangebot bei fehlender Vollabdeckung
 */
function aktualisiereBestesTeilangebot(aktuellBeste, kombi, betten, zimmer, score, zielPersonen) {
  const istBesser = !aktuellBeste ||
    betten > aktuellBeste.gesamtBetten ||
    (betten === aktuellBeste.gesamtBetten && zimmer > aktuellBeste.gesamtZimmer) ||
    (betten === aktuellBeste.gesamtBetten && zimmer === aktuellBeste.gesamtZimmer && score < aktuellBeste.summeScore);

  if (!istBesser) return aktuellBeste;

  return {
    apts: kombi.map(a => a.id),
    gesamtBetten: betten,
    gesamtZimmer: zimmer,
    restB: 0,
    restZimmer: 0,
    anzahl: kombi.length,
    summeScore: score,
    teilangebot: true,
    fehlendePersonen: Math.max(0, zielPersonen - betten)
  };
}

/**
 * Sortiert gefundene Kombinationen nach Wirtschaftlichkeit und Passgenauigkeit
 */
function sortiereApartmentKombinationen(kombis) {
  kombis.sort((a, b) => {
    if (a.teilangebot !== b.teilangebot) return a.teilangebot ? 1 : -1;
    if (a.teilangebot && b.teilangebot && a.gesamtBetten !== b.gesamtBetten) {
      return b.gesamtBetten - a.gesamtBetten;
    }
    if (a.restZimmer !== b.restZimmer) return a.restZimmer - b.restZimmer;
    if (a.restB !== b.restB) return a.restB - b.restB;
    if (a.anzahl !== b.anzahl) return a.anzahl - b.anzahl;
    return a.summeScore - b.summeScore;
  });
}

/**
 * Schreibt Top-Kombinationen ins Array und loggt das Ergebnis
 */
function protokolliereUndSetzeKombiTreffer(leadsDaten, zeilenIndex, kombis, zeilennummer) {
  if (kombis.length > 0) {
    const topKombis = kombis.slice(0, 3).map(k => ({
      apts: k.apts,
      betten: k.gesamtBetten,
      schlafzimmer: k.gesamtZimmer,
      restB: k.restB,
      restZ: k.restZimmer,
      anzahl: k.anzahl,
      teilangebot: Boolean(k.teilangebot),
      fehlendePersonen: k.fehlendePersonen || 0
    }));

    leadsDaten[zeilenIndex][6] = JSON.stringify(topKombis);
    Logger.log(`Zeile ${zeilennummer}: ${topKombis.length} Kombi-Option(en) eingetragen (Teilangebot: ${topKombis[0].teilangebot}).`);
  } else {
    leadsDaten[zeilenIndex][6] = "[]";
    Logger.log(`Zeile ${zeilennummer}: Keine passende Apartment-Kombination gefunden.`);
  }
}

/**
 * Header für Spalte G sicherstellen
 */
function stelleEngineBHeaderSicher(tabelle) {
  const kopfZelle = tabelle.getRange(1, 7);
  if (!kopfZelle.getValue()) {
    kopfZelle.setValue("Optionen Kombi (>6P)");
    kopfZelle.setFontWeight("bold");
    tabelle.getRange("G:G").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

/**
 * Bereinigt Kürzel von Steuer- und Sonderzeichen
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