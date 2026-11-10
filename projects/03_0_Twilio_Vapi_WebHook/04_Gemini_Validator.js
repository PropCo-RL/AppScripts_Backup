// =================================================================
// SKRIPT: VALIDATOR & GESCHÄFTSLOGIK FÜR AUDIO-CALLS (MMZ-KOMPATIBEL)
// =================================================================
function runBusinessValidatorAudioMMZ() {
  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const letzteZeile = tabellenBlatt.getLastRow();
  if (letzteZeile <= 1) return;

  stelleValidatorHeaderSicher(tabellenBlatt);

  const anzahlZeilen = letzteZeile - 1;
  const datenBereich = tabellenBlatt.getRange(2, 10, anzahlZeilen, 2); // Spalte J (10) & K (11)
  const zeilenWerte = datenBereich.getValues();
  let aenderungenVorhanden = false;

  for (let i = 0; i < zeilenWerte.length; i++) {
    const rawJsonJ = String(zeilenWerte[i][0] || "").trim();
    const validiertK = String(zeilenWerte[i][1] || "").trim();

    if (!rawJsonJ || rawJsonJ.startsWith("Fehler") || validiertK !== "") {
      continue;
    }

    const geparstesJson = parseValidiereJson(rawJsonJ);
    if (!geparstesJson) {
      zeilenWerte[i][1] = JSON.stringify({ fehler: "JSON-Parsing fehlgeschlagen" });
      aenderungenVorhanden = true;
      continue;
    }

    const validiertesErgebnis = verarbeiteAudioCallValidierung(geparstesJson);
    zeilenWerte[i][1] = JSON.stringify(validiertesErgebnis);
    aenderungenVorhanden = true;
  }

  if (aenderungenVorhanden) {
    datenBereich.setValues(zeilenWerte);
    Logger.log("Business Validator für Audio-Calls erfolgreich abgeschlossen.");
  }
}

// =================================================================
// MODUL: VALIDIERUNG & ZUSAMMENFÜHRUNG
// =================================================================
function verarbeiteAudioCallValidierung(daten) {
  const personen = ermittlePersonenWerte(daten);
  const ort = daten.suchortText || daten.stadt || "";
  const dauer = ermittleDauerWerte(daten);
  const zitate = daten.zitate || {};
  const hatZitate = pruefeZitateGueltigkeit(zitate);

  const vipStatus = berechneVipStatus(personen.berechnung, dauer.maxTage);
  const confidence = berechneConfidenceScores(daten, hatZitate);

  return {
    vip: vipStatus,
    ort: { Ort: ort, suchortText: ort },
    personen: {
      Pers: personen.pers,
      max: personen.maxPers,
      aufteilung: daten.personenData?.aufteilung || "Standard",
      zimmer: daten.personenData?.zimmer || ""
    },
    dauer: dauer,
    spezial: daten.spezialData || {},
    confidence: confidence,
    details: {
      zitate: zitate,
      istStaffelAnfrage: Boolean(daten.istStaffelAnfrage),
      abweichungen: daten.abweichungen_zu_spalte_e || daten.abweichungen || {}
    }
  };
}

// =================================================================
// MODUL: BERECHNUNGEN (STRIKTES EARLY RETURN)
// =================================================================
function ermittlePersonenWerte(daten) {
  let pers = Number(daten.personenData?.Pers);
  if (!pers || isNaN(pers)) {
    const rawAlt = String(daten.anzahl_personen || "");
    const match = rawAlt.match(/\d+/);
    pers = match ? Number(match[0]) : 1;
  }

  const maxPers = Number(daten.personenData?.max) || pers;
  return { pers: pers, maxPers: maxPers, berechnung: Math.max(pers, maxPers) };
}

function ermittleDauerWerte(daten) {
  const zeitData = daten.zeitraumData || {};
  let startStr = zeitData.start || "";
  let endeStr = zeitData.ende || "";
  const hatVerlaengerung = Boolean(zeitData.verlaengerung);

  let tageExakt = berechneTageDifferenz(startStr, endeStr);
  let finalMaxTage = Math.max(Number(zeitData.maxTage) || 0, tageExakt);

  if (hatVerlaengerung && finalMaxTage < 30) {
    finalMaxTage = 30;
  }

  const startD = parseISODate(startStr);
  const endeD = parseISODate(endeStr);

  return {
    Z: buildZVorschau(startD, endeD, tageExakt),
    tage: tageExakt,
    maxTage: finalMaxTage,
    abrechnung: zeitData.abrechnung || null,
    verlaengerung: hatVerlaengerung,
    muster: zeitData.muster || "durchgehend",
    start: startStr,
    ende: endeStr,
    info: zeitData.info || String(daten.zeitraum || "")
  };
}

function berechneVipStatus(berechnungsPersonen, finalMaxTage) {
  if (berechnungsPersonen >= 3 && finalMaxTage >= 120) return "VIP2";
  if (berechnungsPersonen >= 3 && finalMaxTage >= 30) return "VIP1";
  return "-";
}

function berechneConfidenceScores(daten, hatZitate) {
  let kiScore = typeof daten.kiConfidenceScore === "number" ? daten.kiConfidenceScore : 85;
  let kiStatus = "OK";

  if (!hatZitate && daten.zitate) {
    kiScore = Math.min(kiScore, 50);
    kiStatus = "UNSICHER: ZITAT_FEHLT";
  } else if (kiScore < 85) {
    kiStatus = "UNSICHER";
  }

  const gastScore = typeof daten.gastConfidenceScore === "number" ? daten.gastConfidenceScore : 100;
  const gastStatus = daten.gastIstVage ? `Vage (${daten.vageGruende || 'Unpräzise'})` : "Verbindlich";

  return {
    ki: `${kiScore}% | ${kiStatus}`,
    gast: `${gastScore}% | ${gastStatus}`
  };
}

// =================================================================
// MODUL: HILFSFUNKTIONEN
// =================================================================
function berechneTageDifferenz(startStr, endeStr) {
  if (!startStr || !endeStr) return 0;
  const s = parseISODate(startStr);
  const e = parseISODate(endeStr);
  if (!s || !e || e < s) return 0;

  const diffTage = Math.round((e.getTime() - s.getTime()) / (1000 * 60 * 60 * 24));
  return diffTage === 0 ? 1 : diffTage;
}

function pruefeZitateGueltigkeit(zitate) {
  if (!zitate) return false;
  return Boolean(
    zitate.personenFundstelle && zitate.personenFundstelle.trim().length > 0 &&
    zitate.zeitraumFundstelle && zitate.zeitraumFundstelle.trim().length > 0
  );
}

function parseISODate(str) {
  if (!str) return null;
  const match = String(str).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
}

function buildZVorschau(startD, endeD, tage) {
  if (!startD || !endeD || tage <= 0) return "";
  const sTag = String(startD.getUTCDate()).padStart(2, '0');
  const sMon = String(startD.getUTCMonth() + 1).padStart(2, '0');
  const eTag = String(endeD.getUTCDate()).padStart(2, '0');
  const eMon = String(endeD.getUTCMonth() + 1).padStart(2, '0');
  return `${sTag}.${sMon}. - ${eTag}.${eMon}. (${tage}T)`;
}

function parseValidiereJson(text) {
  try {
    return JSON.parse(text);
  } catch (e) {
    return null;
  }
}

function stelleValidatorHeaderSicher(blatt) {
  const headerK = blatt.getRange(1, 11);
  if (!headerK.getValue()) {
    headerK.setValue("Validiert (Gesamt JSON)");
    headerK.setFontWeight("bold");
    blatt.getRange("K:K").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}