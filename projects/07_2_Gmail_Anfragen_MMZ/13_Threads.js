// =================================================================
// 13_ACTION WORKER (MMZ) - MODULAR ARCHITECTURE
// - Liest Live-State aus SPALTE Q (17), Lead (B) und Draft (J)
// - Führt atomare Handler aus (BUCHUNG -> RE, VERLAENGERUNG -> VL)
// - Berechnet Enddaten bei Zyklusangaben ("14 Tage") automatisch
// - Idempotente Ausführung & Bereinigung von 'offeneAktionen'
// - Spiegelt den bereinigten State zurück nach MMZ (Q) und Thread Archive (H)
// =================================================================

const WORKER_KONFIGURATION = {
  TABELLE_HAUPT: 'MMZ',
  TABELLE_ARCHIV: 'Thread Archive',
  TABELLE_RE: 'RE',
  TABELLE_VL: 'VL',
  SPALTE_LEAD: 2,       // B: Lead JSON
  SPALTE_DRAFT: 10,     // J: Draft Status JSON
  SPALTE_STATE: 17      // Q: INBOUND_DISPATCH (State JSON)
};

// -----------------------------------------------------------------
// 1. ZENTRALER ORCHESTRATOR
// -----------------------------------------------------------------
function runActionWorkerMMZ() {
  const arbeitsmappe = SpreadsheetApp.getActiveSpreadsheet();
  const hauptBlatt = arbeitsmappe.getSheetByName(WORKER_KONFIGURATION.TABELLE_HAUPT);
  if (!hauptBlatt) return;

  const anzahlZeilen = hauptBlatt.getLastRow() - 1;
  if (anzahlZeilen <= 0) return;

  const archivBlatt = arbeitsmappe.getSheetByName(WORKER_KONFIGURATION.TABELLE_ARCHIV);
  const batchDaten = ladeWorkerBatchDaten(hauptBlatt, anzahlZeilen);

  const verarbeitungsErgebnis = verarbeiteWorkerBatch(batchDaten, arbeitsmappe, archivBlatt);

  if (verarbeitungsErgebnis.aenderungenVorhanden) {
    schreibeWorkerStatesZurueck(hauptBlatt, verarbeitungsErgebnis.neueStateWerte, anzahlZeilen);
  }
}

// -----------------------------------------------------------------
// 2. SHEET-IO & BATCHING
// -----------------------------------------------------------------
function ladeWorkerBatchDaten(hauptBlatt, anzahlZeilen) {
  return {
    leads: hauptBlatt.getRange(2, WORKER_KONFIGURATION.SPALTE_LEAD, anzahlZeilen, 1).getValues(),
    drafts: hauptBlatt.getRange(2, WORKER_KONFIGURATION.SPALTE_DRAFT, anzahlZeilen, 1).getValues(),
    states: hauptBlatt.getRange(2, WORKER_KONFIGURATION.SPALTE_STATE, anzahlZeilen, 1).getValues()
  };
}

function schreibeWorkerStatesZurueck(hauptBlatt, neueStateWerte, anzahlZeilen) {
  const zielBereich = hauptBlatt.getRange(2, WORKER_KONFIGURATION.SPALTE_STATE, anzahlZeilen, 1);
  zielBereich.setValues(neueStateWerte);
  hauptBlatt.setRowHeights(2, anzahlZeilen, 21);
  SpreadsheetApp.flush();
  Logger.log("Action Worker: Alle Status-Updates erfolgreich nach Spalte Q geschrieben.");
}

// -----------------------------------------------------------------
// 3. BATCH-SCHLEIFE & DISPATCH-EXECUTION
// -----------------------------------------------------------------
function verarbeiteWorkerBatch(batchDaten, arbeitsmappe, archivBlatt) {
  const anzahlZeilen = batchDaten.leads.length;
  const neueStateWerte = batchDaten.states.slice();
  let aenderungenVorhanden = false;

  for (let index = 0; index < anzahlZeilen; index++) {
    const roherState = String(batchDaten.states[index][0] || "").trim();
    if (!roherState || roherState === "-" || !roherState.startsWith("{")) continue;

    const dispatchDaten = parseJsonSicher(roherState, null);
    if (!dispatchDaten || !Array.isArray(dispatchDaten.offeneAktionen) || dispatchDaten.offeneAktionen.length === 0) {
      continue;
    }

    const leadObjekt = parseJsonSicher(batchDaten.leads[index][0], {});
    const draftObjekt = parseJsonSicher(batchDaten.drafts[index][0], {});
    const ausfuehrungsKontext = { leadObjekt: leadObjekt, draftObjekt: draftObjekt, arbeitsmappe: arbeitsmappe };

    const wurdeAusgefuehrt = fuehreOffeneAktionenAus(dispatchDaten, ausfuehrungsKontext);
    if (!wurdeAusgefuehrt) continue;

    const finalJsonText = JSON.stringify(dispatchDaten);
    neueStateWerte[index][0] = finalJsonText;
    aenderungenVorhanden = true;

    if (dispatchDaten.letzteMsgId) {
      spiegleStateInsArchiv(archivBlatt, dispatchDaten.letzteMsgId, finalJsonText);
    }
  }

  return { aenderungenVorhanden: aenderungenVorhanden, neueStateWerte: neueStateWerte };
}

function fuehreOffeneAktionenAus(dispatchDaten, kontext) {
  const erledigtMenge = new Set(dispatchDaten.erledigteAktionen || []);
  let mindestensEineAktionErledigt = false;

  for (const aktion of dispatchDaten.offeneAktionen) {
    if (erledigtMenge.has(aktion.actionId)) continue;

    const zustandigerHandler = INTENT_HANDLERS[aktion.typ];
    if (typeof zustandigerHandler !== "function") {
      Logger.log(`Kein Handler registriert für Intent-Typ: ${aktion.typ}`);
      continue;
    }

    const ausfuehrungErfolgreich = zustandigerHandler(aktion, dispatchDaten, kontext);
    if (ausfuehrungErfolgreich) {
      erledigtMenge.add(aktion.actionId);
      mindestensEineAktionErledigt = true;
    }
  }

  if (mindestensEineAktionErledigt) {
    dispatchDaten.erledigteAktionen = Array.from(erledigtMenge);
    dispatchDaten.offeneAktionen = dispatchDaten.offeneAktionen.filter(a => !erledigtMenge.has(a.actionId));
  }

  return mindestensEineAktionErledigt;
}

// -----------------------------------------------------------------
// 4. HANDLER-REGISTRY & HANDLER-IMPLEMENTIERUNGEN
// -----------------------------------------------------------------
const INTENT_HANDLERS = {
  "BUCHUNG": handleBuchungInRE,
  "VERLAENGERUNG": handleVerlaengerungInVL,
  "UMZUG": handleWohnungsUmzug,
  "SUPPORT_MANGEL": handleSupportTicket,
  "SUPPORT": handleSupportTicket,
  "WOHNUNGSGEBERBESTAETIGUNG": handleDokumentenAnfrage,
  "DOKUMENTE": handleDokumentenAnfrage,
  "EXPANSION": handleExpansionsAnfrage,
  "PREISVERHANDLUNG": handlePreisverhandlungLog
};

function handleBuchungInRE(aktion, dispatchDaten, kontext) {
  if (aktion.status !== "BEREIT_FUER_RE" || aktion.confidence !== "DEFINITIV") {
    return false;
  }

  const blattRE = kontext.arbeitsmappe.getSheetByName(WORKER_KONFIGURATION.TABELLE_RE);
  if (!blattRE) return false;

  const lead = kontext.leadObjekt;
  const draft = kontext.draftObjekt;
  const parameter = aktion.parameter || {};
  const org = dispatchDaten.organisation || {};
  const rechnungsAdresse = org.rechnungsanschriftDefault || {};

  const anreiseIso = parameter.anreise || lead.anreise || draft.anreise || "";
  const abreiseIso = berechneAbreiseDatum(anreiseIso, parameter.abreise || lead.abreise || draft.abreise, parameter.zusatzTage);
  const apartmentId = aktion.apartmentId || draft.apartmentId || "";
  const zyklus = parameter.abrechnungsZyklus || "14 Tage";
  const preisProNacht = parameter.gegenangebotPreis || parameter.preisProNacht || draft.preisProNacht || "";

  const neueZeile = baueRechnungsZeilenArray({
    telefon: lead.telefon || "",
    firma: org.firmenname || lead.firma || "",
    name: rechnungsAdresse.empfaenger || lead.name || lead.gastName || "",
    strasse: rechnungsAdresse.strasse || "",
    plz: rechnungsAdresse.plz || "",
    stadt: rechnungsAdresse.stadt || lead.stadt || "",
    land: rechnungsAdresse.land || "Deutschland",
    ustId: org.ustId || "",
    email: ermittleBesteEmail(dispatchDaten, lead),
    anreiseDE: formatiereDatumDE(anreiseIso),
    abreiseDE: formatiereDatumDE(abreiseIso),
    zyklus: zyklus,
    apartmentId: apartmentId,
    preisProNacht: preisProNacht
  });

  blattRE.appendRow(neueZeile);
  Logger.log(`RE: Buchung für Apartment '${apartmentId}' angelegt (${anreiseIso} bis ${abreiseIso}).`);

  aktualisiereWohnungsStatusNachBuchung(dispatchDaten, apartmentId, zyklus, anreiseIso, abreiseIso, preisProNacht);
  return true;
}

function handleVerlaengerungInVL(aktion, dispatchDaten, kontext) {
  if (aktion.status !== "BEREIT_FUER_VL" && aktion.status !== "BEREIT_FUER_RE") {
    return false;
  }

  const blattVL = kontext.arbeitsmappe.getSheetByName(WORKER_KONFIGURATION.TABELLE_VL);
  if (!blattVL) return false;

  const lead = kontext.leadObjekt;
  const draft = kontext.draftObjekt;
  const parameter = aktion.parameter || {};
  const apartmentId = aktion.apartmentId || draft.apartmentId || "";

  const bisherigeAbreise = (dispatchDaten.wohnungsStatus && dispatchDaten.wohnungsStatus[apartmentId])
    ? dispatchDaten.wohnungsStatus[apartmentId].abreise
    : (lead.abreise || draft.abreise || "");

  const anreiseIso = parameter.anreise || bisherigeAbreise;
  const abreiseIso = berechneAbreiseDatum(anreiseIso, parameter.abreise, parameter.zusatzTage);
  const zyklus = parameter.abrechnungsZyklus || "14 Tage";
  const preisProNacht = parameter.gegenangebotPreis || parameter.preisProNacht || draft.preisProNacht || "";

  const neueZeile = baueVerlaengerungsZeilenArray({
    telefon: lead.telefon || "",
    firma: (dispatchDaten.organisation && dispatchDaten.organisation.firmenname) || lead.firma || "",
    name: (dispatchDaten.organisation && dispatchDaten.organisation.rechnungsanschriftDefault && dispatchDaten.organisation.rechnungsanschriftDefault.empfaenger) || lead.name || "",
    strasse: (dispatchDaten.organisation && dispatchDaten.organisation.rechnungsanschriftDefault && dispatchDaten.organisation.rechnungsanschriftDefault.strasse) || "",
    plz: (dispatchDaten.organisation && dispatchDaten.organisation.rechnungsanschriftDefault && dispatchDaten.organisation.rechnungsanschriftDefault.plz) || "",
    stadt: (dispatchDaten.organisation && dispatchDaten.organisation.rechnungsanschriftDefault && dispatchDaten.organisation.rechnungsanschriftDefault.stadt) || "",
    land: "Deutschland",
    ustId: (dispatchDaten.organisation && dispatchDaten.organisation.ustId) || "",
    email: ermittleBesteEmail(dispatchDaten, lead),
    anreiseDE: formatiereDatumDE(anreiseIso),
    abreiseDE: formatiereDatumDE(abreiseIso),
    zyklus: zyklus,
    apartmentId: apartmentId,
    preisProNacht: preisProNacht
  });

  blattVL.appendRow(neueZeile);
  Logger.log(`VL: Verlängerung für Apartment '${apartmentId}' angelegt (bis ${abreiseIso}).`);

  if (apartmentId && dispatchDaten.wohnungsStatus && dispatchDaten.wohnungsStatus[apartmentId]) {
    dispatchDaten.wohnungsStatus[apartmentId].lebenszyklus = "VERLAENGERT";
    if (abreiseIso) dispatchDaten.wohnungsStatus[apartmentId].abreise = abreiseIso;
  }

  return true;
}

function handleWohnungsUmzug(aktion, dispatchDaten) {
  const parameter = aktion.parameter || {};
  const altesApt = parameter.alteWohnung || aktion.apartmentId;
  const neuesApt = parameter.neueWohnung;

  if (dispatchDaten.wohnungsStatus) {
    if (altesApt && dispatchDaten.wohnungsStatus[altesApt]) {
      dispatchDaten.wohnungsStatus[altesApt].lebenszyklus = "UMGEZOGEN";
    }
    if (neuesApt) {
      dispatchDaten.wohnungsStatus[neuesApt] = {
        apartmentId: neuesApt,
        lebenszyklus: "GEBUCHT",
        abrechnungsZyklus: parameter.abrechnungsZyklus || "14 Tage"
      };
    }
  }
  Logger.log(`Umzug verarbeitet: '${altesApt}' -> '${neuesApt}'.`);
  return true;
}

function handleSupportTicket(aktion) {
  const parameter = aktion.parameter || {};
  Logger.log(`Support/Mangel erfasst (${aktion.apartmentId || "Allgemein"}): ${parameter.konkreteFrage || parameter.frage || parameter.mangelDetail || ""}`);
  return true;
}

function handleDokumentenAnfrage(aktion) {
  Logger.log(`Dokumentenanfrage registriert für Apartment: ${aktion.apartmentId}`);
  return true;
}

function handleExpansionsAnfrage(aktion) {
  const parameter = aktion.parameter || {};
  Logger.log(`Expansionsanfrage: ${parameter.personen || 0} Personen für Ort '${parameter.zielOrt || parameter.ort || ""}'.`);
  return true;
}

function handlePreisverhandlungLog(aktion) {
  const parameter = aktion.parameter || {};
  Logger.log(`Preisverhandlung für '${aktion.apartmentId}': Gegenangebot ${parameter.gegenangebotPreis}`);
  return true;
}

// -----------------------------------------------------------------
// 5. ZEILEN-BUILDER & FORMATIERUNG
// -----------------------------------------------------------------
function baueRechnungsZeilenArray(daten) {
  const zeile = new Array(38).fill("");
  zeile[0] = daten.telefon;
  zeile[1] = daten.firma;
  zeile[2] = daten.name;
  zeile[3] = daten.strasse;
  zeile[4] = daten.plz;
  zeile[5] = daten.stadt;
  zeile[6] = daten.land;
  zeile[7] = daten.ustId;
  zeile[8] = daten.email;
  zeile[9] = "Done";
  zeile[10] = daten.anreiseDE;
  zeile[11] = daten.abreiseDE;
  zeile[12] = daten.zyklus;
  zeile[13] = daten.apartmentId;
  zeile[14] = daten.preisProNacht;
  zeile[20] = "Neu angelegt";
  zeile[30] = "Offen";
  return zeile;
}

function baueVerlaengerungsZeilenArray(daten) {
  const zeile = new Array(39).fill("");
  zeile[0] = daten.telefon;
  zeile[1] = daten.firma;
  zeile[2] = daten.name;
  zeile[3] = daten.strasse;
  zeile[4] = daten.plz;
  zeile[5] = daten.stadt;
  zeile[6] = daten.land;
  zeile[7] = daten.ustId;
  zeile[8] = daten.email;
  zeile[9] = "Done";
  zeile[10] = daten.anreiseDE;
  zeile[11] = daten.abreiseDE;
  zeile[12] = daten.zyklus;
  zeile[13] = daten.apartmentId;
  zeile[14] = daten.preisProNacht;
  zeile[20] = "Verlängerung angelegt";
  zeile[31] = "Offen";
  return zeile;
}

function berechneAbreiseDatum(anreiseIso, abreiseIso, zusatzTage) {
  if (abreiseIso) return abreiseIso;
  if (!anreiseIso || !zusatzTage) return "";

  const basisDatum = new Date(anreiseIso);
  if (isNaN(basisDatum.getTime())) return "";

  basisDatum.setDate(basisDatum.getDate() + Number(zusatzTage));
  return basisDatum.toISOString().split("T")[0];
}

function formatiereDatumDE(isoDatumStr) {
  if (!isoDatumStr) return "";
  const datum = new Date(isoDatumStr);
  if (isNaN(datum.getTime())) return String(isoDatumStr);

  const tag = String(datum.getDate()).padStart(2, '0');
  const monat = String(datum.getMonth() + 1).padStart(2, '0');
  const jahr = datum.getFullYear();
  return `${tag}.${monat}.${jahr}`;
}

function ermittleBesteEmail(dispatchDaten, leadObjekt) {
  if (Array.isArray(dispatchDaten.kontakte)) {
    const buchhaltung = dispatchDaten.kontakte.find(k => k.rolle === "BUCHHALTUNG" && k.email);
    if (buchhaltung) return buchhaltung.email;
    const disponent = dispatchDaten.kontakte.find(k => k.rolle === "DISPONENT" && k.email);
    if (disponent) return disponent.email;
  }
  return leadObjekt.email || "";
}

function aktualisiereWohnungsStatusNachBuchung(dispatchDaten, aptId, zyklus, anreise, abreise, preis) {
  if (!aptId || !dispatchDaten.wohnungsStatus) return;
  if (!dispatchDaten.wohnungsStatus[aptId]) {
    dispatchDaten.wohnungsStatus[aptId] = { apartmentId: aptId };
  }
  dispatchDaten.wohnungsStatus[aptId].lebenszyklus = "GEBUCHT";
  dispatchDaten.wohnungsStatus[aptId].abrechnungsZyklus = zyklus;
  dispatchDaten.wohnungsStatus[aptId].anreise = anreise;
  dispatchDaten.wohnungsStatus[aptId].abreise = abreise;
  dispatchDaten.wohnungsStatus[aptId].preisProNacht = preis;
}

function spiegleStateInsArchiv(archivBlatt, nachrichtId, stateJsonText) {
  if (!archivBlatt || !nachrichtId) return;
  const letzteZeile = archivBlatt.getLastRow();
  if (letzteZeile <= 1) return;

  const idSpalte = archivBlatt.getRange(2, 6, letzteZeile - 1, 1).getValues();
  for (let r = 0; r < idSpalte.length; r++) {
    if (String(idSpalte[r][0] || "").trim() === nachrichtId) {
      archivBlatt.getRange(r + 2, 8).setValue(stateJsonText);
      return;
    }
  }
}

function parseJsonSicher(text, standardWert) {
  if (!text) return standardWert;
  try {
    return JSON.parse(text);
  } catch (fehler) {
    return standardWert;
  }
}