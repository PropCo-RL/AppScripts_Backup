// =================================================================
// 12_INBOUND DISPATCHER (MMZ & THREAD ARCHIVE) - MODULAR ARCHITECTURE
// - Verarbeitet das 'Thread Archive' strikt per Delta über Spalte H
// - Schreibt bei 'out'-Mails sofort "-" (bereits erledigt)
// - Führt Gemini-Analyse NUR bei neuen 'in'-Mails aus (wo Spalte H leer ist)
// - Spiegelt den aktuellen State atomar nach Spalte H & MMZ (Spalte Q)
// =================================================================

const DISPATCHER_KONFIGURATION = {
  MODELL_NAME: 'gemini-3.8-flash',
  TABELLE_HAUPT: 'MMZ',
  TABELLE_ARCHIV: 'Thread Archive',
  TABELLE_STAMMDATEN: 'd',
  SPALTE_LEAD: 2,       // B: Lead JSON
  SPALTE_DRAFT: 10,     // J: Draft Status JSON
  SPALTE_STATE_MMZ: 17, // Q: INBOUND_DISPATCH im Blatt MMZ
  SPALTE_ARCHIV_STATE: 8 // H: State JSON im Thread Archive
};

// -----------------------------------------------------------------
// 1. ZENTRALER ORCHESTRATOR
// -----------------------------------------------------------------
function runGeminiInboundDispatcherMMZ() {
  const geminiApiKey = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  if (!geminiApiKey) {
    Logger.log("FEHLER: Script-Property 'Gemini-Key' fehlt.");
    return;
  }

  const arbeitsmappe = SpreadsheetApp.getActiveSpreadsheet();
  const hauptBlatt = arbeitsmappe.getSheetByName(DISPATCHER_KONFIGURATION.TABELLE_HAUPT);
  const archivBlatt = arbeitsmappe.getSheetByName(DISPATCHER_KONFIGURATION.TABELLE_ARCHIV);

  if (!hauptBlatt || !archivBlatt) {
    Logger.log("FEHLER: Blatt 'MMZ' oder 'Thread Archive' nicht gefunden.");
    return;
  }

  const anzahlArchivZeilen = archivBlatt.getLastRow() - 1;
  if (anzahlArchivZeilen <= 0) return;

  sicherstelleDispatcherHeader(hauptBlatt);
  const wohnungsKatalog = ladeWohnungsKatalogAusTabD(arbeitsmappe);
  
  // 1. Alle Archivdaten und Lead-Referenzen im Batch laden
  const archivBatch = ladeArchivBatchDaten(archivBlatt, anzahlArchivZeilen);
  const mmzLeadMap = ladeMmzLeadReferenzen(hauptBlatt);

  // 2. Nur unberührte Zeilen im Archiv verarbeiten
  const verarbeitungsErgebnis = verarbeiteArchivDelta(
    archivBatch, mmzLeadMap, wohnungsKatalog, geminiApiKey
  );

  // 3. Batch-Rückschreiben bei Änderungen
  if (verarbeitungsErgebnis.archivGeaendert) {
    schreibeArchivStateBatch(archivBlatt, verarbeitungsErgebnis.neueArchivStates, anzahlArchivZeilen);
  }

  if (verarbeitungsErgebnis.mmzUpdates.size > 0) {
    schreibeMmzStateBatch(hauptBlatt, verarbeitungsErgebnis.mmzUpdates);
  }
}

// -----------------------------------------------------------------
// 2. SHEET-IO & BATCHING
// -----------------------------------------------------------------
function ladeArchivBatchDaten(archivBlatt, anzahlZeilen) {
  // Spalten A bis H einlesen (1: Zeit, 2: InOut, 3: E-Mail, 4: Art, 5: Nachricht, 6: ID, 7: JSON, 8: State)
  const daten = archivBlatt.getRange(2, 1, anzahlZeilen, 8).getValues();
  return {
    zeilen: daten,
    anzahl: anzahlZeilen
  };
}

function ladeMmzLeadReferenzen(hauptBlatt) {
  const anzahlZeilen = hauptBlatt.getLastRow() - 1;
  const leadMap = new Map();
  if (anzahlZeilen <= 0) return leadMap;

  // Lese Spalte B (Lead JSON), J (Draft JSON) und Q (State JSON)
  const leads = hauptBlatt.getRange(2, DISPATCHER_KONFIGURATION.SPALTE_LEAD, anzahlZeilen, 1).getValues();
  const drafts = hauptBlatt.getRange(2, DISPATCHER_KONFIGURATION.SPALTE_DRAFT, anzahlZeilen, 1).getValues();
  const states = hauptBlatt.getRange(2, DISPATCHER_KONFIGURATION.SPALTE_STATE_MMZ, anzahlZeilen, 1).getValues();

  for (let i = 0; i < anzahlZeilen; i++) {
    const roherLead = String(leads[i][0] || "").trim();
    if (!roherLead || roherLead === "-") continue;

    const leadObj = parseJsonSicher(roherLead, {});
    const email = String(leadObj.email || "").toLowerCase().trim();
    if (!email) continue;

    leadMap.set(email, {
      rowNum: i + 2,
      leadText: roherLead,
      leadObj: leadObj,
      draftText: String(drafts[i][0] || "").trim(),
      currentStateText: String(states[i][0] || "").trim()
    });
  }
  return leadMap;
}

function schreibeArchivStateBatch(archivBlatt, neueStates, anzahlZeilen) {
  const zielBereich = archivBlatt.getRange(2, DISPATCHER_KONFIGURATION.SPALTE_ARCHIV_STATE, anzahlZeilen, 1);
  zielBereich.setValues(neueStates);
  archivBlatt.setRowHeights(2, anzahlZeilen, 21);
  SpreadsheetApp.flush();
  Logger.log("Archiv-States erfolgreich in Spalte H aktualisiert.");
}

function schreibeMmzStateBatch(hauptBlatt, mmzUpdatesMap) {
  for (const [rowNum, stateJsonText] of mmzUpdatesMap.entries()) {
    hauptBlatt.getRange(rowNum, DISPATCHER_KONFIGURATION.SPALTE_STATE_MMZ).setValue(stateJsonText);
    hauptBlatt.setRowHeight(rowNum, 21);
  }
  SpreadsheetApp.flush();
  Logger.log(`${mmzUpdatesMap.size} Status-Update(s) nach MMZ (Spalte Q) synchronisiert.`);
}

// -----------------------------------------------------------------
// 3. DELTA-VERARBEITUNG & LOGIK
// -----------------------------------------------------------------
function verarbeiteArchivDelta(archivBatch, mmzLeadMap, wohnungsKatalog, geminiApiKey) {
  const anzahl = archivBatch.anzahl;
  const neueArchivStates = [];
  let archivGeaendert = false;
  const mmzUpdates = new Map();

  for (let i = 0; i < anzahl; i++) {
    const zeile = archivBatch.zeilen[i];
    const bisherigerState = String(zeile[7] || "").trim(); // Spalte H

    // 1. GATEKEEPER: Bereits analysiert?
    if (bisherigerState !== "") {
      neueArchivStates.push([bisherigerState]);
      continue;
    }

    const inOut = String(zeile[1] || "").toLowerCase().trim(); // Spalte B
    const gastEmail = String(zeile[2] || "").toLowerCase().trim(); // Spalte C
    const nachrichtId = String(zeile[5] || "").trim(); // Spalte F

    // 2. Outbound-Mails sofort mit "-" abschließen
    if (inOut === "out") {
      neueArchivStates.push(["-"]);
      archivGeaendert = true;
      continue;
    }

    // 3. Nur eingehende Kunden-Nachrichten ("in") weiter prüfen
    if (inOut !== "in" || !nachrichtId || !gastEmail) {
      neueArchivStates.push(["-"]);
      archivGeaendert = true;
      continue;
    }

    Logger.log(`[Zeile ${i + 2}] Neue Inbound-Nachricht erkannt (${nachrichtId}) für: ${gastEmail}`);

    const mmzReferenz = mmzLeadMap.get(gastEmail);
    if (!mmzReferenz) {
      Logger.log(`Keine aktive Lead-Referenz in MMZ für ${gastEmail} gefunden.`);
      neueArchivStates.push(["-"]);
      archivGeaendert = true;
      continue;
    }

    // Thread-Verlauf bis zu dieser Nachricht aufbauen
    const threadHistorie = baueHistorieFuerGast(archivBatch.zeilen, i, gastEmail);
    const bisherigerZustand = parseJsonSicher(mmzReferenz.currentStateText, null);

    // Gemini-Aufruf
    const analysierterState = analysiereInboundMitGemini(
      mmzReferenz.leadText,
      mmzReferenz.draftText,
      threadHistorie,
      zeile,
      geminiApiKey,
      nachrichtId,
      wohnungsKatalog,
      bisherigerZustand
    );

    if (!analysierterState) {
      neueArchivStates.push(["-"]);
      archivGeaendert = true;
      continue;
    }

    const finalerState = fuehreZustaendeZusammen(bisherigerZustand, analysierterState, nachrichtId);
    const finalStateJson = JSON.stringify(finalerState);

    neueArchivStates.push([finalStateJson]);
    archivGeaendert = true;

    // Für MMZ-Rückschrieb vormerken und lokalen Cache aktualisieren
    mmzUpdates.set(mmzReferenz.rowNum, finalStateJson);
    mmzReferenz.currentStateText = finalStateJson;
  }

  return {
    archivGeaendert: archivGeaendert,
    neueArchivStates: neueArchivStates,
    mmzUpdates: mmzUpdates
  };
}

function baueHistorieFuerGast(alleZeilen, bisIndex, gastEmail) {
  const MAX_HISTORIE_NACHRICHTEN = 6; // Begrenzung auf die relevanten letzten Nachrichten
  const verlauf = [];
  for (let j = 0; j <= bisIndex; j++) {
    const z = alleZeilen[j];
    const email = String(z[2] || "").toLowerCase().trim();
    if (email === gastEmail) {
      verlauf.push({
        id: String(z[5] || ""),
        t: String(z[1] || "").toUpperCase(),
        date: String(z[0] || ""),
        text: String(z[4] || "")
      });
    }
  }
  // Schneidet automatisch nur die letzten N Nachrichten ab
  return verlauf.length > MAX_HISTORIE_NACHRICHTEN 
    ? verlauf.slice(-MAX_HISTORIE_NACHRICHTEN) 
    : verlauf;
}

// -----------------------------------------------------------------
// 4. GEMINI DISPATCHER CALL & PROMPT
// -----------------------------------------------------------------
function analysiereInboundMitGemini(leadText, draftText, verlauf, aktuelleZeile, apiKey, nachrichtId, katalog, alterZustand) {
  const draftObjekt = parseJsonSicher(draftText, {});
  const bekannteAptIds = sammleAngeboteneApartmentIds(draftObjekt, verlauf);
  const letztesAngebot = findeLetztesAngebotImVerlauf(verlauf);
  const relevanterKatalog = katalog.filter(k => bekannteAptIds.includes(k.id));
  const uebergebenerKatalog = relevanterKatalog.length > 0 ? relevanterKatalog : katalog.slice(0, 30);

  const empfangsZeit = String(aktuelleZeile[0] || new Date().toISOString());
  const promptText = baueDispatcherPrompt(
    leadText, draftObjekt, letztesAngebot, verlauf, empfangsZeit, nachrichtId, bekannteAptIds, uebergebenerKatalog, alterZustand
  );

  return sendeGeminiInteractionsAnfrage(promptText, apiKey);
}

function baueDispatcherPrompt(leadText, draftObj, letztesAngebot, verlauf, empfangsZeit, nachrichtId, aptIds, katalog, alterZustand) {
  const jetztIso = new Date().toISOString();
  const bekannterWohnungsStatus = (alterZustand && alterZustand.wohnungsStatus) ? alterZustand.wohnungsStatus : {};

  return `Du bist der Dispatcher für ein Monteurzimmer-Buchungssystem.
Analysiere die neueste Kundenantwort und zerlege sie in strukturierte Entitäten und atomare Aktionen.
Ein Gast kann mehrere Themen gleichzeitig ansprechen (z. B. Buchung + Expansion + Support).
REFERENZZEITPUNKT: ${empfangsZeit}
REGELN FÜR DATUM & ZEITRÄUME:
- Berechne Datumsangaben zwingend als ISO-Datum (YYYY-MM-DD) bezogen auf den Referenzzeitpunkt.
- Wenn eine Zyklusdauer angegeben ist (z. B. "14 Tage"), berechne "abreise" als (anreise + Tage) ODER setze "zusatzTage": 14. "abreise" darf bei einer Buchung NICHT null bleiben!
KATALOG AUS SYSTEM:
${JSON.stringify(katalog)}
ANGEBOTENE APARTMENT-IDS:
${JSON.stringify(aptIds)}
Verwende in "apartmentId" zwingend diese System-IDs (z. B. "bel0", "cs1").
VORHERIGER WOHNUNGSSTATUS:
${JSON.stringify(bekannterWohnungsStatus)}
INITIALER LEAD:
${leadText}
UNSER LETZTES ANGEBOT:
${letztesAngebot ? JSON.stringify(letztesAngebot) : "Kein vorheriges Angebot"}
VERLAUF MIT NEUESTER ANTWORT:
${JSON.stringify(verlauf)}

REGELN FÜR AKTIONEN (offeneAktionen):
Erstelle für jedes Thema ein eigenes Objekt. actionId MUSS eindeutig sein:
- Buchung: "ACT_BOOK_<apartmentId>_${nachrichtId}"
- Verlängerung: "ACT_EXT_<apartmentId>_${nachrichtId}"
- Expansion: "ACT_EXP_<zielOrt>_${nachrichtId}"
- Umzug: "ACT_MOVE_${nachrichtId}"
- Support: "ACT_SUP_<thema>_${nachrichtId}"
- Dokument: "ACT_DOC_<typ>_${nachrichtId}"

Buchungs-Status:
- "BEREIT_FUER_RE" wenn verbindlich gebucht ("wir nehmen das", Rechnungsadresse geschickt).
- "WARTET_AUF_FREIGABE" wenn Vorbehalt ("kläre mit Chef", "falls Parkplatz frei").

ANTWORTE AUSSCHLIESSLICH ALS REINES JSON (KEINE CODEBLÖCKE, KEIN MARKDOWN):
{
  "letzteMsgId": "${nachrichtId}",
  "letzteVerarbeiteteMsgId": "${nachrichtId}",
  "aktualisiertAm": "${jetztIso}",
  "erkannteThemen": ["BUCHUNG", "EXPANSION", "SUPPORT"],
  "organisation": {
    "firmenname": null,
    "ustId": null,
    "rechnungsanschriftDefault": { "empfaenger": null, "strasse": null, "plz": null, "stadt": null, "land": "Deutschland" }
  },
  "kontakte": [{ "rolle": "DISPONENT", "name": null, "email": null, "telefon": null }],
  "wohnungsStatus": {
    "bel0": { "apartmentId": "bel0", "lebenszyklus": "GEBUCHT", "abrechnungsZyklus": "14 Tage", "preisProNacht": null }
  },
  "offeneAktionen": [
    {
      "actionId": "ACT_BOOK_bel0_${nachrichtId}",
      "typ": "BUCHUNG",
      "status": "BEREIT_FUER_RE",
      "confidence": "DEFINITIV",
      "apartmentId": "bel0",
      "parameter": {
        "anreise": "YYYY-MM-DD",
        "abreise": "YYYY-MM-DD",
        "zusatzTage": 14,
        "abrechnungsZyklus": "14 Tage",
        "preisProNacht": null,
        "gegenangebotPreis": null,
        "personen": null,
        "zielOrt": null,
        "alteWohnung": null,
        "neueWohnung": null,
        "konkreteFrage": null,
        "mangelDetail": null
      }
    }
  ]
}`;
}

function sendeGeminiInteractionsAnfrage(promptText, apiKey) {
  const anfrageUrl = "https://generativelanguage.googleapis.com/v1beta/interactions";
  const nutzdaten = {
    model: DISPATCHER_KONFIGURATION.MODELL_NAME,
    input: promptText,
    store: false
  };

  const optionen = {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(nutzdaten),
    muteHttpExceptions: true
  };

  try {
    const antwort = UrlFetchApp.fetch(anfrageUrl, optionen);
    if (antwort.getResponseCode() !== 200) {
      Logger.log(`Gemini API Fehler (${antwort.getResponseCode()}): ${antwort.getContentText()}`);
      return null;
    }
    return extrahiereJsonAusAntwort(antwort.getContentText());
  } catch (fehler) {
    Logger.log("Netzwerkfehler bei Gemini API: " + fehler.toString());
    return null;
  }
}

function extrahiereJsonAusAntwort(antwortText) {
  try {
    const parsed = JSON.parse(antwortText);
    let rohText = "";
    if (parsed.output_text) {
      rohText = parsed.output_text.trim();
    } else if (parsed.steps && parsed.steps.length > 0) {
      const letzterSchritt = parsed.steps[parsed.steps.length - 1];
      if (letzterSchritt.content && letzterSchritt.content[0] && letzterSchritt.content[0].text) {
        rohText = letzterSchritt.content[0].text.trim();
      }
    }
    const treffer = rohText.match(/\{[\s\S]*\}/);
    return treffer ? JSON.parse(treffer[0]) : null;
  } catch (fehler) {
    Logger.log("Fehler beim Extrahieren des Antwort-JSON: " + fehler.toString());
    return null;
  }
}

// -----------------------------------------------------------------
// 5. HELPER & STATE DEEP-MERGE
// -----------------------------------------------------------------
function fuehreZustaendeZusammen(bisherigerZustand, neuerZustand, nachrichtId) {
  if (!bisherigerZustand) {
    neuerZustand.erledigteAktionen = [];
    neuerZustand.letzteMsgId = nachrichtId;
    neuerZustand.letzteVerarbeiteteMsgId = nachrichtId;
    neuerZustand.aktualisiertAm = new Date().toISOString();
    return neuerZustand;
  }

  const zusammengefuehrt = Object.assign({}, bisherigerZustand);
  const jetztIso = new Date().toISOString();
  zusammengefuehrt.letzteMsgId = nachrichtId;
  zusammengefuehrt.letzteVerarbeiteteMsgId = nachrichtId;
  zusammengefuehrt.aktualisiertAm = jetztIso;

  const erledigteAktionenSet = new Set(bisherigerZustand.erledigteAktionen || []);
  zusammengefuehrt.erledigteAktionen = Array.from(erledigteAktionenSet);

  zusammengefuehrt.offeneAktionen = kombiniereOffeneAktionen(
    bisherigerZustand.offeneAktionen || [],
    neuerZustand.offeneAktionen || [],
    erledigteAktionenSet
  );

  zusammengefuehrt.wohnungsStatus = Object.assign(
    {}, bisherigerZustand.wohnungsStatus || {}, neuerZustand.wohnungsStatus || {}
  );

  zusammengefuehrt.organisation = mergeOrganisationsDaten(
    bisherigerZustand.organisation, neuerZustand.organisation
  );

  if (Array.isArray(neuerZustand.kontakte) && neuerZustand.kontakte.length > 0) {
    zusammengefuehrt.kontakte = neuerZustand.kontakte;
  }

  return zusammengefuehrt;
}

function kombiniereOffeneAktionen(alteAktionen, neueAktionen, erledigteSet) {
  const aktionenMap = new Map();
  alteAktionen.forEach(aktion => aktionenMap.set(aktion.actionId, aktion));
  neueAktionen.filter(aktion => !erledigteSet.has(aktion.actionId))
    .forEach(aktion => aktionenMap.set(aktion.actionId, aktion));
  return Array.from(aktionenMap.values());
}

function mergeOrganisationsDaten(alteOrg, neueOrg) {
  const vorher = alteOrg || {};
  const aktuell = neueOrg || {};
  const vorherAdr = vorher.rechnungsanschriftDefault || {};
  const aktuellAdr = aktuell.rechnungsanschriftDefault || {};
  return {
    firmenname: aktuell.firmenname || vorher.firmenname || null,
    ustId: aktuell.ustId || vorher.ustId || null,
    rechnungsanschriftDefault: {
      empfaenger: aktuellAdr.empfaenger || vorherAdr.empfaenger || null,
      strasse: aktuellAdr.strasse || vorherAdr.strasse || null,
      plz: aktuellAdr.plz || vorherAdr.plz || null,
      stadt: aktuellAdr.stadt || vorherAdr.stadt || null,
      land: aktuellAdr.land || vorherAdr.land || "Deutschland"
    }
  };
}

function sammleAngeboteneApartmentIds(draftObjekt, verlauf) {
  const gefundeneIds = [];
  if (draftObjekt.apartmentId) gefundeneIds.push(draftObjekt.apartmentId);
  if (Array.isArray(draftObjekt.apartmentIds)) {
    draftObjekt.apartmentIds.forEach(id => gefundeneIds.push(id));
  }
  for (const msg of verlauf) {
    if (msg.t !== "OUT" || !msg.text) continue;
    const treffer = msg.text.matchAll(/a\.L8Street\.com\/a\/([a-zA-Z0-9_-]+)/g);
    for (const t of treffer) {
      if (!gefundeneIds.includes(t[1])) gefundeneIds.push(t[1]);
    }
  }
  return gefundeneIds;
}

function findeLetztesAngebotImVerlauf(verlauf) {
  for (let i = verlauf.length - 1; i >= 0; i--) {
    if (verlauf[i].t === "OUT") return verlauf[i];
  }
  return null;
}

function ladeWohnungsKatalogAusTabD(arbeitsmappe) {
  const blattD = arbeitsmappe.getSheetByName(DISPATCHER_KONFIGURATION.TABELLE_STAMMDATEN);
  if (!blattD) return [];
  const letzteZeile = blattD.getLastRow();
  if (letzteZeile <= 1) return [];

  const daten = blattD.getRange(2, 1, letzteZeile - 1, 10).getValues();
  const katalog = [];

  for (let r = 0; r < daten.length; r++) {
    const wohnungsName = String(daten[r][2] || "").trim();
    const wohnungsLink = String(daten[r][3] || "").trim();
    const strasse = String(daten[r][7] || "").trim();
    const hausNummer = String(daten[r][8] || "").trim();
    const ort = String(daten[r][9] || "").trim();
    let slugId = "";

    const treffer = wohnungsLink.match(/\/a\/([a-zA-Z0-9_-]+)/);
    if (treffer) slugId = treffer[1];

    const wohnungId = slugId || wohnungsName;
    if (wohnungId) {
      katalog.push({
        id: wohnungId,
        name: wohnungsName,
        adresse: `${strasse} ${hausNummer}, ${ort}`.trim()
      });
    }
  }
  return katalog;
}

function sicherstelleDispatcherHeader(hauptBlatt) {
  const kopfZelle = hauptBlatt.getRange(1, DISPATCHER_KONFIGURATION.SPALTE_STATE_MMZ);
  if (!kopfZelle.getValue() || kopfZelle.getValue() !== "INBOUND_DISPATCH (JSON)") {
    kopfZelle.setValue("INBOUND_DISPATCH (JSON)").setFontWeight("bold");
    hauptBlatt.getRange(1, DISPATCHER_KONFIGURATION.SPALTE_STATE_MMZ, hauptBlatt.getMaxRows(), 1)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
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