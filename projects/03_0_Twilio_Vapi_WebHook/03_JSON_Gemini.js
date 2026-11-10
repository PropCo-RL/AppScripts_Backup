// =================================================================
// KONFIGURATION & MODELL
// =================================================================
const GEMINI_MODELL = 'gemini-3.8-flash';
const MAX_LAUFZEIT_MS = 5 * 60 * 1000; // 5 Minuten Timeout-Puffer (Apps Script Limit ist 6 Min)

// =================================================================
// HAUPTFUNKTION (TRIGGER / ORCHESTRATOR)
// =================================================================
/**
 * Liest offene Sprachanrufe ein, analysiert die Audio-Datei mit Gemini 3.8 Flash
 * und schreibt das MMZ-kompatible JSON in Spalte J.
 */
function verarbeiteAudioAufnahmenMitGemini() {
  const startZeit = Date.now();
  const apiKey = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  if (!apiKey) {
    Logger.log('ABBRUCH: "Gemini-Key" fehlt in den ScriptProperties.');
    return;
  }

  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const letzteZeile = tabellenBlatt.getLastRow();
  if (letzteZeile <= 1) return;

  // Spalten A bis J komplett einlesen für schnelles Screening
  const datenBereich = tabellenBlatt.getRange(2, 1, letzteZeile - 1, 10);
  const zeilenWerte = datenBereich.getValues();
  let anzahlAktualisiert = 0;

  for (let i = 0; i < zeilenWerte.length; i++) {
    // Schutz vor Apps Script 6-Minuten-Abbruch
    if (Date.now() - startZeit > MAX_LAUFZEIT_MS) {
      Logger.log(`Laufzeit-Limit fast erreicht. Beende vorzeitig bei Zeile ${i + 2}. Nächste Calls laufen im nächsten Trigger.`);
      break;
    }

    const zeilenDaten = zeilenWerte[i];
    const zeilenNummer = i + 2;
    const kanal = String(zeilenDaten[2] || '').trim(); // Spalte C
    const audioUrl = String(zeilenDaten[5] || '').trim(); // Spalte F
    const jsonSpalteJ = String(zeilenDaten[9] || '').trim(); // Spalte J

    // Nur Sprachanrufe mit Audio-Link ohne bestehendes JSON verarbeiten
    if (kanal !== 'Sprachanruf' || !audioUrl || jsonSpalteJ !== '') {
      continue;
    }

    Logger.log(`[Zeile ${zeilenNummer}] Analysiere Audiodatei...`);
    const jsonErgebnis = verarbeiteEinzelnenAudioCall(zeilenDaten, audioUrl, apiKey);

    if (jsonErgebnis) {
      // Direktes Speichern pro Zeile (Schutz vor Datenverlust bei Timeout)
      tabellenBlatt.getRange(zeilenNummer, 10).setValue(jsonErgebnis);
      anzahlAktualisiert++;
    }
  }

  Logger.log(`Erfolg: ${anzahlAktualisiert} Zeile(n) verarbeitet.`);
}

// =================================================================
// MODUL: EINZELVERARBEITUNG & AUDIO-INLINE
// =================================================================
function verarbeiteEinzelnenAudioCall(zeilenDaten, audioUrl, apiKey) {
  const audioBlob = ladeAudioDatei(audioUrl);
  if (!audioBlob) return null;

  const datumUhrzeit = String(zeilenDaten[0] || ''); // Spalte A
  const telefonnummer = String(zeilenDaten[1] || ''); // Spalte B
  const fehlerhaftesTranskript = String(zeilenDaten[4] || ''); // Spalte E
  const anrufDatumIso = extrahiereIsoDatum(datumUhrzeit);

  const promptText = `
Du bist ein extrem präziser Buchungs-Analyst für Monteurzimmer-Vermietung bei L8 Street.
Höre dir die beigefügte Aufnahme des Telefonats aufmerksam an.
Das bisherige Live-Transkript war ungenau.

REFERENZDATUM FÜR RELATIVE ZEITANGABEN: "${anrufDatumIso}" (Datum des Anrufs aus Spalte A).
Alle relativen Angaben ("morgen", "nächste Woche", "ab 1.") beziehen sich zwingend auf dieses Datum!

METADATEN DES ANRUFS:
- Anruf-Zeitpunkt: "${datumUhrzeit}"
- Anrufer-Nummer: "${telefonnummer}"
- Fehlerhaftes Live-Transkript: """${fehlerhaftesTranskript}"""

FALLBACK-REGELN & STANDARDWERTE:
1. ZITATE-PFLICHT (ANTI-HALLUZINATION):
   - zitate.personenFundstelle: Wörtliches Zitat der Aussage zur Personenzahl/Zimmer aus dem Audio (oder "nicht genannt").
   - zitate.zeitraumFundstelle: Wörtliches Zitat der Aussage zum Zeitraum/Datum aus dem Audio (oder "nicht genannt").

2. telefonnummer & zeitpunkt: Exakt die übergebenen Werte ("${telefonnummer}" und "${datumUhrzeit}") einsetzen.

3. personenData & zimmer:
   - Pers: Startanzahl Personen als Zahl.
   - max: Wenn keine Aufstockung genannt, setze max = Pers.
   - aufteilung: Textliche Zusammenfassung der Betten-/Zimmerwünsche (z. B. "2 Einzel, 3 Doppel", "nur Einzelbetten", sonst "Standard").
   - zimmer: Explizite Zimmeraufteilung laut Gast oder "".

4. ZIMMER-BERECHNUNG (spezialData.minZimmer & belegungsTyp):
   - Wenn der Gast konkrete Zimmer wünscht (z. B. "2 Einzel und 3 Doppel" -> 5 Zimmer, oder "wir wollen 4 Zimmer"):
     * belegungsTyp = "WUNSCH_ZIMMER"
     * zimmerAnzahl = Gesamtzahl gewünschter Zimmer als Zahl
     * minZimmer = zimmerAnzahl
   - Wenn der Gast nur "Einzelzimmer" / "jeder ein eigenes Zimmer" will:
     * belegungsTyp = "EINZELZIMMER"
     * zimmerAnzahl = Pers
     * minZimmer = Pers
   - STANDARD-FALL (Gast nennt nur Personenzahl ohne Zimmerwünsche):
     * belegungsTyp = "STANDARD"
     * zimmerAnzahl = null
     * minZimmer = Math.ceil(Pers / 2), mindestens 1.

5. FLAGS (ANTI-HALLUZINATION):
   - Setze alle Flags (alleinnutzung, waschmaschine, wlan, parkplatz, einzelbetten, eigenesBad) STANDARDMÄSSIG AUF false!
   - Setze ein Flag NUR DANN auf true, wenn der Gast es im Audio AUSDRÜCKLICH verlangt (z. B. "wir haben einen großen Sprinter" -> parkplatz: true; "brauchen aber getrennte Betten" -> einzelbetten: true).

6. WEITERE FELDER:
   - spezialData.unterkunftsTyp: Standardmäßig "Monteurzimmer".
   - spezialData.kontaktKanal: Fest auf "Telefon".
   - spezialData.unserZeichen: Immer "".
   - spezialData.direktEmail: Die im Audio genannte E-Mail-Adresse oder "".
   - spezialData.fragenUndSonstiges: Spontane Fragen, Schadensmeldungen oder "".

7. ZEITRAUM:
   - start & ende: Format YYYY-MM-DD. Falls noch unklar: null.
   - verlaengerung: true, wenn Projekt/Verlängerung erwähnt, sonst false.
   - muster: "durchgehend" (oder "mo_do" falls nur Mo-Do gewünscht).
   - info: Details oder "".

8. gastIstVage:
   - true, wenn gastConfidenceScore < 85 (z. B. "ca.", "3 bis 4 Leute", "wissen noch nicht genau"), sonst false.
   - vageGruende: Kurze Begründung oder "".

9. istStaffelAnfrage:
   - true nur bei wechselnden Belegungsstufen, sonst false.

10. abweichungen_zu_spalte_e:
    - Dokumentiere präzise Hörfehler des Live-Transkripts im Vergleich zur tatsächlichen Tonspur.
`;

  const teile = [
    {
      inlineData: {
        mimeType: 'audio/wav',
        data: Utilities.base64Encode(audioBlob.getBytes())
      }
    },
    { text: promptText }
  ];

  const geminiAntwort = executeGeminiAudioRequest(teile, apiKey);
  return geminiAntwort ? JSON.stringify(geminiAntwort) : null;
}

// =================================================================
// MODUL: GEMINI API REQUEST MIT STRUCTURED OUTPUT
// =================================================================
function executeGeminiAudioRequest(teile, apiKey) {
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODELL}:generateContent`;

  const schema = {
    type: "OBJECT",
    properties: {
      telefonnummer: { type: "STRING" },
      zeitpunkt: { type: "STRING" },
      kiConfidenceScore: { type: "INTEGER", description: "0-100 Akustische Verständlichkeit" },
      zitate: {
        type: "OBJECT",
        properties: {
          personenFundstelle: { type: "STRING" },
          zeitraumFundstelle: { type: "STRING" }
        },
        required: ["personenFundstelle", "zeitraumFundstelle"]
      },
      gastConfidenceScore: { type: "INTEGER", description: "0-100 Verbindlichkeit des Kunden" },
      gastIstVage: { type: "BOOLEAN" },
      vageGruende: { type: "STRING" },
      suchortText: { type: "STRING", description: "Ort oder PLZ aus dem Audio" },
      personenData: {
        type: "OBJECT",
        properties: {
          Pers: { type: "INTEGER" },
          max: { type: "INTEGER" },
          aufteilung: { type: "STRING" },
          zimmer: { type: "STRING" }
        },
        required: ["Pers", "max", "aufteilung", "zimmer"]
      },
      zeitraumData: {
        type: "OBJECT",
        properties: {
          start: { type: "STRING", nullable: true, description: "YYYY-MM-DD oder null" },
          ende: { type: "STRING", nullable: true, description: "YYYY-MM-DD oder null" },
          maxTage: { type: "INTEGER", nullable: true },
          abrechnung: { type: "INTEGER", nullable: true },
          verlaengerung: { type: "BOOLEAN" },
          muster: { type: "STRING" },
          info: { type: "STRING" }
        },
        required: ["verlaengerung", "muster", "info"]
      },
      spezialData: {
        type: "OBJECT",
        properties: {
          unterkunftsTyp: { type: "STRING" },
          zimmerAnzahl: { type: "INTEGER", nullable: true },
          minZimmer: { type: "INTEGER" },
          belegungsTyp: { type: "STRING" },
          flags: {
            type: "OBJECT",
            properties: {
              alleinnutzung: { type: "BOOLEAN" },
              waschmaschine: { type: "BOOLEAN" },
              wlan: { type: "BOOLEAN" },
              parkplatz: { type: "BOOLEAN" },
              einzelbetten: { type: "BOOLEAN" },
              eigenesBad: { type: "BOOLEAN" }
            },
            required: ["alleinnutzung", "waschmaschine", "wlan", "parkplatz", "einzelbetten", "eigenesBad"]
          },
          direktEmail: { type: "STRING" },
          kontaktKanal: { type: "STRING" },
          unserZeichen: { type: "STRING" },
          budgetMonatGesamt: { type: "NUMBER", nullable: true },
          fragenUndSonstiges: { type: "STRING" }
        },
        required: ["unterkunftsTyp", "minZimmer", "belegungsTyp", "flags", "direktEmail", "kontaktKanal", "unserZeichen", "fragenUndSonstiges"]
      },
      istStaffelAnfrage: { type: "BOOLEAN" },
      abweichungen_zu_spalte_e: {
        type: "OBJECT",
        properties: {
          live_transkript_fehler: { type: "STRING" },
          korrigierte_details: { type: "STRING" }
        },
        required: ["live_transkript_fehler", "korrigierte_details"]
      }
    },
    required: [
      "telefonnummer",
      "zeitpunkt",
      "kiConfidenceScore",
      "zitate",
      "gastConfidenceScore",
      "gastIstVage",
      "vageGruende",
      "suchortText",
      "personenData",
      "zeitraumData",
      "spezialData",
      "istStaffelAnfrage",
      "abweichungen_zu_spalte_e"
    ]
  };

  const payload = {
    contents: [{ parts: teile }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: schema,
      temperature: 0.1
    }
  };

  const optionen = {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const antwort = UrlFetchApp.fetch(apiUrl, optionen);
    if (antwort.getResponseCode() !== 200) {
      Logger.log(`Gemini Fehler ${antwort.getResponseCode()}: ${antwort.getContentText()}`);
      return null;
    }

    const json = JSON.parse(antwort.getContentText());
    if (json.candidates && json.candidates[0]?.content?.parts?.[0]?.text) {
      return JSON.parse(json.candidates[0].content.parts[0].text);
    }
  } catch (fehler) {
    Logger.log("Netzwerk-/Parsing-Fehler bei Gemini: " + fehler.toString());
  }

  return null;
}

// =================================================================
// MODUL: HILFSFUNKTIONEN & DATEI-DOWNLOAD
// =================================================================
function extrahiereIsoDatum(datumString) {
  if (!datumString) return Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd");
  
  // Erkennt z. B. 09.11.2026 oder 09.11.2026 15:40:00
  const match = String(datumString).match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (match) {
    const tag = match[1].padStart(2, '0');
    const monat = match[2].padStart(2, '0');
    const jahr = match[3];
    return `${jahr}-${monat}-${tag}`;
  }
  
  const d = new Date(datumString);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, "GMT+1", "yyyy-MM-dd");
  }
  return Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd");
}

function ladeAudioDatei(audioUrl) {
  try {
    const vapiKey = PropertiesService.getScriptProperties().getProperty('Vapi-Key');
    if (!vapiKey) {
      Logger.log('FEHLER: Kein "Vapi-Key" in ScriptProperties hinterlegt.');
      return null;
    }

    const treffer = audioUrl.match(/([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i);
    const callId = treffer ? treffer[1] : null;
    if (!callId) {
      Logger.log('Call-ID konnte nicht ermittelt werden: ' + audioUrl);
      return null;
    }

    // 1. Call-Objekt von Vapi abfragen
    const vapiRes = UrlFetchApp.fetch(`https://api.vapi.ai/call/${callId}`, {
      headers: { 'Authorization': `Bearer ${vapiKey}` },
      muteHttpExceptions: true
    });

    if (vapiRes.getResponseCode() !== 200) {
      Logger.log(`Vapi Call-Abruf fehlgeschlagen (${vapiRes.getResponseCode()})`);
      return null;
    }

    const daten = JSON.parse(vapiRes.getContentText());

    // 2. Die signierte URL aus artifact.presignedMonoUrl auslesen
    const signierteUrl = daten.artifact?.presignedMonoUrl ||
                         daten.artifact?.presignedStereoUrl;

    if (!signierteUrl) {
      Logger.log(`Keine presignedMonoUrl für Call ${callId} vorhanden.`);
      return null;
    }

    // 3. Audio über den signierten Link direkt herunterladen
    const audioDownload = UrlFetchApp.fetch(signierteUrl, { muteHttpExceptions: true });
    if (audioDownload.getResponseCode() !== 200) {
      Logger.log(`Audio-Download fehlgeschlagen (${audioDownload.getResponseCode()})`);
      return null;
    }

    const audioBlob = audioDownload.getBlob().setContentType('audio/wav');
    Logger.log(`Audio erfolgreich geladen (${audioBlob.getBytes().length} Bytes) für Call ${callId}`);
    return audioBlob;
  } catch (fehler) {
    Logger.log('Ausnahmefehler beim Audio-Download: ' + fehler.toString());
    return null;
  }
}