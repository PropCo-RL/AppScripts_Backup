// =================================================================
// ORCHESTRATOR / TRIGGER: K.O.- & ROUTING-VORFILTER
// =================================================================
function runLeadPreFilterAudio() {
  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const letzteZeile = tabellenBlatt.getLastRow();
  if (letzteZeile <= 1) return;

  stelleHeaderSicher(tabellenBlatt);

  const anzahlZeilen = letzteZeile - 1;
  const datenBereich = tabellenBlatt.getRange(2, 11, anzahlZeilen, 2);
  const zeilenWerte = datenBereich.getValues();
  let aenderungenVorhanden = false;

  for (let i = 0; i < zeilenWerte.length; i++) {
    const jsonK = String(zeilenWerte[i][0] || "").trim();
    const statusL = String(zeilenWerte[i][1] || "").trim();

    if (!jsonK || jsonK.startsWith("Fehler") || statusL !== "") {
      continue;
    }

    const geparstesJson = parseValidiereJson(jsonK);
    if (!geparstesJson) {
      zeilenWerte[i][1] = JSON.stringify({ status: "FEHLER_JSON", grund: "Spalte K unlesbar" });
      aenderungenVorhanden = true;
      continue;
    }

    zeilenWerte[i][1] = JSON.stringify(ermittleRoutingFuerLead(geparstesJson));
    aenderungenVorhanden = true;
  }

  if (aenderungenVorhanden) {
    datenBereich.setValues(zeilenWerte);
    Logger.log("Routing-Vorfilter für Audio-Calls erfolgreich abgeschlossen.");
  }
}

// =================================================================
// MODUL: BUSINESS-LOGIK & KENNZAHLEN
// =================================================================
function ermittleRoutingFuerLead(validiertesJson) {
  const personen = Number(validiertesJson.personen?.Pers) || 0;
  const dauer = validiertesJson.dauer || {};
  const tage = Number(dauer.tage) || 0;
  const maxTage = Number(dauer.maxTage) || tage;
  const hatVerlaengerung = Boolean(dauer.verlaengerung);

  const zimmerBedarf = berechneEffektiveZimmer(validiertesJson.personen, validiertesJson.spezial, personen);
  const angebotsZeitraum = passeAngebotsZeitraumAn(dauer.start, dauer.ende, tage, hatVerlaengerung);
  const istKombiBedarf = personen > 6 || zimmerBedarf.anzahl >= 4;

  const routingEntscheidung = bestimmeRoutingKaskade(
    personen,
    angebotsZeitraum.tage,
    maxTage,
    istKombiBedarf,
    zimmerBedarf
  );

  return {
    status: routingEntscheidung.status,
    details: routingEntscheidung.details,
    kennzahlen: {
      pers: personen,
      zimmer: zimmerBedarf.anzahl,
      einzelzimmerPflicht: zimmerBedarf.istEinzelzimmer,
      tage: angebotsZeitraum.tage,
      maxTage: maxTage,
      hatVerlaengerung: hatVerlaengerung,
      start: angebotsZeitraum.start,
      ende: angebotsZeitraum.ende
    },
    geprueftAm: new Date().toISOString()
  };
}

// =================================================================
// MODUL: ROUTING-KASKADE (STRIKTES EARLY RETURN)
// =================================================================
function bestimmeRoutingKaskade(personen, angebotsTage, maxTage, istKombiBedarf, zimmerBedarf) {
  if (personen <= 1) {
    return { status: "REJECT_1PERS", details: "Nur 1 Person angefragt (K.O.-Kriterium)" };
  }

  if (angebotsTage < 6) {
    return { status: "REVIEW_SHORT", details: "Kurzzeitbuchung unter 6 Tagen ohne Verlängerung" };
  }

  const istLangzeit = maxTage >= 30;

  if (istKombiBedarf) {
    const grundText = personen > 6
      ? `> 6 Pers. (${personen}P)`
      : `ab 4 Zimmer (${zimmerBedarf.anzahl} Zimmer${zimmerBedarf.istEinzelzimmer ? ' / Einzelzimmer' : ''})`;

    return {
      status: istLangzeit ? "BIG_LONG" : "BIG_SHORT",
      details: istLangzeit ? `Großgruppe Langzeit (>= 30 Tage, ${grundText})` : `Großgruppe Kurzzeit (6-29 Tage, ${grundText})`
    };
  }

  const zimmerZusatz = zimmerBedarf.istEinzelzimmer
    ? ` [${zimmerBedarf.anzahl}x EZ]`
    : (zimmerBedarf.anzahl > 0 ? ` [${zimmerBedarf.anzahl} Zi]` : "");

  return {
    status: istLangzeit ? "STD_LONG" : "STD_SHORT",
    details: istLangzeit
      ? `Standard Langzeit (>= 30 Tage, 2-6 Pers.${zimmerZusatz})`
      : `Standard Kurzzeit (6-29 Tage, 2-6 Pers.${zimmerZusatz})`
  };
}

// =================================================================
// MODUL: HILFSFUNKTIONEN & BERECHNUNGEN
// =================================================================
function berechneEffektiveZimmer(personenDaten, spezialDaten, personenAnzahl) {
  const pData = personenDaten || {};
  const sData = spezialDaten || {};
  const zimmerAnzahl = Number(sData.zimmerAnzahl) || 0;
  const minZimmer = Number(sData.minZimmer) || Math.ceil(personenAnzahl / 2);
  const belegungsTyp = String(sData.belegungsTyp || "").toUpperCase();
  const aufteilung = String(pData.aufteilung || "").toLowerCase();
  const zimmerText = String(pData.zimmer || "").toLowerCase();

  const istEinzelzimmer = belegungsTyp === "EINZELZIMMER" ||
                          aufteilung.includes("einzelzimmer") ||
                          zimmerText.includes("einzelzimmer");

  let anzahl = minZimmer;
  if (istEinzelzimmer) {
    anzahl = Math.max(zimmerAnzahl, personenAnzahl, minZimmer);
  } else if (zimmerAnzahl > 0) {
    anzahl = Math.max(zimmerAnzahl, minZimmer);
  }

  return { anzahl: anzahl, istEinzelzimmer: istEinzelzimmer };
}

function passeAngebotsZeitraumAn(startStr, endeStr, tage, hatVerlaengerung) {
  let angebotsTage = tage;
  if (hatVerlaengerung && angebotsTage < 6) {
    angebotsTage = 6;
  }

  let angebotsStart = String(startStr || "").trim();
  let angebotsEnde = String(endeStr || "").trim();

  if (angebotsStart && (!angebotsEnde || angebotsTage !== tage)) {
    const datumStart = new Date(angebotsStart + "T00:00:00Z");
    if (!isNaN(datumStart.getTime())) {
      const datumEnde = new Date(datumStart.getTime() + angebotsTage * 86400000);
      angebotsEnde = datumEnde.toISOString().split("T")[0];
    }
  }

  return { tage: angebotsTage, start: angebotsStart, ende: angebotsEnde };
}

function parseValidiereJson(text) {
  try {
    return JSON.parse(text);
  } catch (fehler) {
    return null;
  }
}

function stelleHeaderSicher(blatt) {
  const headerZelle = blatt.getRange(1, 12);
  if (!headerZelle.getValue()) {
    headerZelle.setValue("Routing Status (JSON)");
    headerZelle.setFontWeight("bold");
    blatt.getRange("L:L").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}