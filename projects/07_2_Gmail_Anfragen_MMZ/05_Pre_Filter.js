// =================================================================
// STUFE 1: K.O.- & ROUTING-VORFILTER FÜR TAB "MMZ"
// Liest Spalte D (Gesamt-JSON) und schreibt Routing-JSON in Spalte E
// =================================================================
function runLeadPreFilterMMZ() {
  const SHEET_NAME = 'MMZ';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`FEHLER: Tabellenblatt "${SHEET_NAME}" nicht gefunden.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Header für Spalte E sicherstellen
  ensurePreFilterHeaderMMZ(sheet);

  // Spalten B bis E einlesen (4 Spalten: B=Index 0, C=1, D=2, E=3)
  const numRows = lastRow - 1;
  const dataRange = sheet.getRange(2, 2, numRows, 4);
  const rows = dataRange.getValues();
  let updatesMade = false;

  for (let i = 0; i < rows.length; i++) {
    const cellB = String(rows[i][0] || "").trim(); // Spalte B: Unique
    const cellD = String(rows[i][2] || "").trim(); // Spalte D: Validiertes Gesamt-JSON
    const cellE = String(rows[i][3] || "").trim(); // Spalte E: Routing-JSON Ausgabe

    // Überspringen wenn: kein B da, Duplikat "-", kein D vorhanden, oder Spalte E bereits befüllt ist
    if (!cellB || cellB === "-" || !cellD || cellD.startsWith("Fehler") || cellE !== "") {
      continue;
    }

    let valData = {};
    try {
      valData = JSON.parse(cellD);
    } catch (e) {
      rows[i][3] = JSON.stringify({ status: "FEHLER_JSON", grund: "Spalte D unlesbar" });
      updatesMade = true;
      continue;
    }

    // Werte direkt aus dem Spalte-D-JSON beziehen
    const persData = valData.personen || {};
    const dauerData = valData.dauer || {};
    const spezialData = valData.spezial || {};

    const pers = Number(persData.Pers) || 0;
    const tage = Number(dauerData.tage) || 0;
    const maxTage = Number(dauerData.maxTage) || tage;
    const hatVerlaengerung = Boolean(dauerData.verlaengerung);

    // Zimmerbedarf & Einzelzimmer-Prüfung
    const zimmerAnzahl = Number(spezialData.zimmerAnzahl) || 0;
    const aufteilungStr = String(persData.aufteilung || "").toLowerCase();
    const zimmerStr = String(persData.zimmer || "").toLowerCase();
    const istEinzelzimmerWunsch = aufteilungStr.includes("einzelzimmer") || zimmerStr.includes("einzelzimmer");

    // Standardannahme ohne Sonderwunsch: ca. 2 Personen pro Zimmer
    const standardZimmer = Math.ceil(pers / 2);

    // Effektiver Mindest-Zimmerbedarf
    let effektiveZimmer = standardZimmer;
    if (istEinzelzimmerWunsch) {
      effektiveZimmer = Math.max(zimmerAnzahl, pers);
    } else if (zimmerAnzahl > 0) {
      effektiveZimmer = Math.max(zimmerAnzahl, standardZimmer);
    }

    // Ab wann eine Großgruppen-Kombi von vornherein zwingend ist:
    const brauchtKombi = pers > 6 || effektiveZimmer >= 4;

    // --- ANPASSUNG: Mindestaufenthalt bei Verlängerungsoption auf 6 Tage anheben & Enddatum berechnen ---
    let angebotsTage = tage;
    if (hatVerlaengerung && angebotsTage < 6) {
      angebotsTage = 6;
    }

    let angebotsStart = String(dauerData.start || "").trim();
    let angebotsEnde = String(dauerData.ende || "").trim();

    // Start vorhanden, aber Ende fehlt oder Mindestdauer wurde angehoben
    if (angebotsStart && (!angebotsEnde || angebotsTage !== tage)) {
      const dStart = new Date(angebotsStart + "T00:00:00Z");
      if (!isNaN(dStart.getTime())) {
        const dEnd = new Date(dStart.getTime() + angebotsTage * 86400000);
        angebotsEnde = dEnd.toISOString().split("T")[0];
      }
    }

    let routingStatus = "";
    let routingDetails = "";

    // --- DIE ROUTING-KASKADE ---
    // 1. Hartes K.O.: 1 Person (oder 0) fliegt immer raus
    if (pers <= 1) {
      routingStatus = "REJECT_1PERS";
      routingDetails = "Nur 1 Person angefragt (K.O.-Kriterium)";
    }
    // 2. Unter 6 Tage ohne Verlängerung, aber mehrere Personen -> Manuelle Prüfung
    else if (angebotsTage < 6) {
      routingStatus = "REVIEW_SHORT";
      routingDetails = "Kurzzeitbuchung unter 6 Tagen ohne Verlängerung";
    }
    // 3. Standard- & Großanfragen (angebotsTage >= 6 und pers >= 2)
    else {
      const istLangzeit = maxTage >= 30;
      if (brauchtKombi) {
        routingStatus = istLangzeit ? "BIG_LONG" : "BIG_SHORT";
        const grundText = pers > 6
          ? `> 6 Pers. (${pers}P)`
          : `ab 4 Zimmer (${effektiveZimmer} Zimmer${istEinzelzimmerWunsch ? ' / Einzelzimmer' : ''})`;
        routingDetails = istLangzeit
          ? `Großgruppe Langzeit (>= 30 Tage, ${grundText})`
          : `Großgruppe Kurzzeit (6-29 Tage, ${grundText})`;
      } else {
        routingStatus = istLangzeit ? "STD_LONG" : "STD_SHORT";
        const zimmerInfo = istEinzelzimmerWunsch ? ` [${effektiveZimmer}x EZ]` : (effektiveZimmer > 0 ? ` [${effektiveZimmer} Zi]` : "");
        routingDetails = istLangzeit
          ? `Standard Langzeit (>= 30 Tage, 2-6 Pers.${zimmerInfo})`
          : `Standard Kurzzeit (6-29 Tage, 2-6 Pers.${zimmerInfo})`;
      }
    }

    // JSON-Ergebnis für Spalte E zusammenstellen
    const resultObj = {
      status: routingStatus,
      details: routingDetails,
      kennzahlen: {
        pers: pers,
        zimmer: effektiveZimmer,
        einzelzimmerPflicht: istEinzelzimmerWunsch,
        tage: angebotsTage,
        maxTage: maxTage,
        hatVerlaengerung: hatVerlaengerung,
        start: angebotsStart,
        ende: angebotsEnde
      },
      geprueftAm: new Date().toISOString()
    };

    rows[i][3] = JSON.stringify(resultObj);
    updatesMade = true;
  }

  // Zurückschreiben und Zeilenhöhe sichern
  if (updatesMade) {
    dataRange.setValues(rows);
    for (let r = 2; r <= lastRow; r++) {
      sheet.setRowHeight(r, 21);
    }
    Logger.log("Routing-Vorfilter für MMZ erfolgreich abgeschlossen.");
  }
}

function ensurePreFilterHeaderMMZ(sheet) {
  const cell = sheet.getRange(1, 5); // Spalte E
  if (!cell.getValue()) {
    cell.setValue("Routing Status (JSON)");
    cell.setFontWeight("bold");
    sheet.getRange("E:E").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}