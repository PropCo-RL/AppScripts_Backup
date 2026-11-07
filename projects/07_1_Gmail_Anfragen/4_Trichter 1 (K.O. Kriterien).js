// =================================================================
// STUFE 1: K.O.- & ROUTING-VORFILTER (Schreibt Status in Spalte O)
// =================================================================
function runLeadPreFilterDMZ2() {
  const SHEET_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  let sheet;

  try {
    sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DMZ2');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabelle: " + e.toString());
    return;
  }

  if (!sheet) {
    Logger.log("FEHLER: Tabellenblatt 'DMZ2' nicht gefunden.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Header für Spalte O (Spalte 15) sicherstellen
  const headerCell = sheet.getRange(1, 15);
  if (!headerCell.getValue()) {
    headerCell.setValue("Routing Status");
    headerCell.setFontWeight("bold");
    sheet.getRange("O:O").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Spalten B bis O einlesen (14 Spalten: B=0 ... O=13)
  const numRows = lastRow - 1;
  const dataRange = sheet.getRange(2, 2, numRows, 14);
  const rows = dataRange.getValues();
  let updatesMade = false;

  for (let i = 0; i < rows.length; i++) {
    const cellB = String(rows[i][0] || "").trim();
    const cellO = String(rows[i][13] || "").trim();

    // Überspringen wenn: kein B vorhanden, Duplikat "-", oder Spalte O bereits befüllt
    if (!cellB || cellB === "-" || cellO !== "") {
      continue;
    }

    // JSON-Daten aus Spalte F (Personen) und Spalte G (Dauer) parsen
    let persData = {};
    let dauerData = {};
    try {
      persData = JSON.parse(rows[i][4] || "{}"); // Spalte F: Pers
      dauerData = JSON.parse(rows[i][5] || "{}"); // Spalte G: Dauer
    } catch (e) {
      rows[i][13] = "FEHLER (JSON unlesbar)";
      updatesMade = true;
      continue;
    }

    const pers = Number(persData.Pers) || 0;
    const tage = Number(dauerData.tage) || 0;
    const maxTage = Number(dauerData.maxTage) || tage;
    const hatVerlaengerung = Boolean(dauerData.verlaengerung);

    // --- DIE NEUE ROUTING-KASKADE ---

    // 1. Hartes K.O.: 1 Person (oder 0) fliegt immer raus
    if (pers <= 1) {
      rows[i][13] = "REJECT_1PERS";
      updatesMade = true;
      continue;
    }

    // 2. Sonderfall: Verlängerungsoption vorhanden bei Erstbuchung < 6 Tage (ab 2 Pers.)
    if (tage < 6 && hatVerlaengerung) {
      rows[i][13] = "KLÄRUNG_OPTION";
      updatesMade = true;
      continue;
    }

    // 3. Unter 6 Tage ohne Verlängerung, aber mehrere Personen -> Manuelle Prüfung / Mögliches Angebot
    if (tage < 6) {
      rows[i][13] = "REVIEW_SHORT"; // Oder z. B. "ANGEBOT_SHORT"
      updatesMade = true;
      continue;
    }

    // Ab hier: pers >= 2 und tage >= 6
    const istLangzeit = maxTage >= 30;

    // 4. Standard-Anfragen (2 bis 6 Personen)
    if (pers <= 6) {
      rows[i][13] = istLangzeit ? "STD_LONG" : "STD_SHORT";
      updatesMade = true;
      continue;
    }

    // 5. Großanfragen (> 6 Personen)
    if (pers > 6) {
      rows[i][13] = istLangzeit ? "BIG_LONG" : "BIG_SHORT";
      updatesMade = true;
      continue;
    }
  }

  // Änderungen gebündelt zurück in die Tabelle schreiben
  if (updatesMade) {
    dataRange.setValues(rows);
    for (let r = 2; r <= lastRow; r++) {
      sheet.setRowHeight(r, 21);
    }
    Logger.log("Routing-Filter erfolgreich ausgeführt (Ergebnisse in Spalte O).");
  }
}