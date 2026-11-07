/**
 * Bereinigt die MMZ-Tabelle und filtert Duplikate heraus.
 * Erkennt Anfragen desselben Gasts über Name, Zeitraum und Inhalt.
 * Wertet unvollständige Zeilen automatisch auf, sobald die Kontaktdaten nachgeliefert werden.
 */
function verarbeiteUndBereinigeMMZ() {
  const SHEET_NAME = "MMZ";
  const LOOKBACK_ROWS = 150; // Prüft die letzten 150 Zeilen
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Header in Spalte B sicherstellen
  const headerB = sheet.getRange(1, 2).getValue();
  if (!headerB || headerB.toString().trim() === "") {
    sheet.getRange(1, 2).setValue("Bereinigt / Unique");
    sheet.getRange(1, 2).setFontWeight("bold");
    sheet.getRange("B:B").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  const startRow = Math.max(2, lastRow - LOOKBACK_ROWS + 1);
  const numRows = lastRow - startRow + 1;
  const rangeAB = sheet.getRange(startRow, 1, numRows, 2);
  const data = rangeAB.getValues();

  // Speichert: fingerprint -> { index: number, score: number }
  const registered = new Map();

  // 1. Bereits vorhandene Uniques in Spalte B indexieren
  for (let i = 0; i < data.length; i++) {
    const cellB = data[i][1] ? data[i][1].toString().trim() : "";
    if (cellB !== "" && cellB !== "-") {
      try {
        const parsed = JSON.parse(cellB);
        const fp = generateFingerprint(parsed);
        registered.set(fp, { index: i, score: getKontaktdatenScore(parsed) });
      } catch (e) {
        // Fallback für nicht-JSON Zellen
        registered.set(cellB, { index: i, score: 0 });
      }
    }
  }

  // 2. Offene Zeilen verarbeiten (Spalte A gefüllt, Spalte B leer)
  let updatesMade = false;

  for (let i = 0; i < data.length; i++) {
    const cellA = data[i][0] ? data[i][0].toString().trim() : "";
    const cellB = data[i][1] ? data[i][1].toString().trim() : "";

    if (cellA !== "" && cellB === "") {
      try {
        const jsonObj = JSON.parse(cellA);
        const fp = generateFingerprint(jsonObj);
        const currentScore = getKontaktdatenScore(jsonObj);

        // Fall A: Dieser Fingerprint existiert bereits
        if (registered.has(fp)) {
          const prev = registered.get(fp);

          // UPGRADE: Neuer Datensatz hat MEHR Kontaktdaten als der frühere
          if (currentScore > prev.score) {
            // Alte unvollständige Zeile in Spalte B auf "-" setzen
            data[prev.index][1] = "-";
            // Neue, vollständige Zeile übernehmen
            data[i][1] = cellA;
            // Map aktualisieren
            registered.set(fp, { index: i, score: currentScore });
          } else {
            // Neuer Datensatz ist gleichwertig oder schlechter -> als Duplikat abweisen
            data[i][1] = "-";
          }
        } 
        // Fall B: Komplett neuer Fingerprint
        else {
          data[i][1] = cellA;
          registered.set(fp, { index: i, score: currentScore });
        }

        updatesMade = true;
      } catch (e) {
        // Ungültiges JSON in Spalte A
        data[i][1] = "-";
        updatesMade = true;
      }
    }
  }

  // 3. Nur mit einem einzigen Schreibvorgang zurückschreiben
  if (updatesMade) {
    rangeAB.setValues(data);
  }
}

/**
 * Bewertet den Vollständigkeitsgrad der Kontaktdaten:
 * 2 = Telefon UND E-Mail unmaskiert vorhanden
 * 1 = Mindestens E-Mail ODER Telefon vorhanden
 * 0 = Keine brauchbaren Kontaktdaten (nur Maskierung oder leer)
 */
function getKontaktdatenScore(item) {
  const phone = String(item.telefonnummer || "").trim();
  const email = String(item.email || "").trim();

  const hatPhone = phone !== "" && 
                   !phone.includes("••") && 
                   !phone.includes("&bull;") && 
                   phone.replace(/[^\d]/g, "").length >= 6;

  const hatEmail = email !== "" && 
                   !email.includes("••") && 
                   !email.includes("&bull;") && 
                   email.includes("@");

  if (hatPhone && hatEmail) return 2;
  if (hatPhone || hatEmail) return 1;
  return 0;
}

/**
 * Erzeugt einen stabilen Fingerprint für MMZ-Anfragen.
 * Ignoriert wechselnde Inserats-IDs und fehlerhafte Firmennummern.
 */
function generateFingerprint(item) {
  // Name säubern: "Frau", "Herr", Umlaute normalisieren, Sonderzeichen entfernen
  const rawName = String(item.gastName || "").toLowerCase();
  const cleanName = rawName
    .replace(/^(frau|herr)\s+/i, "")
    .replace(/[^a-z0-9äöüß]/g, "");

  // Zeitraum standardisieren: Nur die reinen Ziffern (z. B. "3110202607112026")
  const period = String(item.zeitraum || "").replace(/[^0-9]/g, "");

  // Personenanzahl
  const persons = String(item.personen || "").trim();

  // Nachricht säubern: Erste 60 Zeichen ohne Leerzeichen/Satzzeichen
  const cleanMsg = String(item.nachricht || "")
    .toLowerCase()
    .replace(/[^a-z0-9äöüß]/g, "")
    .substring(0, 60);

  return `${cleanName}|${period}|${persons}|${cleanMsg}`;
}