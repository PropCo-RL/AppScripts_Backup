// =================================================================
// GMAIL LABELING: Markiert Original-Anfragen mit "ENTWURF"
// Liest Spalte T (Entwurf-Status) und markiert in Spalte AG mit "x"
// =================================================================
function markiereEntwuerfeMitLabel() {
  const SHEET_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const LABEL_NAME = "ENTWURF"; // Name deines bestehenden Labels

  // Spalten definieren (1-basiert):
  // T = 20 (Entwurf-Status), AG = 33 (Marker "x")
  const COL_T_INDEX = 19;  // Spalte T (0-basierter Index)
  const COL_AG_INDEX = 32; // Spalte AG (0-basierter Index)
  const TOTAL_COLS = 33;   // Bis Spalte AG einlesen

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
  if (lastRow < 2) return;

  // Gmail-Label abrufen
  const label = GmailApp.getUserLabelByName(LABEL_NAME);
  if (!label) {
    Logger.log(`FEHLER: Label "${LABEL_NAME}" wurde in Gmail nicht gefunden. Bitte prüfe die Schreibweise.`);
    return;
  }

  // Header für Spalte AG (Spalte 33) sicherstellen
  const headerAG = sheet.getRange(1, 33);
  if (!headerAG.getValue()) {
    headerAG.setValue("Gelabelt");
    headerAG.setFontWeight("bold");
    sheet.getRange("AG:AG").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Spalten A bis AG einlesen (33 Spalten: A=0 bis AG=32)
  const numRows = lastRow - 1;
  const range = sheet.getRange(2, 1, numRows, TOTAL_COLS);
  const rows = range.getValues();
  let updatesMade = false;

  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    const cellA = String(rows[i][0] || "").trim();           // Spalte A: Roh-JSON
    const cellB = String(rows[i][1] || "").trim();           // Spalte B: Bereinigtes JSON
    const cellT = String(rows[i][COL_T_INDEX] || "").trim();  // Spalte T: E-Mail Status
    const cellAG = String(rows[i][COL_AG_INDEX] || "").trim(); // Spalte AG: Marker "x"

    // Bedingungen:
    // 1. In Spalte T steht ein erstellter Entwurf ("ENTWURF...")
    // 2. In Spalte AG steht noch KEIN "x"
    const hatEntwurf = cellT.toUpperCase().startsWith("ENTWURF");
    const nochNichtGelabelt = (cellAG === "");

    if (hatEntwurf && nochNichtGelabelt) {
      // JSON bevorzugt aus B (Bereinigt), sonst A (Roh) holen
      const jsonSource = (cellB && cellB !== "-") ? cellB : cellA;
      if (!jsonSource) continue;

      let link = "";
      try {
        const parsed = JSON.parse(jsonSource);
        link = String(parsed.link || "").trim();
      } catch (err) {
        Logger.log(`Zeile ${rowNum}: JSON konnte nicht gelesen werden.`);
        continue;
      }

      if (!link) {
        Logger.log(`Zeile ${rowNum}: Kein Link im JSON vorhanden.`);
        continue;
      }

      // Hex-ID aus dem Gmail-Link extrahieren (z. B. ".../#all/1a163d0995633778")
      const match = link.match(/[#/]([a-f0-9]{16,})$/i);
      if (match && match[1]) {
        const mailId = match[1];
        let labelErfolgreich = false;

        // 1. Versuch: Als Thread laden
        try {
          const thread = GmailApp.getThreadById(mailId);
          if (thread) {
            thread.addLabel(label);
            labelErfolgreich = true;
          }
        } catch (threadErr) {
          // 2. Fallback: Als einzelne Message laden
          try {
            const message = GmailApp.getMessageById(mailId);
            if (message) {
              message.getThread().addLabel(label);
              labelErfolgreich = true;
            }
          } catch (msgErr) {
            Logger.log(`Zeile ${rowNum}: Konnte Gmail-Objekt (${mailId}) nicht abrufen.`);
          }
        }

        // Bei Erfolg "x" in Spalte AG setzen
        if (labelErfolgreich) {
          rows[i][COL_AG_INDEX] = "x";
          updatesMade = true;
          Logger.log(`Zeile ${rowNum}: Label "${LABEL_NAME}" zugewiesen -> 'x' in Spalte AG.`);
        }
      }
    }
  }

  // Änderungen gebündelt zurückschreiben
  if (updatesMade) {
    range.setValues(rows);
    SpreadsheetApp.flush();
    Logger.log("=== ALLE ENTWÜRFE ERFOLGREICH IN GMAIL GELABELT (STATUS IN SPALTE AG) ===");
  }
}