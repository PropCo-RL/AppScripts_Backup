// =================================================================
// 1. DUPLIKATS-PRÜFUNG (Wird als Erster Schritt ausgeführt)
// =================================================================
function checkDuplicateRequests() {
  const SHEET1_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  
  let sheet;
  try {
    sheet = SpreadsheetApp.openById(SHEET1_ID).getSheetByName('DMZ');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabelle: " + e.toString());
    return;
  }

  if (!sheet) {
    Logger.log("FEHLER: Tabellenblatt 'DMZ' nicht gefunden.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  Logger.log("--- DUPLIKATS-PRÜFUNG START ---");

  // Map speichert bereits gesehene Anfragen-Kombinationen
  let seenLeads = new Map();
  let duplicateCount = 0;

  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    const currentStatus = String(data[i][17] || "").trim(); // Spalte R (Status)

    // Nur Zeilen prüfen, die in Spalte R noch keinen Eintrag haben
    if (currentStatus !== "") {
      continue;
    }

    const name = String(data[i][2] || "").trim().toLowerCase();     // Spalte C
    const phone = String(data[i][3] || "").trim().toLowerCase();    // Spalte D
    const period = String(data[i][4] || "").trim().toLowerCase();   // Spalte E
    const persons = String(data[i][5] || "").trim();               // Spalte F

    // Eindeutigen Fingerabdruck erzeugen
    const leadKey = `${name}|${phone}|${period}|${persons}`;

    // Falls die Kombination unvollständig ist, überspringen
    if (!name && !phone && !period) {
      continue;
    }

    if (seenLeads.has(leadKey)) {
      // Es ist eine Dopplung -> Markierung in Spalte S schreiben (Spalte 19)
      sheet.getRange(rowNum, 19).setValue("Dopplung");
      duplicateCount++;
      Logger.log(`Zeile ${rowNum} als Dopplung markiert (Identisch mit Zeile ${seenLeads.get(leadKey)})`);
    } else {
      // Erste Erfassung dieser Kombination -> im Speicher registrieren
      seenLeads.set(leadKey, rowNum);
      
      // Falls vorher evtl. "Dopplung" in S stand, dieses Feld leeren
      const currentDupCol = String(data[i][18] || "").trim(); // Spalte S (Index 18)
      if (currentDupCol === "Dopplung") {
        sheet.getRange(rowNum, 19).clearContent();
      }
    }
  }

  Logger.log(`--- DUPLIKATS-PRÜFUNG BEENDET (${duplicateCount} Dopplungen gefunden) ---`);
}