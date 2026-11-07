function markiereDuplikateMitLabel() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Spalten definieren (1-basiert):
  // A = 1 (JSON), B = 2 ("-"), AF = 32 (Status "x")
  const COL_A = 1;
  const COL_B = 2;
  const COL_AF = 32;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // Beenden, falls keine Daten vorhanden sind

  // Vorhandenes Label abrufen
  const labelName = "Duplikat";
  const label = GmailApp.getUserLabelByName(labelName);
  
  if (!label) {
    Logger.log(`Label "${labelName}" wurde in Gmail nicht gefunden.`);
    return;
  }

  // Daten ab Zeile 2 bis zur letzten Zeile lesen
  const startRow = 2;
  const numRows = lastRow - startRow + 1;
  const totalCols = COL_AF;

  const data = sheet.getRange(startRow, 1, numRows, totalCols).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const currentRowIndex = startRow + i;

    const valA = row[COL_A - 1];   // Spalte A (JSON)
    const valB = row[COL_B - 1];   // Spalte B (Duplikat-Kennzeichnung "-")
    const valAF = row[COL_AF - 1]; // Spalte AF (Status/Marker "x")

    // Bedingungen:
    // 1. Spalte A ist nicht leer
    // 2. Spalte B ist genau "-"
    // 3. Spalte AF ist leer (noch kein "x")
    const isANotEmpty = valA !== null && valA.toString().trim() !== "";
    const isBDuplicate = valB !== null && valB.toString().trim() === "-";
    const isAFEmpty = valAF === null || valAF.toString().trim() === "";

    if (isANotEmpty && isBDuplicate && isAFEmpty) {
      try {
        const jsonData = JSON.parse(valA.toString());
        const link = jsonData.link;

        if (link) {
          // Extrahiert die Hex-ID am Ende des Gmail-Links
          const match = link.match(/[#/]([a-f0-9]{16,})$/i);
          
          if (match && match[1]) {
            const threadId = match[1];
            
            try {
              // Thread laden und Label zuweisen
              const thread = GmailApp.getThreadById(threadId);
              if (thread) {
                thread.addLabel(label);
                sheet.getRange(currentRowIndex, COL_AF).setValue("x");
              }
            } catch (gmailErr) {
              // Fallback: Falls die ID direkt eine Nachrichten-ID ist
              try {
                const message = GmailApp.getMessageById(threadId);
                if (message) {
                  message.getThread().addLabel(label);
                  sheet.getRange(currentRowIndex, COL_AF).setValue("x");
                }
              } catch (fallbackErr) {
                Logger.log(`Fehler bei Zeile ${currentRowIndex}: ${fallbackErr.message}`);
              }
            }
          }
        }
      } catch (parseErr) {
        Logger.log(`Ungültiges JSON in Zeile ${currentRowIndex}: ${parseErr.message}`);
      }
    }
  }
}