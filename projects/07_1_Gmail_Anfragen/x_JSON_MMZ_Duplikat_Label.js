function markiereMMZDuplikateMitLabel() {
  var sheetId = "1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw";
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("MMZ");
  if (!sheet) return;

  const COL_K = 11;
  const COL_L = 12;
  const COL_M = 13;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Header für Spalte M setzen
  var headerM = sheet.getRange(1, COL_M).getValue();
  if (!headerM) {
    sheet.getRange(1, COL_M).setValue("Status Duplikat-Label");
    sheet.getRange(1, COL_M).setFontWeight("bold");
    sheet.getRange("M:M").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  const labelName = "Duplikat";
  const label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);

  const startRow = 2;
  const numRows = lastRow - startRow + 1;

  // Spalten K bis M einlesen (K=0, L=1, M=2)
  const rangeKLM = sheet.getRange(startRow, COL_K, numRows, 3);
  const data = rangeKLM.getValues();

  var updatesMade = false;

  for (let i = 0; i < data.length; i++) {
    const currentRowIndex = startRow + i;
    const valK = data[i][0] ? data[i][0].toString().trim() : "";
    const valL = data[i][1] ? data[i][1].toString().trim() : "";
    const valM = data[i][2] ? data[i][2].toString().trim() : "";

    // Zeilen überspringen, die bereits bearbeitet wurden oder wo noch kein L-Wert vorhanden ist
    if (valM !== "" || valK === "" || valL === "") {
      continue;
    }

    // FALL 1: Es ist ein Duplikat (Spalte L hat "-")
    if (valL === "-") {
      try {
        const jsonData = JSON.parse(valK);
        const link = jsonData.link;

        if (link) {
          // Extrahiert die Hex-ID am Ende des Gmail-Links
          const match = link.match(/[#/]([a-f0-9]{16,})$/i);
          if (match && match[1]) {
            const threadOrMsgId = match[1];
            let labelApplied = false;

            // 1. Als Thread versuchen
            try {
              const thread = GmailApp.getThreadById(threadOrMsgId);
              if (thread) {
                thread.addLabel(label);
                labelApplied = true;
              }
            } catch (gmailErr) {
              // 2. Fallback: Als Einzel-Nachricht versuchen
              try {
                const message = GmailApp.getMessageById(threadOrMsgId);
                if (message) {
                  message.getThread().addLabel(label);
                  labelApplied = true;
                }
              } catch (fallbackErr) {
                Logger.log(`Fehler bei Zeile ${currentRowIndex}: ${fallbackErr.message}`);
              }
            }

            // Status in Spalte M eintragen
            data[i][2] = labelApplied ? "Duplikat" : "Fehler (ID nicht gefunden)";
            updatesMade = true;
          } else {
            data[i][2] = "Fehler (Keine Link-ID)";
            updatesMade = true;
          }
        } else {
          data[i][2] = "Fehler (Kein Link)";
          updatesMade = true;
        }
      } catch (parseErr) {
        Logger.log(`Ungültiges JSON in Zeile ${currentRowIndex}: ${parseErr.message}`);
        data[i][2] = "Fehler (JSON ungültig)";
        updatesMade = true;
      }
    } 
    // FALL 2: Es ist ein Unique-Eintrag (Spalte L ist NICHT "-")
    else {
      // Kein Gmail-Aufruf! Nur als "Kein Duplikat" markieren, damit die Zeile künftig übersprungen wird
      data[i][2] = "Kein Duplikat";
      updatesMade = true;
    }
  }

  // Alles in einem einzigen Schreibvorgang zurück in Spalte M schreiben
  if (updatesMade) {
    const columnMValues = data.map(row => [row[2]]);
    sheet.getRange(startRow, COL_M, numRows, 1).setValues(columnMValues);
  }
}