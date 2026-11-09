function syncPostingCategoriesToLxTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Tabellenblatt "lx" auswählen oder automatisch neu erstellen
  let sheetLx = ss.getSheetByName("lx");
  if (!sheetLx) {
    sheetLx = ss.insertSheet("lx");
  }
  
  // 2. Inhalt leeren und Kopfzeilen setzen
  sheetLx.clear();
  sheetLx.appendRow(["Kategorie Name", "ID (UUID)", "Typ", "Gruppe", "Kontakt Erforderlich", "Split Erlaubt"]);
  sheetLx.getRange("A1:F1").setFontWeight("bold");

  // 3. API-Abruf bei Lexware (posting-categories Endpoint)
  const url = "https://api.lexware.io/v1/posting-categories";
  const options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + LX_KEY,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const categories = JSON.parse(response.getContentText());

    if (!Array.isArray(categories) || categories.length === 0) {
      Logger.log("Keine Kategorien von Lexware zurückgemeldet.");
      return;
    }

    // 4. Alle Kategorien in das Tabellenblatt "lx" eintragen
    let rows = [];
    categories.forEach(cat => {
      rows.push([
        cat.name || "",
        cat.id || "",
        cat.type || "",               // "outgo" (Ausgaben) oder "income" (Einnahmen)
        cat.groupName || "",
        cat.contactRequired ? "Ja" : "Nein",
        cat.splitAllowed ? "Ja" : "Nein"
      ]);
    });

    // Effizientes Schreiben in einem Schwung
    sheetLx.getRange(2, 1, rows.length, 6).setValues(rows);
    Logger.log(`${rows.length} Lexware-Kategorien erfolgreich in das Tabellenblatt 'lx' eingetragen.`);

  } catch (error) {
    Logger.log("Fehler beim Abrufen der posting-categories: " + error.toString());
  }
}