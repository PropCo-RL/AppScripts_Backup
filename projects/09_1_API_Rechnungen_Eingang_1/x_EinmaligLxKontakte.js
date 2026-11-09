function syncLexofficeContactsToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetContacts = ss.getSheetByName("Kontakte");
  
  if (!sheetContacts) {
    sheetContacts = ss.insertSheet("Kontakte");
  }
  
  sheetContacts.clear();
  sheetContacts.appendRow([
    "Firma / Name", 
    "Contact ID (UUID)", 
    "Standard MwSt % (Manuell)", 
    "Standard Kategorie Name (Manuell)", 
    "Standard Kategorie UUID (Manuell)"
  ]);
  sheetContacts.getRange("A1:E1").setFontWeight("bold");

  let page = 0;
  let totalPages = 1;
  let allVendors = [];

  // Automatischer Durchlauf aller API-Seiten
  do {
    const url = `https://api.lexware.io/v1/contacts?size=250&page=${page}`;
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
      const data = JSON.parse(response.getContentText());

      if (data.content && data.content.length > 0) {
        data.content.forEach(c => {
          // Nur Lieferanten (vendor) filtern
          const isVendor = c.roles && c.roles.vendor;

          if (isVendor) {
            let name = c.company ? c.company.name : "";
            if (!name && c.person) {
              name = `${c.person.firstName || ""} ${c.person.lastName || ""}`.trim();
            }
            if (name) {
              allVendors.push([name, c.id, "", "", ""]);
            }
          }
        });
      }

      totalPages = data.totalPages || 1;
      page++;

    } catch (e) {
      Logger.log(`Fehler auf Seite ${page}: ` + e.toString());
      break;
    }

  } while (page < totalPages);

  // Alle gesammelten Lieferanten in einem Rutsch ins Sheet schreiben
  if (allVendors.length > 0) {
    sheetContacts.getRange(2, 1, allVendors.length, 5).setValues(allVendors);
    Logger.log(`Fertig! Insgesamt ${allVendors.length} Lieferanten über ${page} Seiten hinweg eingetragen.`);
  } else {
    Logger.log("Keine Lieferanten gefunden.");
  }
}