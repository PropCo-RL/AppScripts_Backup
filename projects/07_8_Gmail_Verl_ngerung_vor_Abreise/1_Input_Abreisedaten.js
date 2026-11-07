/**
 * SCRIPT 1: Liest Buchungen aus 'RE' und 'VL' aus, sucht nach Folgebuchungen (Anreisedatum in Spalte K),
 * ermittelt pro Buchungsnummer das aktuellste Abreisedatum und schreibt die Daten 
 * sortiert in das Tab 'input'.
 */
function updateInputSheet() {
  const SOURCE_SPREADSHEET_ID = '1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU';
  const SOURCE_SHEETS = ['RE', 'VL'];
  const DEST_SHEET_NAME = 'input';
  
  const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  
  // Zeitfenster berechnen: Ab heute 00:00 Uhr bis heute + 3 Tage um 10:00 Uhr
  const now = new Date();
  const startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const endDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 3, 10, 0, 0);
  
  const allLatestBookings = new Map();
  const allReBookings = []; // Speicher für die Prüfung auf Anschlussbuchungen

  // SCHRITT 1: Alle Daten aus den Quellen 'RE' und 'VL' einlesen
  SOURCE_SHEETS.forEach(sheetName => {
    const sheet = sourceSpreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Warnung: Tabellenblatt "${sheetName}" wurde nicht gefunden.`);
      return;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const phone = row[0];          // Spalte A: Telefonnummer
      const name = row[1];           // Spalte B: Name
      const email = row[8];          // Spalte I: E-Mail
      const rawArrival = row[10];    // Spalte K: Anreisedatum (Anschlussbuchungsstart)
      const rawDeparture = row[11];  // Spalte L: Abreisedatum
      const apartment = row[13];     // Spalte N: Apartment
      const bookingNo = row[18];     // Spalte S: Buchungsnummer
      const statusAF = row[31];      // Spalte AF: Status (z. B. Bezahlt / Storniert)
      const rowNum = i + 1;

      // 1. FILTER: Stornierte Buchungen überspringen (Spalte AF)
      if (statusAF && String(statusAF).toLowerCase().includes('storn')) {
        continue;
      }

      if (!rawDeparture || !bookingNo) continue;
      
      let departureDate = parseDate_(rawDeparture);
      if (!departureDate) continue;

      departureDate.setHours(10, 0, 0, 0);
      const bookingKey = String(bookingNo).trim();

      const bookingEntry = {
        name: name,
        apartment: String(apartment).trim(),
        departureDate: departureDate,
        email: email,
        phone: phone,
        bookingNo: bookingKey,
        sourceInfo: `Zeile ${rowNum}, Tab ${sheetName}`
      };

      // Spätestes Abreisedatum pro Buchungsnummer sichern
      if (allLatestBookings.has(bookingKey)) {
        if (departureDate > allLatestBookings.get(bookingKey).departureDate) {
          allLatestBookings.set(bookingKey, bookingEntry);
        }
      } else {
        allLatestBookings.set(bookingKey, bookingEntry);
      }

      // Alle Anreisedaten aus Spalte K sammeln, um Folgebuchungen im selben Apartment zu ermitteln
      if (sheetName === 'RE') {
        let arrivalDateObj = parseDate_(rawArrival);
        if (arrivalDateObj) {
          allReBookings.push({
            apartment: String(apartment).trim(),
            startDate: arrivalDateObj,
            name: name,
            bookingNo: bookingKey
          });
        }
      }
    }
  });

  // SCHRITT 2: Filtern auf Abreisen im definierten 3-Tages-Fenster & Anschlussbuchung prüfen
  const matchingBookings = [];
  allLatestBookings.forEach(item => {
    if (item.departureDate >= startDate && item.departureDate <= endDate) {
      
      let nextBookingDate = '';
      let nextBookingName = '';
      let templateType = 'Option A (Verlängerung möglich)';

      // Suche nach einer Anschlussbuchung (Anreise aus Spalte K) für dasselbe Apartment am/nach dem Abreisedatum
      const foundNext = allReBookings.find(re => 
        re.apartment === item.apartment &&
        re.bookingNo !== item.bookingNo &&
        re.startDate >= new Date(item.departureDate.getFullYear(), item.departureDate.getMonth(), item.departureDate.getDate(), 0, 0, 0)
      );

      if (foundNext) {
        nextBookingDate = Utilities.formatDate(foundNext.startDate, Session.getScriptTimeZone(), 'dd.MM.yyyy');
        nextBookingName = foundNext.name;
        templateType = 'Option B (Anschlussbuchung)';
      }

      item.nextBookingDate = nextBookingDate;
      item.nextBookingName = nextBookingName;
      item.templateType = templateType;

      matchingBookings.push(item);
    }
  });

  // SCHRITT 3: Aufsteigend nach Abreisedatum sortieren (z. B. 20. Okt -> 21. Okt -> 22. Okt)
  matchingBookings.sort((a, b) => a.departureDate - b.departureDate);

  // SCHRITT 4: Daten in ein 14-spaltiges Array umwandeln
  const outputRows = [];
  matchingBookings.forEach(item => {
    const formattedDate = Utilities.formatDate(item.departureDate, Session.getScriptTimeZone(), 'dd.MM.yyyy');
    const uniqueId = `${item.bookingNo}_${formattedDate}`;

    outputRows.push([
      item.name,             // Spalte A: Name
      item.apartment,        // Spalte B: Apartment
      formattedDate,         // Spalte C: Abreise
      item.email,            // Spalte D: Email
      item.phone,            // Spalte E: Telefon
      item.bookingNo,        // Spalte F: Buchungsnummer
      item.sourceInfo,       // Spalte G: Zeile/Quelle
      uniqueId,              // Spalte H: Unique ID (Buchungsnummer_Abreisedatum)
      item.nextBookingDate,  // Spalte I: Anschlussbuchung ab Datum (aus Spalte K der Quelltabelle)
      item.nextBookingName,  // Spalte J: Anschlussbuchung Name
      item.templateType,     // Spalte K: Email Template Typ
      '',                    // Spalte L: Gmail 1 Gesendet (Status)
      '',                    // Spalte M: Gmail 1 ID
      ''                     // Spalte N: Gmail 1 Zeitstempel
    ]);
  });

  // SCHRITT 5: Tab 'input' leeren und neue Daten schreiben
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEST_SHEET_NAME);
  if (!targetSheet) {
    Logger.log(`Fehler: Tabellenblatt "${DEST_SHEET_NAME}" wurde nicht gefunden.`);
    return;
  }

  const lastRow = targetSheet.getLastRow();
  if (lastRow >= 3) {
    targetSheet.getRange(3, 1, lastRow - 2, 14).clearContent();
  }

  if (outputRows.length > 0) {
    targetSheet.getRange(3, 1, outputRows.length, 14).setValues(outputRows);
    Logger.log(`${outputRows.length} Datenzeilen erfolgreich nach "${DEST_SHEET_NAME}" geschrieben.`);
  } else {
    Logger.log(`Keine passenden Buchungen im 3-Tages-Fenster gefunden.`);
  }
}

/**
 * HILFSFUNKTION: Zuverlässiges Parsen von Datumsangaben (String oder Date-Objekt)
 */
function parseDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return new Date(val.getTime());
  if (typeof val === 'string') {
    const parts = val.trim().split('.');
    if (parts.length === 3) {
      return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}