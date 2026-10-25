function updateDauerhaftBelegt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Tabs: 'M' für die Apartment-Übersicht, 'cal' für die Buchungen
  const mainSheet = ss.getSheetByName('M'); 
  const calSheet = ss.getSheetByName('cal');
  
  if (!mainSheet || !calSheet) {
    Logger.log('FEHLER: Tabellenblatt "M" oder "cal" nicht gefunden!');
    return;
  }

  // 1. Zeitfenster definieren: heute bis heute + 60 Tage (als Tages-Index, zeitzonenunabhängig)
  const now = new Date();
  const todayDayIndex = Math.floor(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) / 86400000);
  const targetEndDayIndex = todayDayIndex + 60;

  Logger.log(`HEUTE: ${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()} (Tag: ${todayDayIndex})`);
  Logger.log(`ZIEL (+60 Tage): Tag ${targetEndDayIndex}`);

  // Hilfsfunktion: Wandelt String oder Date in Tages-Index um
  function toDayIndex(val) {
    if (!val) return null;
    if (val instanceof Date) {
      return Math.floor(Date.UTC(val.getFullYear(), val.getMonth(), val.getDate()) / 86400000);
    }
    const str = String(val).trim();
    const m = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) {
      return Math.floor(Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)) / 86400000);
    }
    return null;
  }

  // Hilfsfunktion: Bereinigt Namen von Leerzeichen & Sonderzeichen
  function cleanName(str) {
    if (!str) return '';
    return String(str).replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  // 2. Buchungsdaten aus Tab 'cal' einlesen
  const calData = calSheet.getDataRange().getValues();
  const bookingsMap = {};

  for (let i = 1; i < calData.length; i++) {
    const rawName = calData[i][0];
    const key = cleanName(rawName);
    const startDay = toDayIndex(calData[i][1]);
    const endDay = toDayIndex(calData[i][2]);

    if (!key || startDay === null || endDay === null) continue;

    if (!bookingsMap[key]) bookingsMap[key] = [];
    bookingsMap[key].push({ start: startDay, end: endDay });
  }

  Logger.log(`Apartments mit Buchungen in Tab "cal": ${Object.keys(bookingsMap).length}`);

  // 3. Tab 'M' einlesen
  const totalRows = mainSheet.getLastRow();
  if (totalRows < 2) {
    Logger.log('Tab "M" enthält keine Datenzeilen.');
    return;
  }

  // Spalte A (1) aus Tab 'M' holen
  const siteNames = mainSheet.getRange(2, 1, totalRows - 1, 1).getValues();
  const output = [];
  let hits = 0;

  for (let r = 0; r < siteNames.length; r++) {
    const rawName = siteNames[r][0];
    const key = cleanName(rawName);
    const intervals = bookingsMap[key] || [];

    if (intervals.length === 0) {
      output.push(['']);
      continue;
    }

    // Sortieren nach Anreise
    intervals.sort((a, b) => a.start - b.start);

    // Buchungsintervalle nahtlos verschmelzen
    const merged = [];
    let cur = { start: intervals[0].start, end: intervals[0].end };

    for (let j = 1; j < intervals.length; j++) {
      if (intervals[j].start <= cur.end) {
        cur.end = Math.max(cur.end, intervals[j].end);
      } else {
        merged.push(cur);
        cur = { start: intervals[j].start, end: intervals[j].end };
      }
    }
    merged.push(cur);

    // Prüfen, ob durchgehender Block von <= HEUTE bis >= HEUTE + 60 TAGE existiert
    let isCovered = false;
    for (const b of merged) {
      if (b.start <= todayDayIndex && b.end >= targetEndDayIndex) {
        isCovered = true;
        break;
      }
    }

    if (isCovered) {
      Logger.log(`[TREFFER] Zeile ${r + 2}: "${rawName}" ist dauerhaft belegt.`);
      output.push(['x']);
      hits++;
    } else {
      output.push(['']);
    }
  }

  // 4. In Spalte AS (Spalte 45) des Tabs 'M' schreiben
  mainSheet.getRange(2, 45, output.length, 1).setValues(output);
  Logger.log(`ERGEBNIS: ${hits} Treffer erfolgreich in Spalte AS von Tab "M" eingetragen.`);
}