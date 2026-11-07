function rWeeklyInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetRE = ss.getSheetByName("wRE"); 
  const sheetDB = ss.getSheetByName("db");
  const sheetLXLog = ss.getSheetByName("lxLog");

  if (!sheetRE || !sheetDB || !sheetLXLog) {
    Logger.log("Eines der Blätter (wRE, db oder lxLog) wurde nicht gefunden!");
    return;
  }

  const stichTagVal = sheetRE.getRange("C2").getValue();
  if (!stichTagVal) return;
  const stichTag = new Date(stichTagVal).getTime();

  // 1. IDs aus wRE Spalte B (ID = Index 0 in diesem Array)
  const lastRowB = sheetRE.getRange("B:B").getValues().filter(String).length + 3; 
  if (lastRowB < 4) return; 
  const idsInRE = sheetRE.getRange("B4:B" + lastRowB).getValues();

  // 2. DATEN LADEN
  const dbRaw = sheetDB.getRange("A4:W" + Math.max(sheetDB.getLastRow(), 4)).getValues();
  const lxLogRaw = sheetLXLog.getRange("A2:Q" + Math.max(sheetLXLog.getLastRow(), 2)).getValues();

  // Hilfsfunktionen
  const formatDate = (dateVal) => {
    if (!dateVal) return "-";
    let d = new Date(dateVal);
    return isNaN(d.getTime()) ? "-" : d.toLocaleDateString('de-DE', {day: 'numeric', month: 'short'});
  };

  const normalize = (str) => {
    if (!str) return "";
    return str.toString().toLowerCase()
      .replace(/[äàáâ]/g, "ae").replace(/[öòóô]/g, "oe").replace(/[üùúû]/g, "ue").replace(/ß/g, "ss")
      .replace(/[^a-z0-9]/g, "").trim();
  };

  // 3. lxLog VORBEREITEN & GRUPPIEREN
  // Wir gruppieren lxLog nach Rechnungs-UUID (Spalte A), um Items zusammenzufassen
  const rechnungsMap = {}; // Key: Rechnungs-UUID
  const gueltigeStatus = ["open", "paid", "overdue"];

  lxLogRaw.forEach(row => {
    const uuid = row[0]; // Spalte A: Rechnungs UUID
    const status = (row[11] || "").toString().toLowerCase(); // Spalte L: Status

    if (uuid && gueltigeStatus.includes(status)) {
      if (!rechnungsMap[uuid]) {
        rechnungsMap[uuid] = {
          normName: normalize(row[7]), // Spalte H: Name
          zeitraumEnde: row[9] ? new Date(row[9]).getTime() : 0, // Spalte J: Ende
          zeitraumEndeRaw: row[9],
          status: row[11],
          items: []
        };
      }
      // Item Daten sammeln: N=Pos(13), O=Name(14), P=Menge(15), Q=Preis(16)
      rechnungsMap[uuid].items.push({
        pos: parseInt(row[13]) || 999,
        name: row[14],
        menge: row[15],
        preis: row[16]
      });
    }
  });

  // Items innerhalb jeder Rechnung nach Position sortieren
  Object.values(rechnungsMap).forEach(re => {
    re.items.sort((a, b) => a.pos - b.pos);
  });

  // 4. VERARBEITUNG
  const mainResults = [];
  const itemResults = [];

  idsInRE.forEach(row => {
    const searchId = row[0].toString().trim();
    if (!searchId) {
      mainResults.push(["", "", "", ""]);
      itemResults.push([]);
      return;
    }

    let gefundenerName = "";
    let aufenthaltEnde = "-";
    let rechnungZeitraumEnde = "-";
    let status = "Keine RE";
    let gefundeneItems = [];

    // Suche in DB (identisch zum vorherigen Stand)
    for (let i = 0; i < dbRaw.length; i++) {
      const propId = dbRaw[i][5].toString().trim();
      const von = new Date(dbRaw[i][3]).getTime();
      const bis = new Date(dbRaw[i][4]).getTime();
      const source = (dbRaw[i][22] || "").toString();

      if (propId === searchId && von <= stichTag && bis > stichTag) {
        if (source !== "BookingCom" && source !== "AirbnbIntegration") {
          let nameRaw = (dbRaw[i][13] || "").toString();
          if (!nameRaw.toLowerCase().startsWith("re ")) {
            gefundenerName = nameRaw.split(/[-(–]/)[0].trim();
            aufenthaltEnde = formatDate(dbRaw[i][4]);
            break;
          }
        }
      }
    }

    // Match mit gruppiertem lxLog
    if (gefundenerName) {
      const normName = normalize(gefundenerName);
      let neuestesEnde = -1;
      let besteRE = null;

      // Finde die Rechnung mit dem neuesten Zeitraum-Enddatum für diesen Mieter
      for (const uuid in rechnungsMap) {
        const re = rechnungsMap[uuid];
        if (re.normName.includes(normName) || normName.includes(re.normName)) {
          if (re.zeitraumEnde > neuestesEnde) {
            neuestesEnde = re.zeitraumEnde;
            besteRE = re;
          }
        }
      }

      if (besteRE) {
        status = besteRE.status;
        rechnungZeitraumEnde = formatDate(besteRE.zeitraumEndeRaw);
        
        // Items in flaches Array umwandeln: Name, Menge, Preis, Name, Menge, Preis...
        besteRE.items.forEach(it => {
          gefundeneItems.push(it.name || "-", it.menge || 0, it.preis || 0);
        });
      }
    }

    mainResults.push([gefundenerName, aufenthaltEnde, rechnungZeitraumEnde, status]);
    itemResults.push(gefundeneItems);
  });

  // 5. SCHREIBEN
  sheetRE.getRange("C4:F150").clearContent(); 
  if (mainResults.length > 0) {
    sheetRE.getRange(4, 3, mainResults.length, 4).setValues(mainResults);
  }

  sheetRE.getRange("J4:Z150").clearContent(); 
  if (itemResults.length > 0) {
    for (let i = 0; i < itemResults.length; i++) {
      if (itemResults[i].length > 0) {
        // Schreibt Name, Menge, Preis ab Spalte J (10)
        sheetRE.getRange(4 + i, 10, 1, itemResults[i].length).setValues([itemResults[i]]);
      }
    }
  }
}