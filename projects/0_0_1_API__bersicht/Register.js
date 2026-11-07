const REGISTRY_CONFIG = {
  TARGET_SS_ID: "1rLFAl4pdgADo6j_vaDxmPMd8kAHL0V8Pqp-DjeXBxLo",
  TAB_NAME: "Übersicht",
  MAX_STRING_LENGTH: 50
};

// ==========================================
// 1. ORCHESTRATOR (HAUPTFUNKTION)
// ==========================================
function registerThisScript() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptId = ScriptApp.getScriptId();
  const sheetId = ss ? ss.getId() : "";

  // Datenbausteine isoliert ermitteln
  const folderPath = getParentFolderName(sheetId);
  const { headersOnly, singleSample } = extractSpreadsheetSchemas(ss, REGISTRY_CONFIG.MAX_STRING_LENGTH);

  // Datenzeile (A bis H) vorbereiten
  const rowData = [
    ss ? ss.getName() : "Standalone",
    "Google Sheet",
    sheetId,
    ss ? ss.getUrl() : "",
    scriptId,
    folderPath,
    JSON.stringify(headersOnly),
    JSON.stringify(singleSample)
  ];

  // In Ziel-Sheet upserten
  saveRowToRegistry(REGISTRY_CONFIG, scriptId, rowData);
}

// ==========================================
// 2. UNTERFUNKTION: ORDNERPFAD ERMITTELN
// ==========================================
function getParentFolderName(sheetId) {
  if (!sheetId) return "";
  try {
    const parents = DriveApp.getFileById(sheetId).getParents();
    return parents.hasNext() ? parents.next().getName() : "Root";
  } catch (e) {
    return "Kein Zugriff / Root";
  }
}

// ==========================================
// 3. UNTERFUNKTION: TABS & SCHEMAS EXTRAHIEREN
// ==========================================
function extractSpreadsheetSchemas(ss, maxLen) {
  const headersOnly = {};
  const singleSample = {};

  if (!ss) return { headersOnly, singleSample };

  for (const sheet of ss.getSheets()) {
    const data = sheet.getDataRange().getValues();
    const tabName = sheet.getName();

    if (data.length === 0) continue;

    // Header aus Zeile 1
    headersOnly[tabName] = data[0].map(h => String(h ?? ""));

    // Beispielwerte aus Zeile 2 (gekürzt)
    if (data.length > 1) {
      singleSample[tabName] = data[1].map(cell => {
        const str = String(cell ?? "");
        return str.length > maxLen ? str.substring(0, maxLen - 3) + "..." : cell;
      });
    } else {
      singleSample[tabName] = [];
    }
  }

  return { headersOnly, singleSample };
}

// ==========================================
// 4. UNTERFUNKTION: SHEET-UPSERT (SCHREIBEN/ÜBERSCHREIBEN)
// ==========================================
function saveRowToRegistry(config, scriptId, rowData) {
  const targetSheet = SpreadsheetApp.openById(config.TARGET_SS_ID).getSheetByName(config.TAB_NAME);
  const lastRow = targetSheet.getLastRow();
  let targetRowIndex = -1;

  // Prüfen, ob Script-ID (Spalte E = Index 5) bereits existiert
  if (lastRow > 1) {
    const ids = targetSheet.getRange(2, 5, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === scriptId) {
        targetRowIndex = i + 2;
        break;
      }
    }
  }

  // Schreiben: Update oder Append
  if (targetRowIndex !== -1) {
    targetSheet.getRange(targetRowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    targetSheet.appendRow(rowData);
    targetRowIndex = targetSheet.getLastRow();
  }

  // Optik sicherstellen: 21px Höhe & CLIP auf Spalten G und H
  targetSheet.setRowHeight(targetRowIndex, 21);
  targetSheet.getRange(targetRowIndex, 7, 1, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}