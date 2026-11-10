const REGISTRY_CONFIG = {
TARGET_SS_ID: "1rLFAl4pdgADo6j_vaDxmPMd8kAHL0V8Pqp-DjeXBxLo",
TARGET_FOLDER_ID: "1jqJ74Ix5lv1QVk8qFmyYTaCLwVnXOaPk",
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

// 1. Zuerst in Zielordner verschieben (falls noch nicht dort)
moveToTargetFolder(sheetId, REGISTRY_CONFIG.TARGET_FOLDER_ID);

// 2. Metadaten & Schemas nach dem Verschieben ermitteln
const folderPath = getParentFolderName(sheetId);
const { headersOnly, singleSample } = extractSpreadsheetSchemas(ss, REGISTRY_CONFIG.MAX_STRING_LENGTH);

// 3. Datenzeile (A bis H) vorbereiten
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

// 4. In Ziel-Sheet speichern (Upsert)
saveRowToRegistry(REGISTRY_CONFIG, scriptId, rowData);
}

// ==========================================
// 2. UNTERFUNKTION: IN ZIELORDNER VERSCHIEBEN
// ==========================================
function moveToTargetFolder(sheetId, targetFolderId) {
if (!sheetId || !targetFolderId) return;

try {
const file = DriveApp.getFileById(sheetId);
const targetFolder = DriveApp.getFolderById(targetFolderId);

// Prüfen, ob die Datei bereits im Zielordner liegt
const parents = file.getParents();
while (parents.hasNext()) {
if (parents.next().getId() === targetFolderId) {
return; // Liegt bereits dort -> nichts tun
}
}

file.moveTo(targetFolder);
Logger.log("Datei erfolgreich in Zielordner verschoben.");
} catch (e) {
Logger.log("Fehler beim Verschieben in den Zielordner: " + e.toString());
}
}

// ==========================================
// 3. UNTERFUNKTION: ORDNERNAME ERMITTELN
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
// 4. UNTERFUNKTION: TABS & SCHEMAS EXTRAHIEREN
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
// 5. UNTERFUNKTION: SHEET-UPSERT (SCHREIBEN/UPDATE)
// ==========================================
function saveRowToRegistry(config, scriptId, rowData) {
const targetSheet = SpreadsheetApp.openById(config.TARGET_SS_ID).getSheetByName(config.TAB_NAME);
if (!targetSheet) {
Logger.log("Fehler: Tab '" + config.TAB_NAME + "' nicht gefunden.");
return;
}

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

// Schreiben: Bestehende Zeile aktualisieren oder neue anfügen
if (targetRowIndex !== -1) {
targetSheet.getRange(targetRowIndex, 1, 1, rowData.length).setValues([rowData]);
} else {
targetSheet.appendRow(rowData);
targetRowIndex = targetSheet.getLastRow();
}

// Zeilenhöhe auf 21px fixieren & Spalten G/H auf CLIP setzen
targetSheet.setRowHeight(targetRowIndex, 21);
targetSheet.getRange(targetRowIndex, 7, 1, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}