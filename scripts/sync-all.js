const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// ==========================================
// KONFIGURATION: MASTER ÜBERSICHTS-SHEET
// ==========================================
const CONFIG = {
  MASTER_SHEET_ID: "1rLFAl4pdgADo6j_vaDxmPMd8kAHL0V8Pqp-DjeXBxLo",
  RANGE: "Übersicht!E2:E" // Liest Spalte E ab Zeile 2 dynamisch aus
};

// ==========================================
// HILFSFUNKTION: SCRIPT-IDS AUS GOOGLE SHEET LESEN
// ==========================================
async function fetchDynamicScriptIds(auth) {
  try {
    const sheets = google.sheets({ version: 'v4', auth });
    console.log(`Lese Script-IDs aus Master-Sheet (${CONFIG.MASTER_SHEET_ID})...`);

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.MASTER_SHEET_ID,
      range: CONFIG.RANGE,
    });

    const rows = res.data.values || [];
    
    // Leere Zellen filtern, Leerzeichen entfernen und IDs validieren
    const validIds = rows
      .map(r => (r[0] ? String(r[0]).trim() : ""))
      .filter(id => id.length > 20 && !id.includes(" "));

    const uniqueIds = [...new Set(validIds)];
    console.log(`✓ ${uniqueIds.length} Script-IDs dynamisch aus dem Sheet geladen.`);
    return uniqueIds;
  } catch (error) {
    console.error("Fehler beim Abrufen der IDs aus dem Google Sheet:", error.message);
    throw error;
  }
}

// ==========================================
// HAUPTFUNKTION: SYNC ALLER SKRIPTE
// ==========================================
async function syncAllScripts() {
  try {
    const saCredentialsJson = process.env.GAPPSCRIPT_SA_CREDENTIALS;
    if (!saCredentialsJson) {
      throw new Error("Secret GAPPSCRIPT_SA_CREDENTIALS fehlt oder ist leer!");
    }

    const credentials = JSON.parse(saCredentialsJson);

    // Strenger READ-ONLY Zugriff (Drive, Apps Script AND Sheets)
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/script.projects.readonly',
        'https://www.googleapis.com/auth/spreadsheets.readonly'
      ]
    });

    const drive = google.drive({ version: 'v3', auth });
    const scriptApi = google.script({ version: 'v1', auth });

    // 1. Freistehende Standalone-Skripte via Drive API ermitteln
    console.log("Suche freistehende Skripte in Google Drive...");
    const driveRes = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.script' and trashed=false",
      fields: 'files(id, name)',
      pageSize: 1000,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      corpora: 'allDrives'
    });

    const standaloneScripts = driveRes.data.files || [];
    console.log(`${standaloneScripts.length} freistehende Skripte gefunden.`);

    // 2. Eingebettete Skript-IDs live aus dem Sheet holen
    const dynamicSheetIds = await fetchDynamicScriptIds(auth);

    // 3. Zusammenführen ohne Duplikate
    const processedIds = new Set();
    const allScripts = [];

    for (const file of standaloneScripts) {
      processedIds.add(file.id);
      allScripts.push({ id: file.id, name: file.name });
    }

    for (const id of dynamicSheetIds) {
      if (!processedIds.has(id)) {
        processedIds.add(id);
        allScripts.push({ id: id, name: null });
      }
    }

    console.log(`Insgesamt zu verarbeiten: ${allScripts.length} Projekte. Starte Download...`);

    const projectsDir = path.join(__dirname, '../projects');
    if (!fs.existsSync(projectsDir)) {
      fs.mkdirSync(projectsDir, { recursive: true });
    }

    let downloadedCount = 0;
    let failedCount = 0;

    // 4. Download-Schleife
    for (const file of allScripts) {
      try {
        let projectName = file.name;

        if (!projectName) {
          const projectInfo = await scriptApi.projects.get({ scriptId: file.id });
          projectName = projectInfo.data.title || `Unnamed_Script_${file.id}`;
        }

        const content = await scriptApi.projects.getContent({ scriptId: file.id });
        const scriptFiles = content.data.files || [];

        if (scriptFiles.length === 0) continue;

        const folderName = projectName.replace(/[^a-zA-Z0-9_-]/g, '_');
        const folderPath = path.join(projectsDir, folderName);

        console.log(`Lade herunter: ${projectName} (${file.id})...`);

        if (!fs.existsSync(folderPath)) {
          fs.mkdirSync(folderPath, { recursive: true });
        }

        const claspJson = { scriptId: file.id, rootDir: "." };
        fs.writeFileSync(path.join(folderPath, '.clasp.json'), JSON.stringify(claspJson, null, 2));

        for (const sFile of scriptFiles) {
          let ext = '.js';
          if (sFile.type === 'HTML') ext = '.html';
          if (sFile.type === 'JSON') ext = '.json';

          const filename = (sFile.name === 'appsscript' && ext === '.json')
            ? 'appsscript.json'
            : `${sFile.name}${ext}`;

          const filePath = path.join(folderPath, filename);
          fs.writeFileSync(filePath, sFile.source || '');
        }

        downloadedCount++;
      } catch (err) {
        console.error(`Fehler bei Skript ID ${file.id}: Zugriff verweigert oder ID ungültig.`);
        failedCount++;
      }
    }

    console.log(`\nSync abgeschlossen! ${downloadedCount} Apps Script Projekte erfolgreich heruntergeladen.`);

    if (failedCount > 0) {
      console.warn(`Achtung: ${failedCount} Skripte konnten nicht gesichert werden.`);
      process.exit(1);
    }

  } catch (error) {
    console.error("Kritischer Fehler:", error.message || error);
    process.exit(1);
  }
}

syncAllScripts();
