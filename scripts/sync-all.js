const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
const generateOverview = require('./generate-overview');
const { extractBackupProjects } = require('./backup-index');
const BACKUP_INDEX_SPREADSHEET_ID = '1rLFAl4pdgADo6j_vaDxmPMd8kAHL0V8Pqp-DjeXBxLo';

async function syncAllScripts() {
  try {
    const saCredentialsJson = process.env.GAPPSCRIPT_SA_CREDENTIALS;
    if (!saCredentialsJson) {
      throw new Error("Secret GAPPSCRIPT_SA_CREDENTIALS fehlt oder ist leer!");
    }

    const credentials = JSON.parse(saCredentialsJson);

    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: [
        'https://www.googleapis.com/auth/spreadsheets.readonly',
        'https://www.googleapis.com/auth/script.projects.readonly'
      ]
    });

    const sheets = google.sheets({ version: 'v4', auth });
    const scriptApi = google.script({ version: 'v1', auth });

    console.log('Lese Apps-Script-Projekte aus dem Backup-Index...');
    const index = await sheets.spreadsheets.get({
      spreadsheetId: BACKUP_INDEX_SPREADSHEET_ID,
      fields: 'sheets.properties.title'
    });
    const ranges = (index.data.sheets || []).map(sheet => {
      const title = sheet.properties.title.replace(/'/g, "''");
      return `'${title}'!A:E`;
    });
    if (ranges.length === 0) {
      throw new Error('Der Backup-Index enthält keine Tabellenblätter.');
    }
    const indexValues = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: BACKUP_INDEX_SPREADSHEET_ID,
      ranges,
      valueRenderOption: 'FORMATTED_VALUE'
    });
    const projects = extractBackupProjects(indexValues.data.valueRanges || []);

    console.log(`${projects.length} Projekte im Index gefunden. Starte Download...`);

    const projectsDir = path.join(__dirname, '../projects');
    if (!fs.existsSync(projectsDir)) {
      fs.mkdirSync(projectsDir, { recursive: true });
    }

    let downloadedCount = 0;

    const expectedFolders = new Set();
    for (const project of projects) {
      try {
        const content = await scriptApi.projects.getContent({ scriptId: project.scriptId });
        const scriptFiles = content.data.files || [];

        if (scriptFiles.length === 0) {
          throw new Error('Google hat keine Projektdateien zurückgegeben.');
        }

        const folderPath = path.join(projectsDir, project.folderName);

        console.log(`Lade herunter: ${project.projectName} (${project.scriptId})...`);

        if (!fs.existsSync(folderPath)) {
          fs.mkdirSync(folderPath, { recursive: true });
        }

        const claspJson = { scriptId: project.scriptId, rootDir: '.' };
        fs.writeFileSync(path.join(folderPath, '.clasp.json'), JSON.stringify(claspJson, null, 2));
        const expectedFiles = new Set(['.clasp.json']);

        for (const sFile of scriptFiles) {
          let ext = '.js';
          if (sFile.type === 'HTML') ext = '.html';
          if (sFile.type === 'JSON') ext = '.json';

          const filename = (sFile.name === 'appsscript' && ext === '.json')
            ? 'appsscript.json' 
            : `${sFile.name}${ext}`;
          if (!filename || path.basename(filename) !== filename) {
            throw new Error(`Ungültiger Dateiname von Google: ${filename}`);
          }

          const filePath = path.join(folderPath, filename);
          fs.writeFileSync(filePath, sFile.source || '');
          expectedFiles.add(filename);
        }

        for (const entry of fs.readdirSync(folderPath, { withFileTypes: true })) {
          if (entry.isFile() && /\.(gs|js|html|json)$/i.test(entry.name) && !expectedFiles.has(entry.name)) {
            fs.unlinkSync(path.join(folderPath, entry.name));
          }
        }

        expectedFolders.add(project.folderName);
        downloadedCount++;
      } catch (err) {
        throw new Error(`Fehler bei ${project.projectName} (${project.scriptId}): ${err.message}`);
      }
    }

    console.log(`Sync abgeschlossen! ${downloadedCount} Apps Script Projekte erfolgreich heruntergeladen.`);
    for (const entry of fs.readdirSync(projectsDir, { withFileTypes: true })) {
      if (entry.isDirectory() && !expectedFolders.has(entry.name)) {
        fs.rmSync(path.join(projectsDir, entry.name), { recursive: true, force: true });
        console.log(`Entferne nicht mehr gelistetes Projekt: ${entry.name}`);
      }
    }

    const overviewPath = path.join(__dirname, '../projects-overview.csv');
    const scriptFileCount = generateOverview(projectsDir, overviewPath);
    console.log(`Projektübersicht erstellt: ${scriptFileCount} Skriptdateien in ${overviewPath}.`);

  } catch (error) {
    console.error("Kritischer Fehler:", error.message || error);
    process.exit(1);
  }
}

syncAllScripts();
