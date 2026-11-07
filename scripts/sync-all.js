const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
const generateOverview = require('./generate-overview');

// ==========================================
// 1. EINGEBETTETE GOOGLE SHEET SKRIPTE
// ==========================================
const SHEET_SCRIPT_IDS = [
  "15ET_O-o0uHT5dSVcseb28ShbwPFSPJA0bbidmopCyFfqgX6qmYOzfjQk",
  "1LMpO3_vD2BGzx4kvkfD--p0RKhJBmLMtZrFfD1S8V7d_SFNswV2JqeLN",
  "1Kl6tQ1siwaIjTLeOqEUsWFPNdKUS0wULctWC_nLHrmtlrt1jmFgJg-AJ",
  "1SFxPHAOScvy62ON5LvJwPofrKprvu8J1PwGDzDk8zNLLu-noHwSDUZ38",
  "1C0e89p_ZlMvy8OGbpj7WbLFvEW79T1WCcen_usjq1NbSEGa1g9b6EThY",
  "18M_DPau_8kcheWA83HYCRZBGfYbHfIgTM7oiYhtox95I9mwCzDo8iuHF",
  "1W-r2hbSYs1g8ZXFjI1-eWqx4AG9U7ttBrFhD16_Mwsps60-98HLbfJ1N",
  "1JNfOWh0Ai7cpGB3g4TWYQcg4wiK88NFjK9IMgajD22m9ODk3_w-n7Jus",
  "1DtJ5J8-hNqzmME7e8WBq3zME8Oj_p8gQvTwyzoMTJBLUBV3nuhwPG17t",
  "1kbpC7OBDlEQ1kp30-zDB6dVZwAD0J1_l1M-n6FtEwHOF4k2ayeIT4qhd",
  "1KiwCB5ETXqrLJD-oOmsB71t9r_gLgq1ADJdkukCWyzO78Y9v6m8kyKsn",
  "18K6oe1xL6kKzagWCmYx2T9ZxYnigcPnLNPlT_AvLkzFxxSWwfOBUukaf",
  "1av8e-T1nAAv1NOgGN_d3bVy3k1uzmeBtDu-oJ5gSejGWYOot0tquTnqk",
  "1WIM6bRqwBXQp0vYi1jMMSdIcOSHk3d9-9nVXWn07dZsR92HiUaPs3sym",
  "18xAQ7M669a8PAD9GMFSxUZnnghn7ThWGymb-cIget4S52dvzRUc9IzjE",
  "1QXnG2kZc1YWoGnCk7kRTULYlTEUjlcuFxl1QzHgBS4r7-DUMfjkETKPq",
  "1e3L-vQPfm09-jlfQX4e74xcEUcL-FxlVneM4lfx4Ns2ATXzJVnKE8pfp",
  "1KbFzROEK8Devsv5ClWEk0Xa56R8B5Wb80OLE2cczhE5zZKEerbRQSnzb",
  "1XTHzi5USyP0glUWfr18WgZOkI2m0LW4fK5LxYxK9C0cSPZpU9Iuvgles",
  "11OKBinxNFi1mEz4q9_pKDZYy8_QK5fZI4q3UCUYcgy0s39BtnppcJ_ZM",
  "1Yf-BgGRh6_HpyLlni-K8_HN6KWXp_qg4pIWtrhsdke31PwhDQvGk2SK",
  "1o98ykJvnSRIR4r1eMGWW-6BCB1DYkNf-0UX-EliGjoCil122X3K61xF-1XuNrxCj-cKzOXVcPxZe2gIi0tLP19KOqRZGBdmxcCqKAgRvjqXTkTT1z"
];

async function syncAllScripts() {
  try {
    const saCredentialsJson = process.env.GAPPSCRIPT_SA_CREDENTIALS;
    if (!saCredentialsJson) {
      throw new Error("Secret GAPPSCRIPT_SA_CREDENTIALS fehlt oder ist leer!");
    }

    const credentials = JSON.parse(saCredentialsJson);

    // Strenger READ-ONLY Zugriff!
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/script.projects.readonly'
      ]
    });

    const drive = google.drive({ version: 'v3', auth });
    const scriptApi = google.script({ version: 'v1', auth });

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
    console.log(`${standaloneScripts.length} freistehende Skripte automatisch gefunden.`);

    // Kombinieren: Automatische Skripte + Ihre manuellen Sheet-Skripte
    const allScripts = [...standaloneScripts];
    for (const id of SHEET_SCRIPT_IDS) {
      if (id) {
        allScripts.push({ id: id, name: null }); 
      }
    }

    console.log(`Insgesamt zu verarbeiten: ${allScripts.length} Projekte. Starte Download...`);

    const projectsDir = path.join(__dirname, '../projects');
    if (!fs.existsSync(projectsDir)) {
      fs.mkdirSync(projectsDir, { recursive: true });
    }

    let downloadedCount = 0;

    for (const file of allScripts) {
      try {
        let projectName = file.name;
        
        // Namen direkt über die API abrufen, falls er aus der manuellen ID-Liste kommt
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
      }
    }

    console.log(`Sync abgeschlossen! ${downloadedCount} Apps Script Projekte erfolgreich heruntergeladen.`);
    const overviewPath = path.join(__dirname, '../projects-overview.csv');
    const scriptFileCount = generateOverview(projectsDir, overviewPath);
    console.log(`Projektübersicht erstellt: ${scriptFileCount} Skriptdateien in ${overviewPath}.`);

  } catch (error) {
    console.error("Kritischer Fehler:", error.message || error);
    process.exit(1);
  }
}

syncAllScripts();
