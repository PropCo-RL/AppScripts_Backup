// ============================================================================
// SEPARATES ERGÄNZUNGSSCRIPT: PDF-SPEICHERUNG IN GOOGLE DRIVE & LINK-GENERIERUNG
// ============================================================================

const TARGET_FOLDER_ID = "13CMmLQzGkfh46SSGHoW7qh79ItDaOZLf";

/**
 * Erstellt und speichert die WGB-PDFs für den Tab 'WGB' in Google Drive.
 * Liest den Status aus Spalte G und schreibt den Link in Spalte H (8).
 */
function syncWGBToDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWGB = ss.getSheetByName("wgb") || ss.getSheetByName("WGB");
  const sheetD = ss.getSheetByName("d");

  if (!sheetWGB || !sheetD) return;
  const lastRowWGB = sheetWGB.getLastRow();
  if (lastRowWGB < 2) return;

  const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  const DEFAULT_LANDLORD = "L8 Street GmbH, Forststraße 65, 75223 Niefern-Öschelbronn";

  // Apartment Mappings aus Tab 'd' laden
  const aptMap = {};
  const lastRowD = sheetD.getLastRow();
  if (lastRowD >= 2) {
    sheetD.getRange(2, 1, lastRowD - 1, 14).getValues().forEach(row => {
      const match = row[3] ? row[3].toString().match(/\/a\/([^\/\?]+)/i) : null;
      if (!match) return;

      const code = match[1].toLowerCase().trim();
      let ident = [row[11], row[12]].filter(x => x && x.toString().trim()).join(", ");

      aptMap[code] = {
        landlord: (row[5] && row[5].toString().trim()) ? row[5].toString().trim() : DEFAULT_LANDLORD,
        street: (row[7] + " " + row[8]).trim(),
        cityZip: (row[9] + " " + row[10]).trim(),
        ident: ident,
        owner: row[13] ? row[13].toString().trim() : ""
      };
    });
  }

  const values = sheetWGB.getRange(2, 1, lastRowWGB - 1, 8).getValues();

  values.forEach((row, i) => {
    const rowIndex = i + 2;
    const aptCode = row[1] ? row[1].toString().toLowerCase().trim() : "";
    const guestNames = row[2] ? row[2].toString().trim() : "";
    const moveInRaw = row[3];
    const guestEmail = row[5] ? row[5].toString().trim() : "";
    const status = row[6] ? row[6].toString().trim() : "";
    const currentLink = row[7] ? row[7].toString().trim() : ""; // Spalte H (8)

    // Abbrechen, wenn noch nicht versendet ODER wenn bereits ein Drive-Link in Spalte H steht
    if (!status.includes("versendet") || currentLink.startsWith("http")) return;

    const apt = aptMap[aptCode];
    if (!apt) return;

    const moveInDate = (moveInRaw instanceof Date)
      ? Utilities.formatDate(moveInRaw, "Europe/Berlin", "dd.MM.yyyy")
      : moveInRaw.toString();
    const today = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy");

    try {
      const html = `
      <!DOCTYPE html><html><head><meta charset="utf-8">
      <link href="https://fonts.googleapis.com/css2?family=Dancing+Script:wght@700&display=swap" rel="stylesheet">
      <style>
      body { font-family: Arial, sans-serif; font-size: 12pt; line-height: 1.5; margin: 30px; color: #111; }
      h1 { font-size: 16pt; text-align: center; margin-bottom: 5px; text-transform: uppercase; }
      .sub { text-align: center; font-size: 10pt; font-weight: bold; margin-bottom: 30px; color: #444; }
      .sec { margin-bottom: 22px; }
      .title { font-weight: bold; font-size: 11pt; border-bottom: 1px solid #000; padding-bottom: 3px; margin-bottom: 8px; text-transform: uppercase; }
      .box { background: #f9f9f9; padding: 10px; border-radius: 4px; }
      </style></head><body>
      <h1>Wohnungsgeberbestätigung</h1>
      <div class="sub">nach § 19 Abs. 3 des Bundesmeldegesetzes (BMG)</div>
      <p>Hiermit wird der <strong>Einzug</strong> der nachstehend genannten Person(en) bestätigt:</p>

      <div class="sec">
      <div class="title">1. Angaben zum Wohnungsgeber / Vermieter</div>
      <div class="box">${apt.landlord.replace(/\n/g, '<br>')}</div>
      </div>

      <div class="sec">
      <div class="title">2. Anschrift & Identifikation der Wohnung</div>
      <div><strong>Straße / Hausnr.:</strong> ${apt.street}</div>
      <div><strong>PLZ / Ort:</strong> ${apt.cityZip}</div>
      ${apt.ident ? `<div><strong>Lage/Identifikation:</strong> ${apt.ident}</div>` : ''}
      ${apt.owner ? `<div style="margin-top:5px;"><strong>Eigentümer:</strong> ${apt.owner}</div>` : ''}
      </div>

      <div class="sec">
      <div class="title">3. Angaben zu den einziehenden Personen</div>
      <div>Folgende Person(en) sind am <strong>${moveInDate}</strong> eingezogen:</div><br>
      <div style="font-weight:bold; font-size:12pt;">${guestNames.replace(/,/g, '<br>')}</div>
      </div>

      <div style="margin-top:40px;">
      <div><strong>Ausstellungsdatum:</strong> ${today}</div>
      <div style="width:300px; margin-top:25px;">
      <div style="font-family:'Dancing Script', cursive; font-size:24pt; font-weight:700; color:#002b66; transform:rotate(-4deg); margin-left:10px;">R. Leneweit</div>
      <div style="border-top:1px solid #000; padding-top:5px; font-size:10pt;">
      <strong>Raul Leneweit</strong><br>L8 Street GmbH
      </div>
      </div>
      </div>
      </body></html>`;

      const pdfBlob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
      const fileName = `WGB_${aptCode.toUpperCase()}_${guestNames.replace(/[^a-zA-Z0-9]/g, "_")}.pdf`;
      pdfBlob.setName(fileName);

      const file = folder.createFile(pdfBlob);
      const fileId = file.getId();
      const directDownloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

      sheetWGB.getRange(rowIndex, 8).setValue(directDownloadUrl); // Spalte H (8)
      Logger.log(`[WGB] PDF gesichert für Zeile ${rowIndex}: ${directDownloadUrl}`);

    } catch (err) {
      Logger.log(`[WGB] Fehler Zeile ${rowIndex}: ${err.message}`);
    }
  });
}