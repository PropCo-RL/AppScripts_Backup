/**
 * Empfängt Aufrufe von Gästen im Browser (GET-Request)
 */
function doGet(e) {
  const rowIndex = e.parameter.row || '';
  const name = e.parameter.name || 'Gast';
  const apartment = e.parameter.apt || '';
  const departureRaw = e.parameter.depDate || '';
  const maxDateRaw = e.parameter.maxDate || '';
  const gmailId = e.parameter.draftId || e.parameter.gmailId || '';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName('output');

  // 1. Prüfung: Link älter als 3 Tage nach Abreisedatum?
  if (departureRaw) {
    let depDateParsed = parseDateString_(departureRaw);
    if (depDateParsed) {
      let expireDate = new Date(depDateParsed.getTime());
      expireDate.setDate(expireDate.getDate() + 3);
      if (new Date() > expireDate) {
        return renderErrorPage_("Link abgelaufen", "Dieser Bestätigungslink ist leider abgelaufen (Gültigkeit max. 3 Tage nach Abreisedatum). Bitte kontaktieren Sie uns direkt per WhatsApp oder E-Mail.");
      }
    }
  }

  // 2. Prüfung: Wurde bereits geantwortet?
  let existingResponse = null;
  if (outputSheet && rowIndex >= 2) {
    const colIndex = findColumnIndexByName_(outputSheet, 'webApp Antwort');
    if (colIndex !== -1) {
      const cellVal = outputSheet.getRange(rowIndex, colIndex).getValue();
      if (cellVal && String(cellVal).trim().startsWith('{')) {
        try {
          existingResponse = JSON.parse(cellVal);
        } catch (err) {}
      }
    }
  }

  // Standard-Datum: Abreisedatum + 7 Tage
  let defaultExtDate = '';
  let depDateParsed = parseDateString_(departureRaw);
  if (depDateParsed) {
    let targetDate = new Date(depDateParsed.getTime());
    targetDate.setDate(targetDate.getDate() + 7);
    defaultExtDate = formatDateToIso_(targetDate);
  } else {
    let todayPlus7 = new Date();
    todayPlus7.setDate(todayPlus7.getDate() + 7);
    defaultExtDate = formatDateToIso_(todayPlus7);
  }

  // Maximales Datum berechnen
  let maxDateAttr = '';
  if (maxDateRaw && !maxDateRaw.includes('Option A')) {
    let parsedMax = parseDateString_(maxDateRaw);
    if (parsedMax) {
      maxDateAttr = formatDateToIso_(parsedMax);
      if (defaultExtDate > maxDateAttr) {
        defaultExtDate = maxDateAttr;
      }
    }
  }

  const maxDateGerman = maxDateAttr ? formatIsoToGerman_(maxDateAttr) : '';
  const existingJson = existingResponse ? JSON.stringify(existingResponse) : 'null';

  // Schnellere Ausführung: HTML direkt als String generieren
  const htmlContent = `
    <!DOCTYPE html>
    <html lang="de">
      <head>
        <meta charset="utf-8">
        <base target="_top">
        <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
        <title>L8 Street – Abreise & Verlängerung</title>
        <style>
          * { box-sizing: border-box; }
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; margin: 0; padding: 12px; background-color: #f0f4f8; color: #1a202c; font-size: 16px; }
          .container { max-width: 580px; margin: 0 auto; background: #ffffff; padding: 22px 18px; border-radius: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
          .text-center { text-align: center; }
          h2 { color: #1a202c; font-size: 22px; margin-top: 0; margin-bottom: 6px; }
          .sub-title { font-size: 15px; color: #4a5568; margin-bottom: 22px; }
          .radio-box { display: flex; align-items: center; background: #f7fafc; padding: 16px; margin-bottom: 12px; border-radius: 10px; border: 2px solid #e2e8f0; cursor: pointer; font-weight: 600; font-size: 17px; }
          .radio-box:hover { border-color: #cbd5e0; }
          input[type="radio"] { margin-right: 14px; width: 22px; height: 22px; accent-color: #3182ce; }
          .extend-fields { display: none; margin-top: 18px; background: #faf5ff; padding: 18px 14px; border-radius: 12px; border: 1px solid #e9d8fd; }
          .field-group { margin-bottom: 16px; }
          .field-group label { display: block; font-size: 14px; font-weight: bold; margin-bottom: 6px; color: #2d3748; }
          input[type="date"], input[type="text"], select { padding: 12px; font-size: 16px; width: 100%; border: 1px solid #cbd5e0; border-radius: 8px; background: #fff; }
          .checkbox-group { margin-top: 20px; font-size: 14px; color: #4a5568; border-top: 1px solid #e2e8f0; padding-top: 16px; }
          .checkbox-group label { display: flex; align-items: flex-start; margin-bottom: 14px; cursor: pointer; line-height: 1.4; }
          .checkbox-group input[type="checkbox"] { margin-right: 12px; margin-top: 2px; width: 22px; height: 22px; accent-color: #3182ce; flex-shrink: 0; }
          button.btn-primary { background: #3182ce; color: white; border: none; padding: 16px 20px; font-size: 18px; border-radius: 10px; width: 100%; cursor: pointer; font-weight: bold; margin-top: 15px; box-shadow: 0 4px 12px rgba(49,130,206,0.25); }
          button.btn-primary:hover { background: #2b6cb0; }
          button.btn-danger { background: #e53e3e; color: white; border: none; padding: 14px 20px; font-size: 16px; border-radius: 8px; width: 100%; cursor: pointer; font-weight: bold; margin-top: 10px; }
          button:disabled { background: #a0aec0; cursor: not-allowed; box-shadow: none; }
          .info-box { background: #ebf8ff; border: 1px solid #bee3f8; padding: 16px; border-radius: 10px; margin: 20px 0; text-align: left; font-size: 15px; color: #2b6cb0; line-height: 1.5; }
          .storno-box { background: #fff5f5; border: 1px solid #fed7d7; padding: 16px; border-radius: 10px; margin-top: 25px; text-align: center; }
          .info-hint { font-size: 12px; color: #e53e3e; font-weight: bold; margin-top: 4px; }
          .icon { font-size: 52px; margin-bottom: 10px; }
          a { color: #3182ce; text-decoration: underline; }
        </style>
      </head>
      <body>
        <div class="container">
          
          <div id="formView">
            <h2>Abreise & Verlängerung</h2>
            <div class="sub-title">Hallo <strong>${name}</strong> (Apartment ${apartment})</div>
            
            <form id="departureForm" onsubmit="submitForm(event)">
              <input type="hidden" name="rowIndex" value="${rowIndex}">
              <input type="hidden" name="gmailId" value="${gmailId}">
              
              <div class="option-group">
                <label class="radio-box">
                  <input type="radio" name="responseType" value="Keine Verlängerung" onclick="toggleExtension(false)" required>
                  Ich reise wie geplant aus
                </label>
                
                <label class="radio-box">
                  <input type="radio" name="responseType" value="Verlängerung" onclick="toggleExtension(true)">
                  Ich möchte verlängern
                </label>
              </div>

              <div id="extensionDetails" class="extend-fields">
                <div class="field-group">
                  <label for="newDate">Verlängerung gewünscht bis (Abreisedatum): *</label>
                  <input type="date" id="newDate" name="newDate" value="${defaultExtDate}" ${maxDateAttr ? `max="${maxDateAttr}"` : ''}>
                  ${maxDateGerman ? `<div class="info-hint">Hinweis: Aufgrund von Folgebuchungen ist eine Verlängerung maximal bis zum ${maxDateGerman} möglich.</div>` : ''}
                </div>

                <div class="field-group">
                  <label for="billingCycle">Gewünschter Abrechnungszyklus:</label>
                  <select id="billingCycle" name="billingCycle">
                    <option value="Gesamter Zeitraum auf einmal">Gesamter Zeitraum auf einmal</option>
                    <option value="Wöchentlich">Wöchentlich</option>
                    <option value="Alle 14 Tage">Alle 14 Tage</option>
                    <option value="Monatlich">Monatlich</option>
                  </select>
                </div>

                <div class="field-group">
                  <label for="invoiceAddressType">Rechnungsanschrift:</label>
                  <select id="invoiceAddressType" name="invoiceAddressType" onchange="toggleAddressField(this.value)">
                    <option value="Unverändert (wie bisher)">Unverändert (wie bisher)</option>
                    <option value="Neue Rechnungsanschrift">Neue Rechnungsanschrift (Hinweis: Die bisherige Firma bleibt weiterhin haftbar)</option>
                  </select>
                </div>

                <div class="field-group" id="customAddressContainer" style="display: none;">
                  <label for="customInvoiceAddress">Neue Rechnungsanschrift eingeben:</label>
                  <input type="text" id="customInvoiceAddress" name="customInvoiceAddress" placeholder="Firma / Name, Straße, PLZ, Ort">
                </div>

                <div class="checkbox-group">
                  <label>
                    <input type="checkbox" id="agbAccepted" name="agbAccepted" value="Ja">
                    <span>Ich akzeptiere die <a href="https://L8Street.com/agb" target="_blank">AGB von L8Street.com</a>. *</span>
                  </label>
                  
                  <label>
                    <input type="checkbox" id="stornoAccepted" name="stornoAccepted" value="Ja">
                    <span><strong>Stornobedingungen:</strong> Bei Irrtum ist eine kostenfreie Stornierung per E-Mail/Button innerhalb von 3 Stunden möglich. Danach ist die Verlängerung verbindlich und nicht mehr stornierbar. *</span>
                  </label>
                </div>
              </div>

              <button type="submit" id="submitBtn" class="btn-primary">Abreise bestätigen</button>
            </form>
          </div>

          <div id="thanksView" class="text-center" style="display: none;">
            <div class="icon" id="thanksIcon" style="color:#38a169;">✓</div>
            <h2 id="thanksTitle">Vielen Dank für Ihre Rückmeldung!</h2>
            <p id="thanksSub">Ihre Angaben wurden erfolgreich verarbeitet.</p>
            
            <div class="info-box" id="infoBox" style="display: none;">
              <strong>So geht es weiter:</strong><br>
              Sie erhalten innerhalb von ca. <strong>40 Minuten</strong> Ihre Buchungsbestätigung bzw. Rechnung per E-Mail.
            </div>

            <div class="storno-box" id="stornoBox" style="display: none;">
              <p style="margin:0; font-size:14px; color:#c53030;">
                <strong>Aus Versehen geklickt?</strong><br>Sie können diese Verlängerung kostenfrei stornieren bis:<br><strong id="stornoDeadlineText"></strong>
              </p>
              <button class="btn-danger" onclick="cancelBooking()">Irrtum? Kostenfrei stornieren</button>
            </div>

            <p style="margin-top: 25px; font-size: 13px; color: #a0aec0;">Ihr L8 Street Team</p>
          </div>

          <div id="cancelledView" class="text-center" style="display: none;">
            <div class="icon" style="color:#e53e3e;">✕</div>
            <h2 style="color:#e53e3e;">Verlängerung storniert</h2>
            <p style="color:#4a5568; line-height:1.6;">Ihre Verlängerungsanfrage wurde wunschgemäß kostenfrei storniert. Es fallen keine Kosten für Sie an.</p>
            <p style="margin-top: 25px; font-size: 13px; color: #a0aec0;">Ihr L8 Street Team</p>
          </div>

        </div>

        <script>
          const existingData = ${existingJson};
          const curRowIndex = ${rowIndex || 0};

          window.onload = function() {
            if (existingData && existingData !== null && typeof existingData === 'object') {
              showStatusView(existingData);
            }
          };

          function showStatusView(data) {
            document.getElementById('formView').style.display = 'none';

            if (data.status === 'STORNIERT_DURCH_GAST') {
              document.getElementById('cancelledView').style.display = 'block';
              return;
            }

            const antwort = data.antwortTyp || data.responseType || 'Eingegangen';

            document.getElementById('thanksView').style.display = 'block';
            document.getElementById('thanksSub').innerHTML = 'Ihre Rückmeldung (<strong>' + antwort + '</strong>) liegt uns vor.';

            if (antwort === 'Verlängerung') {
              document.getElementById('infoBox').style.display = 'block';

              if (data.stornoMoeglichBis) {
                const parts = data.stornoMoeglichBis.split(' ');
                const dParts = parts[0].split('.');
                const tParts = parts[1].split(':');
                const deadlineDate = new Date(dParts[2], dParts[1] - 1, dParts[0], tParts[0], tParts[1], tParts[2]);

                if (new Date() <= deadlineDate) {
                  document.getElementById('stornoDeadlineText').innerText = data.stornoMoeglichBis + ' Uhr';
                  document.getElementById('stornoBox').style.display = 'block';
                }
              }
            }
          }

          function toggleExtension(isExtend) {
            const extDiv = document.getElementById('extensionDetails');
            const newDate = document.getElementById('newDate');
            const agb = document.getElementById('agbAccepted');
            const storno = document.getElementById('stornoAccepted');
            const btn = document.getElementById('submitBtn');

            if (isExtend) {
              extDiv.style.display = 'block';
              newDate.required = true;
              agb.required = true;
              storno.required = true;
              btn.innerText = 'Kostenpflichtig / Verbindlich verlängern';
            } else {
              extDiv.style.display = 'none';
              newDate.required = false;
              agb.required = false;
              storno.required = false;
              btn.innerText = 'Abreise bestätigen';
            }
          }

          function toggleAddressField(val) {
            const container = document.getElementById('customAddressContainer');
            container.style.display = (val.startsWith('Neue Rechnungsanschrift')) ? 'block' : 'none';
          }

          function submitForm(e) {
            e.preventDefault();
            const btn = document.getElementById('submitBtn');
            btn.disabled = true;
            btn.innerText = 'Wird übermittelt...';

            google.script.run
              .withSuccessHandler(function(response) {
                if (response && response.status === 'success') {
                  window.location.reload();
                } else {
                  alert('Fehler: ' + (response ? response.message : 'Unbekannter Fehler'));
                  btn.disabled = false;
                  btn.innerText = 'Erneut versuchen';
                }
              })
              .withFailureHandler(function(err) {
                alert('Serverfehler: ' + err.message);
                btn.disabled = false;
              })
              .processGuestResponse(document.getElementById('departureForm'));
          }

          function cancelBooking() {
            if (confirm('Möchten Sie diese Verlängerungsanfrage wirklich kostenfrei stornieren?')) {
              google.script.run
                .withSuccessHandler(function(res) {
                  if (res.status === 'success') {
                    window.location.reload();
                  } else {
                    alert(res.message);
                  }
                })
                .processGuestStorno(curRowIndex);
            }
          }
        </script>
      </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('L8 Street – Abreisebestätigung')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processGuestResponse(formObject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const outputSheet = ss.getSheetByName('output');
    if (!outputSheet) return { status: 'error', message: 'Tabellenblatt "output" nicht gefunden.' };

    const rowIndex = parseInt(formObject.rowIndex);
    if (isNaN(rowIndex) || rowIndex < 2) return { status: 'error', message: 'Ungültige Zeilenangabe.' };

    const isExtension = formObject.responseType === 'Verlängerung';
    const now = new Date();
    const timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');
    const stornoDeadline = new Date(now.getTime() + (3 * 60 * 60 * 1000));
    const stornoDeadlineStr = Utilities.formatDate(stornoDeadline, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');

    const responseData = {
      timestamp: timestampStr,
      timestampIso: now.toISOString(),
      gmailId: formObject.gmailId || 'N/A',
      antwortTyp: formObject.responseType,
      verlaengerungBis: isExtension ? formObject.newDate : null,
      abrechnungszyklus: isExtension ? formObject.billingCycle : null,
      rechnungsanschriftTyp: isExtension ? formObject.invoiceAddressType : null,
      abweichendeRechnungsadresse: isExtension ? formObject.customInvoiceAddress : null,
      agbZugestimmt: isExtension ? (formObject.agbAccepted === 'Ja') : false,
      stornobedingungenAkzeptiert: isExtension ? (formObject.stornoAccepted === 'Ja') : false,
      stornoMoeglichBis: isExtension ? stornoDeadlineStr : null,
      storniertAm: null,
      status: 'VERBINDLICH_GEBUCHT'
    };

    const colIndex = findColumnIndexByName_(outputSheet, 'webApp Antwort');
    if (colIndex === -1) return { status: 'error', message: 'Spalte "webApp Antwort" nicht gefunden.' };

    outputSheet.getRange(rowIndex, colIndex).setValue(JSON.stringify(responseData));

    return { status: 'success' };

  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function processGuestStorno(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const outputSheet = ss.getSheetByName('output');
    if (!outputSheet) return { status: 'error', message: 'Sheet nicht gefunden' };

    const colIndex = findColumnIndexByName_(outputSheet, 'webApp Antwort');
    if (colIndex === -1) return { status: 'error', message: 'Spalte nicht gefunden' };

    const cellRange = outputSheet.getRange(rowIndex, colIndex);
    const cellVal = cellRange.getValue();

    if (!cellVal) return { status: 'error', message: 'Keine Buchung gefunden' };

    let data = JSON.parse(cellVal);
    const now = new Date();

    data.status = 'STORNIERT_DURCH_GAST';
    data.storniertAm = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');

    cellRange.setValue(JSON.stringify(data));

    return { status: 'success' };

  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function renderErrorPage_(title, message) {
  const html = `<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1.0"><style>body{font-family:sans-serif; padding:30px 15px; text-align:center; background:#f4f6f9;}.card{max-width:450px; margin:0 auto; background:white; padding:25px; border-radius:12px;}</style></head><body><div class="card"><h2 style="color:#e53e3e;">${title}</h2><p style="color:#4a5568; line-height:1.5;">${message}</p></div></body></html>`;
  return HtmlService.createHtmlOutput(html);
}

function findColumnIndexByName_(sheet, columnName) {
  const headers = sheet.getRange(1, 1, 2, sheet.getLastColumn()).getValues();
  for (let row = 0; row < headers.length; row++) {
    for (let col = 0; col < headers[row].length; col++) {
      const cellValue = String(headers[row][col]).trim();
      if (cellValue.toLowerCase() === columnName.toLowerCase()) return col + 1;
    }
  }
  return -1;
}

function parseDateString_(str) {
  if (!str) return null;
  if (typeof str === 'object' && str instanceof Date) return new Date(str.getTime());
  str = String(str).trim();
  if (str.includes('.')) {
    const p = str.split('.');
    if (p.length === 3) return new Date(p[2], p[1] - 1, p[0]);
  }
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateToIso_(dateObj) {
  const year = dateObj.getFullYear();
  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
  const day = String(dateObj.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatIsoToGerman_(isoStr) {
  const p = isoStr.split('-');
  return `${p[2]}.${p[1]}.${p[0]}`;
}