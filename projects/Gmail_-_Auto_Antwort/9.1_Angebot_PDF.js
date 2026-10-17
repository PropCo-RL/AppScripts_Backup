/**
 * Erstellt ein professionelles 2-seitiges PDF-Angebot für Monteurwohnungen.
 * Größere Schriften, 12 große Fotos (3x4 Raster) und feste Fußzeilen am Seitenende.
 * Liest Spalte AC aus und schreibt das Ergebnis in Spalte AI.
 */
function createPdfOffersFromColAC() {
  const sheet1Id = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const sheet2Id = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';
  const targetFolderId = '15Cnz16cF33J6aPkpXSuIgacY8Q200Ep1';

  const s1 = SpreadsheetApp.openById(sheet1Id).getSheetByName('DMZ');
  const s2_m = SpreadsheetApp.openById(sheet2Id).getSheetByName('M');
  const targetFolder = DriveApp.getFolderById(targetFolderId);

  const dmzData = s1.getDataRange().getValues();
  const mData = s2_m.getDataRange().getValues();
  const apartmentsMap = parseApartmentsForPdf(mData);

  // Array für Ergebnisse in Spalte AI (Index 34 / 35. Spalte)
  let outputColAI = [["PDF Angebot Link"]];
  const ONE_DAY_MS = 1000 * 60 * 60 * 24;

  for (let i = 1; i < dmzData.length; i++) {
    let colA = String(dmzData[i][0] || "").trim();
    if (colA === "") break;

    let sourceAC = String(dmzData[i][28] || "").trim();   // Spalte AC (Index 28)
    let existingAI = String(dmzData[i][34] || "").trim(); // Spalte AI (Index 34)

    // WENN AC leer ist OR in AI bereits etwas steht -> Überspringen
    if (sourceAC === "" || existingAI !== "") {
      outputColAI.push([dmzData[i][34] || ""]);
      continue;
    }

    let firstSlugKey = extractFirstSlugKey(sourceAC);
    let apt = apartmentsMap[firstSlugKey];
    if (!apt) {
      outputColAI.push(["Fehler: Objekt in Tab M nicht gefunden"]);
      continue;
    }

    let jsonRaw = dmzData[i][10];
    let jsonMeta = parseJsonMeta(jsonRaw);
    let dateStr = dmzData[i][14] || dmzData[i][4];
    let dates = parseGermanDateRange(dateStr);
    if (!dates) {
      outputColAI.push(["Fehler: Datum ungültig"]);
      continue;
    }

    let nightCount = Math.round((dates.end - dates.start) / ONE_DAY_MS);
    if (nightCount <= 0) nightCount = 1;

    let netPricePerNight = apt.netPricePerNight;
    let totalRentalNet = nightCount * netPricePerNight;
    
    // Fallback-Logik für Endreinigung
    let cleaningFeeNet = (apt.cleaningFeeNet && apt.cleaningFeeNet < 500) ? apt.cleaningFeeNet : 89.00;
    let subtotalNet = totalRentalNet + cleaningFeeNet;
    let vatAmount = subtotalNet * 0.07;
    let totalBrutto = subtotalNet + vatAmount;

    let nowMs = new Date().getTime();
    let validUntilMs = nowMs + (3 * ONE_DAY_MS);

    let displayAddress = apt.address.replace(/\s+/g, ' ').replace(/\s,/g, ',');

    let htmlContent = generateOfferHtml({
      guestName: jsonMeta.gastName || jsonMeta.firma || "Sehr geehrte Damen und Herren",
      customerNo: jsonMeta.kundenNummer || "10003",
      startDateStr: formatDateGerman(dates.start),
      endDateStr: formatDateGerman(dates.end),
      nights: nightCount,
      persons: dmzData[i][13] || dmzData[i][5] || jsonMeta.personen || 1,
      apt: { ...apt, address: displayAddress },
      netPricePerNight: netPricePerNight,
      totalRentalNet: totalRentalNet,
      cleaningFeeNet: cleaningFeeNet,
      subtotalNet: subtotalNet,
      vatAmount: vatAmount,
      totalBrutto: totalBrutto,
      offerNumber: "ANG-" + (20260000 + i),
      currentDateStr: formatDateGerman(nowMs),
      validUntilDateStr: formatDateGerman(validUntilMs)
    });

    let cleanAddressName = displayAddress.replace(/[^a-zA-Z0-9_\-]/g, '_').replace(/_+/g, '_');
    let fileName = `Angebot_${cleanAddressName}_Zeile_${i+1}.pdf`;

    let pdfBlob = Utilities.newBlob(htmlContent, 'text/html', `temp_${i}.html`).getAs('application/pdf').setName(fileName);
    let file = targetFolder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    let pdfUrl = file.getUrl();
    outputColAI.push([pdfUrl]);
  }

  // Rückschreiben in Spalte 35 (Spalte AI)
  if (outputColAI.length > 0) {
    s1.getRange(1, 35, outputColAI.length, 1).setValues(outputColAI);
  }

  Logger.log("--- PDF ERSTELLUNG ABGESCHLOSSEN ---");
}

// ======================================================================
// HTML-TEMPLATE
// ======================================================================

function generateOfferHtml(data) {
  let photosHtml = "";
  if (data.apt.photos && data.apt.photos.length > 0) {
    let topPhotos = data.apt.photos.slice(0, 12);
    let imgTags = topPhotos.map(url => `<div class="photo-card"><img src="${url}" class="photo-img" /></div>`).join("");
    photosHtml = `<div class="photos-grid">${imgTags}</div>`;
  } else {
    photosHtml = `<div style="padding: 20px; color: #777; font-style: italic; font-size: 11pt;">Fotos der Unterkunft stehen online zur Verfügung.</div>`;
  }

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    @page { size: A4; margin: 12mm 15mm 12mm 15mm; }
    body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #222; font-size: 11pt; line-height: 1.5; margin: 0; padding: 0; }
    .page-container { width: 100%; height: 265mm; position: relative; box-sizing: border-box; }
    .page-break { page-break-before: always; }
    .header-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    .brand-logo-text { font-size: 24pt; font-weight: 900; letter-spacing: -0.5px; color: #1a2530; margin: 0; }
    .brand-subtext { font-size: 9pt; color: #666; text-transform: uppercase; letter-spacing: 1px; margin-top: 2px; }
    .company-contact { font-size: 10pt; color: #444; text-align: right; line-height: 1.4; }
    .sender-line { font-size: 8.5pt; color: #777; border-bottom: 1px solid #ddd; padding-bottom: 4px; margin-bottom: 15px; }
    .info-container { width: 100%; border-collapse: collapse; margin-bottom: 25px; }
    .recipient-cell { vertical-align: top; width: 55%; font-size: 11.5pt; line-height: 1.4; color: #111; }
    .meta-cell { vertical-align: top; width: 45%; }
    .meta-table { width: 100%; border-collapse: collapse; font-size: 10pt; }
    .meta-table td { padding: 3px 0; vertical-align: top; }
    .meta-label { color: #666; text-align: right; padding-right: 12px; }
    .meta-value { font-weight: bold; color: #111; text-align: right; white-space: nowrap; }
    .doc-title { font-size: 16pt; font-weight: bold; color: #1a2530; margin-bottom: 15px; border-bottom: 2px solid #1a2530; padding-bottom: 5px; }
    .items-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    .items-table th { background-color: #f5f7fa; border-bottom: 1px solid #1a2530; border-top: 1px solid #ccc; padding: 8px 5px; font-size: 10pt; text-align: left; font-weight: bold; color: #1a2530; }
    .items-table td { padding: 10px 5px; vertical-align: top; font-size: 10.5pt; border-bottom: 1px dotted #e0e0e0; }
    .item-details { font-size: 9.5pt; color: #555; margin-top: 5px; line-height: 1.4; }
    .summary-wrapper { width: 100%; margin-bottom: 20px; }
    .summary-table { width: 50%; float: right; border-collapse: collapse; font-size: 10.5pt; }
    .summary-table td { padding: 4px 0; text-align: right; }
    .summary-label { padding-right: 12px; color: #555; }
    .total-row { font-weight: bold; font-size: 12pt; color: #1a2530; border-top: 1px solid #1a2530; border-bottom: 2px solid #1a2530; }
    .clear { clear: both; }
    .cta-box { background-color: #f0f4f8; border-left: 4px solid #1a2530; padding: 12px 15px; margin: 20px 0; border-radius: 2px; }
    .cta-title { font-weight: bold; font-size: 11pt; color: #1a2530; margin-bottom: 4px; }
    .cta-link { color: #0056b3; font-weight: bold; text-decoration: underline; font-size: 10.5pt; }
    .notes-block { font-size: 9.5pt; color: #444; margin-top: 15px; line-height: 1.5; border-top: 1px solid #eee; padding-top: 10px; }
    .badge-container { margin: 12px 0; }
    .badge { display: inline-block; background-color: #e9ecef; color: #2b2b2b; padding: 5px 9px; font-size: 9pt; font-weight: bold; border-radius: 3px; margin-right: 6px; margin-bottom: 6px; }
    .photos-grid { width: 100%; text-align: left; margin-top: 15px; font-size: 0; }
    .photo-card { width: 32%; height: 160px; display: inline-block; margin: 0 1.3% 15px 0; vertical-align: top; box-sizing: border-box; font-size: 11pt; }
    .photo-card:nth-child(3n) { margin-right: 0; }
    .photo-img { width: 100%; height: 100%; object-fit: cover; border-radius: 4px; border: 1px solid #ccc; }
    .footer-wrapper { position: absolute; bottom: 0; left: 0; width: 100%; }
    .footer-table { width: 100%; border-collapse: collapse; border-top: 1px solid #ccc; padding-top: 8px; font-size: 8pt; color: #666; line-height: 1.4; }
    .footer-col { width: 33%; vertical-align: top; }
  </style>
</head>
<body>
  <!-- SEITE 1 -->
  <div class="page-container">
    <table class="header-table">
      <tr>
        <td style="vertical-align: top;">
          <div class="brand-logo-text">L8 STREET</div>
          <div class="brand-subtext">Serviced Apartments & Worker Housing</div>
        </td>
        <td class="company-contact">
          <strong>L8 Street GmbH</strong><br>
          Postfach CAYA595616<br>
          96035 Bamberg<br>
          Tel.: +49 176 84801295<br>
          support@L8Street.com<br>
          www.L8Street.com
        </td>
      </tr>
    </table>

    <div class="sender-line">L8 Street GmbH • Postfach CAYA595616 • 96035 Bamberg</div>

    <table class="info-container">
      <tr>
        <td class="recipient-cell">
          <strong>${data.guestName}</strong>
        </td>
        <td class="meta-cell">
          <table class="meta-table">
            <tr><td class="meta-label">Angebotsnr.:</td><td class="meta-value">${data.offerNumber}</td></tr>
            <tr><td class="meta-label">Kundennr.:</td><td class="meta-value">${data.customerNo}</td></tr>
            <tr><td class="meta-label">Datum:</td><td class="meta-value">${data.currentDateStr}</td></tr>
            <tr><td class="meta-label">Gültig bis:</td><td class="meta-value" style="color: #c00;">${data.validUntilDateStr}</td></tr>
            <tr><td class="meta-label">Leistungszeitraum:</td><td class="meta-value">${data.startDateStr} bis ${data.endDateStr}</td></tr>
          </table>
        </td>
      </tr>
    </table>

    <div class="doc-title">Unterkunftsangebot ${data.offerNumber}</div>

    <table class="items-table">
      <thead>
        <tr>
          <th width="5%">Pos.</th>
          <th width="47%">Bezeichnung</th>
          <th width="16%">Menge & Einheit</th>
          <th style="text-align:right;" width="10%">Einzel €</th>
          <th style="text-align:right;" width="8%">USt. %</th>
          <th style="text-align:right;" width="14%">Gesamt €</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>1</td>
          <td>
            <strong>Monteurwohnung: ${data.apt.address}</strong>
            <div class="item-details">
              <strong>Belegung:</strong> ${data.persons} Person(en) | <strong>Zeitraum:</strong> ${data.startDateStr} bis ${data.endDateStr}<br>
              <strong>Kapazität:</strong> ${data.apt.bedrooms} Schlafzimmer | ${data.apt.beds} Betten | ${data.apt.bathrooms} Bad
            </div>
          </td>
          <td>${data.nights} Nacht/Nächte</td>
          <td style="text-align:right;">${data.netPricePerNight.toFixed(2)}</td>
          <td style="text-align:right;">7,00</td>
          <td style="text-align:right;">${data.totalRentalNet.toFixed(2)}</td>
        </tr>
        <tr>
          <td>2</td>
          <td>
            <strong>Endreinigung & Herrichtung</strong>
            <div class="item-details">Einmalige Pauschale für Endreinigung und Desinfektion</div>
          </td>
          <td>1 Pauschale</td>
          <td style="text-align:right;">${data.cleaningFeeNet.toFixed(2)}</td>
          <td style="text-align:right;">7,00</td>
          <td style="text-align:right;">${data.cleaningFeeNet.toFixed(2)}</td>
        </tr>
      </tbody>
    </table>

    <div class="summary-wrapper">
      <table class="summary-table">
        <tr>
          <td class="summary-label">Zwischensumme (netto)</td>
          <td>${data.subtotalNet.toFixed(2)} €</td>
        </tr>
        <tr>
          <td class="summary-label">Umsatzsteuer 7%</td>
          <td>${data.vatAmount.toFixed(2)} €</td>
        </tr>
        <tr class="total-row">
          <td class="summary-label">Gesamtbetrag (brutto)</td>
          <td>${data.totalBrutto.toFixed(2)} €</td>
        </tr>
      </table>
      <div class="clear"></div>
    </div>

    <div class="cta-box">
      <div class="cta-title">Möchten Sie dieses Angebot sofort verbindlich annehmen?</div>
      Buchen Sie direkt online über unseren Buchungslink: <br>
      <a class="cta-link" href="${data.apt.bookingUrl.replace(/18street/g, 'l8street')}">${data.apt.bookingUrl.replace(/18street/g, 'l8street')}</a>
    </div>

    <div class="notes-block">
      <strong>Stornierung:</strong> Eine kostenfreie Stornierung ist bis 7 Tage vor der Anreise möglich.<br>
      - <strong>Reservierung:</strong> Dieses Angebot ist freibleibend bis zur verbindlichen Buchung.<br>
      - <strong>AGB & Datenschutz:</strong> Es gelten unsere AGB unter <a href="https://L8Street.com/agb">L8Street.com/agb</a> sowie die Datenschutzerklärung unter <a href="https://L8Street.com/datenschutz">L8Street.com/datenschutz</a>.<br><br>
      Vielen Dank für Ihre Anfrage! Bei Fragen stehen wir Ihnen gerne zur Verfügung.
    </div>

    <div class="footer-wrapper">
      <table class="footer-table">
        <tr>
          <td class="footer-col">
            <strong>L8 Street GmbH</strong><br>
            Postfach CAYA595616<br>
            96035 Bamberg<br>
            Tel.: +49 176 84801295<br>
            support@L8Street.com
          </td>
          <td class="footer-col">
            USt-IdNr.: DE314603427<br>
            Steuernummer: 48051/24783<br>
            Amtsgericht Mannheim: HRB 728552<br>
            Sitz: 75223 Niefern-Öschelbronn<br>
            GF: Raul Leneweit
          </td>
          <td class="footer-col">
            <strong>Bankverbindung:</strong><br>
            Sparkasse Pforzheim Calw<br>
            IBAN: DE41 6665 0085 0008 9793 32<br>
            BIC: PZHSDE66XXX
          </td>
        </tr>
      </table>
    </div>
  </div>

  <!-- SEITE 2 -->
  <div class="page-break"></div>
  <div class="page-container" style="padding-top: 15px;">
    <table class="header-table" style="margin-bottom: 10px;">
      <tr>
        <td><div class="brand-logo-text" style="font-size: 18pt;">L8 STREET</div></td>
        <td style="text-align: right; font-size: 10pt; color: #555;">Objektübersicht zu Angebot <strong>${data.offerNumber}</strong></td>
      </tr>
    </table>

    <div class="doc-title" style="font-size: 14pt; margin-bottom: 10px;">Unterkunft: ${data.apt.address}</div>

    <div class="badge-container">
      <span class="badge">✓ Highspeed WLAN</span>
      <span class="badge">✓ Einzelbetten</span>
      <span class="badge">✓ Voll ausgestattete Küche</span>
      <span class="badge">✓ Waschmaschine</span>
      <span class="badge">✓ Eigenes Bad</span>
      <span class="badge">✓ Kostenfreie Parkmöglichkeiten</span>
    </div>

    <div style="font-size: 10.5pt; color: #444; margin-bottom: 15px; margin-top: 5px;">
      Unsere Unterkünfte sind speziell auf die Bedürfnisse von Firmen und Montage-Teams ausgelegt. Alle Wohnungen sind komplett möbliert und einsatzbereit.
    </div>

    <div style="font-weight: bold; font-size: 11pt; color: #1a2530; margin-bottom: 5px; border-bottom: 1px solid #ddd; padding-bottom: 5px;">
      Impressionen der Unterkunft
    </div>

    ${photosHtml}

    <div class="footer-wrapper">
      <table class="footer-table">
        <tr>
          <td class="footer-col"><strong>L8 Street GmbH</strong> | support@L8Street.com</td>
          <td class="footer-col" style="text-align: center;">www.L8Street.com</td>
          <td class="footer-col" style="text-align: right;">Seite 2 von 2</td>
        </tr>
      </table>
    </div>
  </div>
</body>
</html>`;
}

// ======================================================================
// PARSER & HILFSFUNKTIONEN
// ======================================================================

function cleanSlugKey(rawStr) {
  if (!rawStr) return "";
  let str = String(rawStr).trim().toLowerCase();
  str = str.replace(/18street/g, "l8street");
  str = str.replace(/https?:\/\/(a\.)?l8street\.com/g, "");
  str = str.replace(/^\/+|\/+$/g, "");
  if (str.startsWith("a/")) {
    str = str.substring(2);
  }
  return str.trim();
}

function parseApartmentsForPdf(mData) {
  let map = {};
  for (let i = 1; i < mData.length; i++) {
    let rawSlug = String(mData[i][6] || "").trim();
    if (!rawSlug) continue;

    let key = cleanSlugKey(rawSlug);
    let rawPhotos = String(mData[i][21] || "");
    let photoUrls = [];

    try {
      if (rawPhotos.startsWith("[")) {
        let parsed = JSON.parse(rawPhotos);
        photoUrls = parsed.map(formatGoogleDriveImageUrl);
      }
    } catch (e) {
      photoUrls = [];
    }

    let bookingUrlPath = rawSlug.startsWith('/') ? rawSlug : '/' + rawSlug;
    if (!bookingUrlPath.endsWith('/')) bookingUrlPath += '/';

    let finalBookingUrl = "https://a.l8street.com" + bookingUrlPath;
    finalBookingUrl = finalBookingUrl.replace(/18street/g, "l8street");

    map[key] = {
      name: String(mData[i][0]).trim(),
      bookingUrl: finalBookingUrl,
      address: String(mData[i][5] || "").trim(),
      bedrooms: parseInt(mData[i][14]) || 1,
      beds: parseInt(mData[i][15]) || 1,
      bathrooms: parseInt(mData[i][16]) || 1,
      netPricePerNight: parseFloat(mData[i][17]) || 0,
      cleaningFeeNet: parseFloat(mData[i][18]) || 0,
      photos: photoUrls
    };
  }
  return map;
}

function extractFirstSlugKey(linksStr) {
  let matches = linksStr.match(/https?:\/\/[^\s,]+/g);
  if (!matches || matches.length === 0) return "";
  let firstUrl = matches[0].split(" ")[0];
  return cleanSlugKey(firstUrl);
}

function formatGoogleDriveImageUrl(url) {
  if (!url) return "";
  if (url.includes("drive.google.com") || url.includes("lh3.googleusercontent.com")) {
    let match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      return "https://lh3.googleusercontent.com/d/" + match[1] + "=s800";
    }
  }
  return url;
}

function parseJsonMeta(jsonRaw) {
  try {
    if (!jsonRaw) return {};
    if (typeof jsonRaw === 'object') return jsonRaw;
    return JSON.parse(jsonRaw);
  } catch (e) {
    return {};
  }
}

function formatDateGerman(timeMs) {
  let d = new Date(timeMs);
  let day = ("0" + d.getUTCDate()).slice(-2);
  let month = ("0" + (d.getUTCMonth() + 1)).slice(-2);
  return day + "." + month + "." + d.getUTCFullYear();
}

function parseGermanDateRange(dateStr) {
  try {
    let parts = String(dateStr).split('-');
    if (parts.length !== 2) return null;

    let d1 = parts[0].trim().split('.');
    let d2 = parts[1].trim().split('.');

    return {
      start: Date.UTC(d1[2], d1[1] - 1, d1[0]),
      end: Date.UTC(d2[2], d2[1] - 1, d2[0])
    };
  } catch (e) {
    return null;
  }
}