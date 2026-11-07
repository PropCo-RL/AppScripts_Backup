// =================================================================
// STUFE 3: GMAIL DRAFT GENERATOR (Schreibt Status in Spalte T)
// HTML-Layout: Klickbare Links, fette Kopfpreise & schlanke Kalkulation
// =================================================================
function generateEmailDraftsDMZ2() {
  const SHEET_LEADS_ID = '1VB253Cq_hr4YfFrWDXNG-pgv8qOOKAzb7jN_KYciEHw';
  const SHEET_INVENTORY_ID = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';

  let leadsSheet, invSheetM;
  try {
    const leadsSS = SpreadsheetApp.openById(SHEET_LEADS_ID);
    leadsSheet = leadsSS.getSheetByName('DMZ2');

    const invSS = SpreadsheetApp.openById(SHEET_INVENTORY_ID);
    invSheetM = invSS.getSheetByName('M');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Tabellen: " + e.toString());
    return;
  }

  if (!leadsSheet || !invSheetM) {
    Logger.log("FEHLER: Tabellenblatt DMZ2 oder M fehlt.");
    return;
  }

  const lastRow = leadsSheet.getLastRow();
  if (lastRow <= 1) return;

  // Header für Spalte T sicherstellen
  const headerCell = leadsSheet.getRange(1, 20);
  if (!headerCell.getValue()) {
    headerCell.setValue("E-Mail Status");
    headerCell.setFontWeight("bold");
    leadsSheet.getRange("T:T").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // 1. Apartment-Stammdaten & Preise aus Tab M laden
  const mData = invSheetM.getDataRange().getValues();
  const aptDetailsMap = {};
  for (let r = 1; r < mData.length; r++) {
    const rawSlug = String(mData[r][6] || "").trim();
    const code = cleanShortCode(rawSlug);
    if (!code) continue;

    aptDetailsMap[code] = {
      address: String(mData[r][5] || "").trim() || "Monteurwohnung",
      schlafzimmer: Number(mData[r][14]) || 2,
      maxPersonen: Number(mData[r][15]) || 4,
      netPricePerNight: parseFloat(mData[r][17]) || 85.00,
      isCustomPrice: !isNaN(parseFloat(mData[r][17])) && parseFloat(mData[r][17]) > 0
    };
  }

  // 2. Leads einlesen
  const numRows = lastRow - 1;
  const dataRange = leadsSheet.getRange(2, 1, numRows, 20);
  const rows = dataRange.getValues();
  let updatesMade = false;

  const footerText =
    "\n\nFreundliche Grüße\n\n" +
    "L8 Street\n\n" +
    "© 2017-2026\n" +
    "L8 Street GmbH\n" +
    "Postfach CAYA595616\n" +
    "96035 Bamberg\n\n" +
    "WhatsApp: +4917684801295\n" +
    "E-Mail: support@L8Street.com\n" +
    "Web: https://www.L8Street.com\n" +
    "Geschäftsführer: Raul Leneweit\n" +
    "Sitz der Gesellschaft: Forststraße 65,\n" +
    "75223 Niefern-Öschelbronn\n" +
    "Amtsgericht Mannheim HRB: 72 85 52\n" +
    "USt-IdNr.: DE 314 603 427";

  for (let i = 0; i < rows.length; i++) {
    const cellB = String(rows[i][1] || "").trim();
    const routing = String(rows[i][14] || "").trim();
    const optSingle = String(rows[i][15] || "").trim();
    const optKombi = String(rows[i][16] || "").trim();
    const mailStatus = String(rows[i][19] || "").trim();

    if (!cellB || cellB === "-" || mailStatus !== "") {
      continue;
    }

    let leadJson = {}, dauerData = {}, spezialData = {}, persData = {};
    try {
      leadJson = JSON.parse(cellB);
      persData = JSON.parse(rows[i][5] || "{}");
      dauerData = JSON.parse(rows[i][6] || "{}");
      spezialData = JSON.parse(rows[i][7] || "{}");
    } catch (e) {
      continue;
    }

    const recipientEmail = String(spezialData.direktEmail || leadJson.email || "").trim();
    if (!recipientEmail || !recipientEmail.includes("@")) {
      rows[i][19] = "FEHLER: Keine E-Mail";
      updatesMade = true;
      continue;
    }

    const anredeZeile = buildSalutation(spezialData.ansprechpartner, leadJson.gastName);
    const kundenNummer = leadJson.kundenNummer || "";
    let mailSubject = "Buchungsanfrage - Monteurwohnung L8 Street";
    if (kundenNummer && kundenNummer !== "N/A") {
      mailSubject += " (Anfrage ID: " + kundenNummer + ")";
    }

    const persCount = Math.max(1, Number(persData.Pers) || Number(leadJson.personen) || 1);
    const nights = Math.max(1, Number(dauerData.tage) > 0 ? Number(dauerData.tage) : 1);

    // -------------------------------------------------------------
    // FALL 1: K.O. ABSAGE
    // -------------------------------------------------------------
    if (routing === "REJECT_KO" || routing === "REJECT_1PERS") {
      const mailBodyPlain = anredeZeile + "\n\n" +
        "vielen Dank für Ihre Buchungsanfrage.\n\n" +
        "Tut mir schrecklich leid, wir können aktuell leider garnichts anbieten.\n\n" +
        "Vielen Dank für Ihr Verständnis." +
        footerText;

      createGmailDraftFormatted(recipientEmail, mailSubject, mailBodyPlain);
      rows[i][19] = "ENTWURF: Absage (KO)";
      updatesMade = true;
      continue;
    }

    // -------------------------------------------------------------
    // FALL 2: SINGLE APARTMENT (<= 6 Personen)
    // -------------------------------------------------------------
    if (routing === "STD_SHORT" || routing === "STD_LONG") {
      let singleList = [];
      try { singleList = JSON.parse(optSingle || "[]"); } catch (e) {}

      if (singleList.length > 0) {
        const best = singleList[0];
        const cleanId = cleanShortCode(best.id);
        const apt = aptDetailsMap[cleanId] || {
          address: "Monteurwohnung",
          schlafzimmer: 2,
          maxPersonen: persCount,
          netPricePerNight: 85.00,
          isCustomPrice: false
        };
        const link = `https://a.L8Street.com/a/${cleanId}/`;
        const bettenProZimmer = Math.max(1, Math.round(apt.maxPersonen / apt.schlafzimmer));

        // Kalkulation
        const nachtNetto = roundEur(apt.netPricePerNight);
        const reinigungNetto = nachtNetto;
        const mieteNetto = roundEur(nights * nachtNetto);
        const zwischensummeNetto = roundEur(mieteNetto + reinigungNetto);
        const mwstBetrag = roundEur(zwischensummeNetto * 0.07);
        const gesamtBrutto = roundEur(zwischensummeNetto + mwstBetrag);

        const kopfNetto = roundEur(nachtNetto / persCount);
        const kopfBrutto = roundEur(kopfNetto * 1.07);

        // Plain-Text Fallback
        const plainBody = anredeZeile + "\n\n" +
          "sehr gerne, wir haben noch die folgende Unterkunft frei:\n\n" +
          `Adresse: ${apt.address}\n` +
          `- ${apt.schlafzimmer} separate Schlafzimmer mit ${apt.maxPersonen} Einzelbetten (${bettenProZimmer} pro Schlafzimmer)\n\n` +
          "- Voll ausgestattete Küche, Bad, Waschmaschine, WLAN & Parkplätze\n\n" +
          `Fotos & Details: ${link}\n\n` +
          "Kosten:\n" +
          `- ${formatEur(kopfNetto)} netto (${formatEur(kopfBrutto)} brutto) pro Person / Nacht\n` +
          `  (entspricht ${formatEur(nachtNetto)} netto pro Nacht für die gesamte Unterkunft)\n\n` +
          "Für den gesamten Zeitraum:\n" +
          `- Miete: ${nights} Nächte x ${formatEur(nachtNetto)} = ${formatEur(mieteNetto)}\n` +
          `- Endreinigung: ${formatEur(reinigungNetto)}\n` +
          `- Gesamtbetrag: ${formatEur(gesamtBrutto)} brutto (${formatEur(zwischensummeNetto)} netto zzgl. 7% MwSt.)\n\n` +
          "Keine weiteren Kosten. AGB: https://L8Street.com/agb\n\n" +
          "Passt diese Wohnung? Antworten Sie einfach kurz hier oder direkt per WhatsApp unter +49 176 84801295 für eine sofortige Reservierung." +
          footerText;

        // Klickbarer HTML-Body
        const htmlBody = escapeHtml(anredeZeile) + "<br><br>" +
          "sehr gerne, wir haben noch die folgende Unterkunft frei:<br><br>" +
          `Adresse: ${escapeHtml(apt.address)}<br>` +
          `- ${apt.schlafzimmer} separate Schlafzimmer mit ${apt.maxPersonen} Einzelbetten (${bettenProZimmer} pro Schlafzimmer)<br><br>` +
          "- Voll ausgestattete Küche, Bad, Waschmaschine, WLAN & Parkplätze<br><br>" +
          `Fotos & Details: <a href="${link}">${link}</a><br><br>` +
          "Kosten:<br>" +
          `<b>- ${formatEur(kopfNetto)} netto (${formatEur(kopfBrutto)} brutto) pro Person / Nacht</b><br>` +
          `&nbsp;&nbsp;(entspricht ${formatEur(nachtNetto)} netto pro Nacht für die gesamte Unterkunft)<br><br>` +
          "Für den gesamten Zeitraum:<br>" +
          `- Miete: ${nights} Nächte x ${formatEur(nachtNetto)} = ${formatEur(mieteNetto)}<br>` +
          `- Endreinigung: ${formatEur(reinigungNetto)}<br>` +
          `- Gesamtbetrag: ${formatEur(gesamtBrutto)} brutto (${formatEur(zwischensummeNetto)} netto zzgl. 7% MwSt.)<br><br>` +
          `Keine weiteren Kosten. AGB: <a href="https://L8Street.com/agb">https://L8Street.com/agb</a><br><br>` +
          `Passt diese Wohnung? Antworten Sie einfach kurz hier oder direkt per WhatsApp unter <a href="https://wa.me/4917684801295">+49 176 84801295</a> für eine sofortige Reservierung.` +
          footerTextToHtml(footerText);

        createGmailDraftFormatted(recipientEmail, mailSubject, plainBody, htmlBody);
        const warnTag = !apt.isCustomPrice ? " (Achtung: Standardpreis)" : "";
        rows[i][19] = `ENTWURF: Angebot (${cleanId})${warnTag}`;
        updatesMade = true;
      } else {
        const plainFail = buildNoAvailabilityMail(anredeZeile, footerText);
        createGmailDraftFormatted(recipientEmail, mailSubject, plainFail);
        rows[i][19] = "ENTWURF: Keine Verfügbarkeit";
        updatesMade = true;
      }
      continue;
    }

    // -------------------------------------------------------------
    // FALL 3: KOMBI-ANGEBOT (> 6 Personen)
    // -------------------------------------------------------------
    if (routing === "BIG_SHORT" || routing === "BIG_LONG") {
      let kombiList = [];
      try { kombiList = JSON.parse(optKombi || "[]"); } catch (e) {}

      if (kombiList.length > 0) {
        const bestKombi = kombiList[0];
        const isSingleUnitKombi = (bestKombi.apts.length === 1);
        let summeNachtNetto = 0;
        let plainWohnungen = "";
        let htmlWohnungen = "";
        let usedStandardPrice = false;

        bestKombi.apts.forEach((id, idx) => {
          const cleanId = cleanShortCode(id);
          const apt = aptDetailsMap[cleanId] || {
            address: "Monteurwohnung",
            schlafzimmer: 2,
            maxPersonen: 4,
            netPricePerNight: 85.00,
            isCustomPrice: false
          };
          if (!apt.isCustomPrice) usedStandardPrice = true;

          const nNetto = roundEur(apt.netPricePerNight);
          summeNachtNetto = roundEur(summeNachtNetto + nNetto);
          const bettenProZimmer = Math.max(1, Math.round(apt.maxPersonen / apt.schlafzimmer));
          const prefix = isSingleUnitKombi ? "Adresse:" : `Wohnung ${idx + 1}:`;
          const link = `https://a.L8Street.com/a/${cleanId}/`;

          plainWohnungen += `${prefix} ${apt.address}\n` +
            `- ${apt.schlafzimmer} separate Schlafzimmer mit ${apt.maxPersonen} Einzelbetten (${bettenProZimmer} pro Schlafzimmer)\n\n` +
            "- Voll ausgestattete Küche, Bad, Waschmaschine, WLAN & Parkplätze\n\n`" +
            `Fotos & Details: ${link}\n\n`;

          htmlWohnungen += `${prefix} ${escapeHtml(apt.address)}<br>` +
            `- ${apt.schlafzimmer} separate Schlafzimmer mit ${apt.maxPersonen} Einzelbetten (${bettenProZimmer} pro Schlafzimmer)<br><br>` +
            "- Voll ausgestattete Küche, Bad, Waschmaschine, WLAN & Parkplätze<br><br>" +
            `Fotos & Details: <a href="${link}">${link}</a><br><br>`;
        });

        const summeReinigungNetto = summeNachtNetto;
        const mieteNetto = roundEur(nights * summeNachtNetto);
        const zwischensummeNetto = roundEur(mieteNetto + summeReinigungNetto);
        const mwstBetrag = roundEur(zwischensummeNetto * 0.07);
        const gesamtBrutto = roundEur(zwischensummeNetto + mwstBetrag);

        const kopfNetto = roundEur(summeNachtNetto / persCount);
        const kopfBrutto = roundEur(kopfNetto * 1.07);

        const introTextPlain = isSingleUnitKombi
          ? "sehr gerne, wir haben noch die folgende Unterkunft frei:\n\n"
          : `sehr gerne, wir können Ihr Team in folgenden ${bestKombi.apts.length} Unterkünften gemeinsam unterbringen:\n\n`;

        const introTextHtml = isSingleUnitKombi
          ? "sehr gerne, wir haben noch die folgende Unterkunft frei:<br><br>"
          : `sehr gerne, wir können Ihr Team in folgenden ${bestKombi.apts.length} Unterkünften gemeinsam unterbringen:<br><br>`;

        const subNote = isSingleUnitKombi
          ? `(entspricht ${formatEur(summeNachtNetto)} netto pro Nacht für die gesamte Unterkunft)`
          : `(entspricht ${formatEur(summeNachtNetto)} netto pro Nacht für alle Wohnungen gesamt)`;

        const plainBody = anredeZeile + "\n\n" +
          introTextPlain +
          plainWohnungen +
          "Kosten:\n" +
          `- ${formatEur(kopfNetto)} netto (${formatEur(kopfBrutto)} brutto) pro Person / Nacht\n` +
          `  ${subNote}\n\n` +
          "Für den gesamten Zeitraum:\n" +
          `- Miete: ${nights} Nächte x ${formatEur(summeNachtNetto)} = ${formatEur(mieteNetto)}\n` +
          `- Endreinigung: ${formatEur(summeReinigungNetto)}\n` +
          `- Gesamtbetrag: ${formatEur(gesamtBrutto)} brutto (${formatEur(zwischensummeNetto)} netto zzgl. 7% MwSt.)\n\n` +
          "Keine weiteren Kosten. AGB: https://L8Street.com/agb\n\n" +
          "Passt diese Aufteilung für Ihr Team? Antworten Sie einfach kurz hier oder direkt per WhatsApp unter +49 176 84801295 für eine sofortige Reservierung." +
          footerText;

        const htmlBody = escapeHtml(anredeZeile) + "<br><br>" +
          introTextHtml +
          htmlWohnungen +
          "Kosten:<br>" +
          `<b>- ${formatEur(kopfNetto)} netto (${formatEur(kopfBrutto)} brutto) pro Person / Nacht</b><br>` +
          `&nbsp;&nbsp;${subNote}<br><br>` +
          "Für den gesamten Zeitraum:<br>" +
          `- Miete: ${nights} Nächte x ${formatEur(summeNachtNetto)} = ${formatEur(mieteNetto)}<br>` +
          `- Endreinigung: ${formatEur(summeReinigungNetto)}<br>` +
          `- Gesamtbetrag: ${formatEur(gesamtBrutto)} brutto (${formatEur(zwischensummeNetto)} netto zzgl. 7% MwSt.)<br><br>` +
          `Keine weiteren Kosten. AGB: <a href="https://L8Street.com/agb">https://L8Street.com/agb</a><br><br>` +
          `Passt diese Aufteilung für Ihr Team? Antworten Sie einfach kurz hier oder direkt per WhatsApp unter <a href="https://wa.me/4917684801295">+49 176 84801295</a> für eine sofortige Reservierung.` +
          footerTextToHtml(footerText);

        createGmailDraftFormatted(recipientEmail, mailSubject, plainBody, htmlBody);
        const warnTag = usedStandardPrice ? " (Achtung: Standardpreis)" : "";
        rows[i][19] = `ENTWURF: Kombi (${bestKombi.apts.map(cleanShortCode).join('+')})${warnTag}`;
        updatesMade = true;
      } else {
        const plainFail = buildNoAvailabilityMail(anredeZeile, footerText);
        createGmailDraftFormatted(recipientEmail, mailSubject, plainFail);
        rows[i][19] = "ENTWURF: Keine Verfügbarkeit";
        updatesMade = true;
      }
      continue;
    }
  }

  if (updatesMade) {
    dataRange.setValues(rows);
    SpreadsheetApp.flush();
    Logger.log("=== ALLE ENTWÜRFE ERFOLGREICH ANGELEGT (SPALTE T BEFÜLLT) ===");
  }
}

// -----------------------------------------------------------------
// HILFSFUNKTIONEN
// -----------------------------------------------------------------

function buildSalutation(ansprechpartner, fallbackName) {
  let name = (ansprechpartner && ansprechpartner.trim().length > 2) ? ansprechpartner.trim() : (fallbackName || "").trim();

  if (!name || name === "Damen und Herren" || name.toLowerCase().startsWith("firma")) {
    return "Guten Tag,";
  }

  if (/^herr\s+/i.test(name)) {
    const parts = name.replace(/^herr\s+/i, "").trim().split(/\s+/);
    return `Sehr geehrter Herr ${parts[parts.length - 1]},`;
  }

  if (/^frau\s+/i.test(name)) {
    const parts = name.replace(/^frau\s+/i, "").trim().split(/\s+/);
    return `Sehr geehrte Frau ${parts[parts.length - 1]},`;
  }

  return `Guten Tag ${name},`;
}

function roundEur(val) {
  const num = Number(val);
  if (isNaN(num)) return 0;
  return Math.round((num + Number.EPSILON) * 100) / 100;
}

function formatEur(num) {
  const val = roundEur(num);
  return val.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
}

function cleanShortCode(rawStr) {
  if (!rawStr) return "";
  return String(rawStr).trim().toLowerCase()
    .replace(/^https?:\/\/a\.l8street\.com\/a\//i, "")
    .replace(/^\/?a\//i, "")
    .replace(/^\//, "")
    .replace(/\/$/, "");
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function footerTextToHtml(footer) {
  return "<br><br>" + escapeHtml(footer.trim())
    .replace(/\n/g, "<br>")
    .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1">$1</a>');
}

/**
 * Erstellt den Gmail Draft mit HTML-Formatierung (fett, Links etc.)
 */
function createGmailDraftFormatted(to, subject, plainText, htmlText) {
  try {
    if (htmlText) {
      GmailApp.createDraft(to, subject, plainText, {
        htmlBody: htmlText
      });
    } else {
      GmailApp.createDraft(to, subject, plainText);
    }
  } catch (err) {
    Logger.log("Fehler bei Entwurf-Erstellung für " + to + ": " + err.toString());
  }
}

function buildNoAvailabilityMail(anredeZeile, footerText) {
  return anredeZeile + "\n\n" +
    "vielen Dank für Ihre Buchungsanfrage.\n\n" +
    "Leider haben wir im von Ihnen gewünschten Zeitraum direkt vor Ort keine exakt passende Kapazität frei.\n\n" +
    "Auf unserer Website finden Sie eine Übersicht weiterer Wohnungen die in den kommenden Tagen frei werden: a.L8Street.com\n\n" +
    "Schauen Sie dort gerne durch oder geben Sie uns kurz Bescheid, falls auch ein leicht verschobener Zeitraum für Sie infrage kommt." +
    footerText;
}