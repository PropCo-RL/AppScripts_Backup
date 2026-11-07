// =================================================================
// GMAIL & TEMPLATE SERVICE (KOMPLETT)
// Mail-Generierung, Thread-Handling, Footer, Zitatblock & Formatierungen
// Inklusive Unterstützung für Next-Available Alternativangebote & Sheets
// =================================================================

function cleanPortalSubject(rawSubject) {
  if (!rawSubject) return "Re: Neue Buchungsanfrage auf mein-Monteurzimmer.de";
  let s = String(rawSubject).trim();
  s = s.replace(/^(re|fwd|aw|wg):\s*/i, "");
  s = s.replace(/\[recovered\]/gi, "").trim();
  s = s.replace(/\s*-\s*\[MMZ\s*#[^\]]+\]/gi, "").trim();
  if (!s) s = "Neue Buchungsanfrage auf mein-Monteurzimmer.de";
  return "Re: " + s;
}

function createOrReplyDraft(to, subject, plainText, htmlText, forwardMessageId, forwardThreadId) {
  if (!to) {
    Logger.log("FEHLER: Kein Empfänger (to) übergeben.");
    return { draftId: "", inThread: false };
  }

  // 1. Versuch: Über forwardMessageId antworten
  if (forwardMessageId) {
    try {
      const origMsg = GmailApp.getMessageById(forwardMessageId);
      if (origMsg) {
        const origDate = Utilities.formatDate(origMsg.getDate(), "Europe/Zurich", "dd.MM.yyyy HH:mm");
        const origSender = origMsg.getFrom();
        const origSubject = origMsg.getSubject();
        const quotedPlain = `\n\n--- Ursprüngliche Nachricht ---\nVon: ${origSender}\nDatum: ${origDate}\nBetreff: ${origSubject}\n\n` +
          origMsg.getPlainBody().split("\n").map(line => `> ${line}`).join("\n");
        const quotedHtml = `<br><br><hr style="border:none;border-top:1px solid #ccc;margin:20px 0;">` +
          `<div style="color:#555;font-size:13px;"><b>Ursprüngliche Nachricht:</b><br>` +
          `<b>Von:</b> ${escapeHtml(origSender)}<br>` +
          `<b>Datum:</b> ${origDate}<br>` +
          `<b>Betreff:</b> ${escapeHtml(origSubject)}</div><br>` +
          `<blockquote style="border-left:2px solid #ccc;margin:0 0 0 10px;padding-left:10px;color:#444;">` +
          origMsg.getBody() +
          `</blockquote>`;

        const fullPlain = plainText + quotedPlain;
        const fullHtml = htmlText ? (htmlText + quotedHtml) : "";
        const replyOptions = { replyTo: "support@l8street.com" };
        if (fullHtml) replyOptions.htmlBody = fullHtml;

        const draft = origMsg.createDraftReply(fullPlain, replyOptions);
        if (draft) {
          const baseSub = origMsg.getSubject() || origMsg.getThread().getFirstMessageSubject() || subject;
          const finalSubject = cleanPortalSubject(baseSub);
          draft.update(to, finalSubject, fullPlain, replyOptions);
          return { draftId: draft.getId(), inThread: true };
        }
      }
    } catch (e) {
      Logger.log("Antwort via forwardMessageId fehlgeschlagen: " + e.toString());
    }
  }

  // 2. Versuch: Über forwardThreadId antworten
  if (forwardThreadId) {
    try {
      const thread = GmailApp.getThreadById(forwardThreadId);
      if (thread) {
        const messages = thread.getMessages();
        const lastMsg = messages[messages.length - 1];
        let quotedPlain = "";
        let quotedHtml = "";
        if (lastMsg) {
          const origDate = Utilities.formatDate(lastMsg.getDate(), "Europe/Zurich", "dd.MM.yyyy HH:mm");
          const origSender = lastMsg.getFrom();
          const origSubject = lastMsg.getSubject();
          quotedPlain = `\n\n--- Ursprüngliche Nachricht ---\nVon: ${origSender}\nDatum: ${origDate}\nBetreff: ${origSubject}\n\n` +
            lastMsg.getPlainBody().split("\n").map(line => `> ${line}`).join("\n");
          quotedHtml = `<br><br><hr style="border:none;border-top:1px solid #ccc;margin:20px 0;">` +
            `<div style="color:#555;font-size:13px;"><b>Ursprüngliche Nachricht:</b><br>` +
            `<b>Von:</b> ${escapeHtml(origSender)}<br>` +
            `<b>Datum:</b> ${origDate}<br>` +
            `<b>Betreff:</b> ${escapeHtml(origSubject)}</div><br>` +
            `<blockquote style="border-left:2px solid #ccc;margin:0 0 0 10px;padding-left:10px;color:#444;">` +
            lastMsg.getBody() +
            `</blockquote>`;
        }
        const fullPlain = plainText + quotedPlain;
        const fullHtml = htmlText ? (htmlText + quotedHtml) : "";
        const replyOptions = { replyTo: "support@l8street.com" };
        if (fullHtml) replyOptions.htmlBody = fullHtml;

        const draft = thread.createDraftReply(fullPlain, replyOptions);
        if (draft) {
          const baseSub = thread.getFirstMessageSubject() || subject;
          const finalSubject = cleanPortalSubject(baseSub);
          draft.update(to, finalSubject, fullPlain, replyOptions);
          return { draftId: draft.getId(), inThread: true };
        }
      }
    } catch (e) {
      Logger.log("Antwort via forwardThreadId fehlgeschlagen: " + e.toString());
    }
  }

  // 3. Fallback: Standalone Draft
  try {
    const finalSubject = cleanPortalSubject(subject);
    const replyOptions = { replyTo: "support@l8street.com" };
    if (htmlText) replyOptions.htmlBody = htmlText;
    const draft = GmailApp.createDraft(to, finalSubject, plainText, replyOptions);
    return { draftId: draft ? draft.getId() : "", inThread: false };
  } catch (err) {
    Logger.log("Fehler bei Standalone-Entwurf: " + err.toString());
    return { draftId: "", inThread: false };
  }
}

function handleMissingCustomerEmail(leadJson, labelName) {
  const forwardThreadId = leadJson.forwardThreadId || "";
  let labelApplied = false;
  if (forwardThreadId) {
    try {
      let labelObj = GmailApp.getUserLabelByName(labelName);
      if (!labelObj) {
        labelObj = GmailApp.createLabel(labelName);
      }
      const thread = GmailApp.getThreadById(forwardThreadId);
      if (thread) {
        const existingLabels = thread.getLabels().map(l => l.getName());
        if (!existingLabels.includes(labelName)) {
          thread.addLabel(labelObj);
          const msgs = thread.getMessages();
          if (msgs.length > 0) {
            msgs[msgs.length - 1].markUnread();
          }
          labelApplied = true;
        }
      }
    } catch (lblErr) {
      Logger.log("Konnte Label nicht anheften: " + lblErr.toString());
    }
  }
  return {
    status: "MANUELL_PORTAL_BEANTWORTEN",
    label: labelName,
    threadGelabelt: labelApplied,
    portalUrl: leadJson.anfrageUrl || "https://mein-monteurzimmer.de",
    hinweis: "Keine Direkt-Mail vorhanden. Bitte manuell im Portal antworten.",
    zeitpunkt: new Date().toISOString()
  };
}

function loadApartmentInventoryMap(invSheetM) {
  const mData = invSheetM.getDataRange().getValues();
  const map = {};
  for (let r = 1; r < mData.length; r++) {
    const rawSlug = String(mData[r][6] || "").trim();
    const code = cleanShortCode(rawSlug);
    if (!code) continue;
    map[code] = {
      address: String(mData[r][5] || "").trim() || "Monteurwohnung",
      schlafzimmer: Number(mData[r][14]) || 2,
      maxPersonen: Number(mData[r][15]) || 4,
      netPricePerNight: parseFloat(mData[r][17]) || 85.00,
      isCustomPrice: !isNaN(parseFloat(mData[r][17])) && parseFloat(mData[r][17]) > 0
    };
  }
  return map;
}

// -------------------------------------------------------------
// TEXT- & TEMPLATE-GENERATOREN
// -------------------------------------------------------------

function formatGermanDate(isoStr) {
  if (!isoStr) return "";
  const parts = String(isoStr).split("-");
  if (parts.length === 3) return `${parts[2]}.${parts[1]}.${parts[0]}`;
  return isoStr;
}

function checkIsLongstayVip(persCount, nights, leadJson) {
  if (persCount < 3) return false;
  if (nights >= 60) return true;
  let textSource = "";
  if (leadJson && typeof leadJson === "object") {
    textSource = String(leadJson.nachricht || leadJson.bemerkung || leadJson.text || leadJson.notiz || "");
  }
  const keywords = ["verlänger", "verlaenger", "monat", "1 jahr", "ein jahr", "unbefristet", "dauermiete", "länger", "laenger"];
  const lowerText = textSource.toLowerCase();
  return keywords.some(kw => lowerText.includes(kw));
}

function getCtaBlock(isVip, istKombi) {
  if (isVip) {
    const frage = istKombi ? "Passt diese Aufteilung für Ihr Team?" : "Passt diese Wohnung für Ihr Team?";
    const plain = `${frage}\n\n` +
      "Da Sie eine längerfristige Unterbringung mit Verlängerungsoption anfragen: Wir passen die Ausstattung der Wohnungen sehr gerne an die individuellen Bedürfnisse Ihres Teams an. Unser Hausmeister vor Ort sowie ich persönlich stehen Ihnen jederzeit zur Verfügung, damit alles absolut reibungslos läuft.\n\n" +
      "Lassen Sie uns die Details gerne kurz telefonisch besprechen oder eine Besichtigung vor Ort vereinbaren. Sie erreichen mich direkt per WhatsApp oder telefonisch unter +49 176 84801295.";
    const html = `<b>${frage}</b><br><br>` +
      "Da Sie eine längerfristige Unterbringung mit Verlängerungsoption anfragen: Wir passen die Ausstattung der Wohnungen sehr gerne an die individuellen Bedürfnisse Ihres Teams an. Unser Hausmeister vor Ort sowie ich persönlich stehen Ihnen jederzeit zur Verfügung, damit alles absolut reibungslos läuft.<br><br>" +
      `Lassen Sie uns die Details gerne kurz telefonisch besprechen oder eine Besichtigung vor Ort vereinbaren. Sie erreichen mich direkt mobil oder per WhatsApp unter <a href="https://wa.me/4917684801295">+49 176 84801295</a>.`;
    return { plain, html };
  }

  const frage = istKombi ? "Passt diese Aufteilung für Sie?" : "Passt diese Wohnung für Sie?";
  const plain = `${frage} Gerne erreichen Sie mich auch per WhatsApp an +49 176 84801295.`;
  const html = `${frage} Gerne erreichen Sie mich auch per WhatsApp an <a href="https://wa.me/4917684801295">+49 176 84801295</a>.`;
  return { plain, html };
}

function formatZimmerText(schlafzimmer, betten) {
  const zimmerWort = schlafzimmer === 1 ? "1 separates Schlafzimmer" : `${schlafzimmer} separate Schlafzimmer`;
  const bettenWort = betten === 1 ? "1 Einzelbett" : `${betten} Einzelbetten`;
  const bettenProZimmer = Math.max(1, Math.round(betten / schlafzimmer));
  return `${zimmerWort} mit ${bettenWort} (${bettenProZimmer} pro Schlafzimmer)`;
}

function buildRejectionContent(anredeZeile, footerText) {
  const plainBody = anredeZeile + "\n\n" +
    "vielen Dank für Ihre Buchungsanfrage.\n\n" +
    "Tut mir schrecklich leid, wir können aktuell leider garnichts anbieten.\n\n" +
    "Vielen Dank für Ihr Verständnis." +
    footerText;

  const htmlBody = escapeHtml(anredeZeile) + "<br><br>" +
    "vielen Dank für Ihre Buchungsanfrage.<br><br>" +
    "Tut mir schrecklich leid, wir können aktuell leider garnichts anbieten.<br><br>" +
    "Vielen Dank für Ihr Verständnis." +
    footerTextToHtml(footerText);

  return { plainBody, htmlBody };
}

function buildNoAvailabilityContent(anredeZeile, footerText) {
  const plainBody = anredeZeile + "\n\n" +
    "vielen Dank für Ihre Buchungsanfrage.\n\n" +
    "Leider haben wir im von Ihnen gewünschten Zeitraum direkt vor Ort keine exakt passende Kapazität frei.\n\n" +
    "Auf unserer Website finden Sie eine Übersicht weiterer Wohnungen die in den kommenden Tagen frei werden: a.L8Street.com\n\n" +
    "Schauen Sie dort gerne durch oder geben Sie uns kurz Bescheid, falls auch ein leicht verschobener Zeitraum für Sie infrage kommt." +
    footerText;

  const htmlBody = escapeHtml(anredeZeile) + "<br><br>" +
    "vielen Dank für Ihre Buchungsanfrage.<br><br>" +
    "Leider haben wir im von Ihnen gewünschten Zeitraum direkt vor Ort keine exakt passende Kapazität frei.<br><br>" +
    `Auf unserer Website finden Sie eine Übersicht weiterer Wohnungen die in den kommenden Tagen frei werden: <a href="https://a.L8Street.com">a.L8Street.com</a><br><br>` +
    "Schauen Sie dort gerne durch oder geben Sie uns kurz Bescheid, falls auch ein leicht verschobener Zeitraum für Sie infrage kommt." +
    footerTextToHtml(footerText);

  return { plainBody, htmlBody };
}

function buildSingleOfferContent(best, aptDetailsMap, persCount, nights, anredeZeile, footerText, leadJson, altObj) {
  const cleanId = cleanShortCode(best.id);
  const apt = aptDetailsMap[cleanId] || {
    address: "Monteurwohnung",
    schlafzimmer: 2,
    maxPersonen: persCount,
    netPricePerNight: 85.00,
    isCustomPrice: false
  };

  const link = `https://a.L8Street.com/a/${cleanId}/`;
  const zimmerInfo = formatZimmerText(apt.schlafzimmer, apt.maxPersonen);
  const nachtNetto = roundEur(apt.netPricePerNight);
  const reinigungNetto = nachtNetto;
  const mieteNetto = roundEur(nights * nachtNetto);
  const zwischensummeNetto = roundEur(mieteNetto + reinigungNetto);
  const mwstBetrag = roundEur(zwischensummeNetto * 0.07);
  const gesamtBrutto = roundEur(zwischensummeNetto + mwstBetrag);
  const kopfNetto = roundEur(nachtNetto / persCount);
  const isVip = checkIsLongstayVip(persCount, nights, leadJson);
  const cta = getCtaBlock(isVip, false);

  let introPlain = "sehr gerne, wir haben noch die folgende Unterkunft frei:\n\n";
  let introHtml = "sehr gerne, wir haben noch die folgende Unterkunft frei:<br><br>";

  if (altObj && altObj.altStart) {
    const formattedAltStart = formatGermanDate(altObj.altStart);
    introPlain = `vielen Dank für Ihre Anfrage. Zum direkt angefragten Termin sind wir leider bereits voll belegt. Wir können Ihnen diese passende Unterkunft jedoch gerne ab dem ${formattedAltStart} anbieten:\n\n`;
    introHtml = `vielen Dank für Ihre Anfrage. Zum direkt angefragten Termin sind wir leider bereits voll belegt. Wir können Ihnen diese passende Unterkunft jedoch gerne <b>ab dem ${formattedAltStart}</b> anbieten:<br><br>`;
  }

  const plainBody = anredeZeile + "\n\n" +
    introPlain +
    `Adresse: ${apt.address}\n` +
    `Fotos & Details auf unserer Website: ${link}\n` +
    `- ${zimmerInfo}\n` +
    "- Voll ausgestattete Küche, Bad, Waschmaschine, WLAN & Parkplätze\n\n" +
    `Kosten: ${formatEur(kopfNetto)} netto pro Person / Nacht (entspricht ${formatEur(nachtNetto)} netto pro Nacht für die gesamte Unterkunft). Für den gesamten Zeitraum also ${nights} Nächte x ${formatEur(nachtNetto)} = ${formatEur(mieteNetto)}, Endreinigung nur Einmalig: ${formatEur(reinigungNetto)}, Gesamtbetrag: ${formatEur(gesamtBrutto)} brutto (${formatEur(zwischensummeNetto)} netto zzgl. 7% MwSt.). Keine weiteren Kosten. AGB: https://L8Street.com/agb\n\n` +
    cta.plain + footerText;

  const htmlBody = escapeHtml(anredeZeile) + "<br><br>" +
    introHtml +
    `<b>Adresse:</b> ${escapeHtml(apt.address)}<br>` +
    `Fotos & Details auf unserer Website: <a href="${link}">${link}</a><br>` +
    `- ${escapeHtml(zimmerInfo)}<br>` +
    "- Voll ausgestattete Küche, Bad, Waschmaschine, WLAN & Parkplätze<br><br>" +
    `<b>Kosten: ${formatEur(kopfNetto)} netto pro Person / Nacht</b> (entspricht ${formatEur(nachtNetto)} netto pro Nacht für die gesamte Unterkunft). Für den gesamten Zeitraum also ${nights} Nächte x ${formatEur(nachtNetto)} = ${formatEur(mieteNetto)}, Endreinigung nur Einmalig: ${formatEur(reinigungNetto)}, Gesamtbetrag: <b>${formatEur(gesamtBrutto)} brutto</b> (${formatEur(zwischensummeNetto)} netto zzgl. 7% MwSt.). Keine weiteren Kosten. AGB: <a href="https://L8Street.com/agb">https://L8Street.com/agb</a><br><br>` +
    cta.html + footerTextToHtml(footerText);

  return {
    cleanId,
    plainBody,
    htmlBody,
    gesamtBrutto,
    usedStandardPrice: !apt.isCustomPrice
  };
}

function buildKombiOfferContent(bestKombi, aptDetailsMap, persCount, nights, anredeZeile, footerText, leadJson, altObj) {
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
    const zimmerInfo = formatZimmerText(apt.schlafzimmer, apt.maxPersonen);
    const link = `https://a.L8Street.com/a/${cleanId}/`;

    plainWohnungen += `Wohnung ${idx + 1}: ${apt.address}\n` +
      `Fotos & Details auf unserer Website: ${link}\n` +
      `- ${zimmerInfo}\n` +
      "- Voll ausgestattete Küche, Bad, Waschmaschine, WLAN & Parkplätze\n\n";

    htmlWohnungen += `<b>Wohnung ${idx + 1}:</b> ${escapeHtml(apt.address)}<br>` +
      `Fotos & Details auf unserer Website: <a href="${link}">${link}</a><br>` +
      `- ${escapeHtml(zimmerInfo)}<br>` +
      "- Voll ausgestattete Küche, Bad, Waschmaschine, WLAN & Parkplätze<br><br>";
  });

  const summeReinigungNetto = summeNachtNetto;
  const mieteNetto = roundEur(nights * summeNachtNetto);
  const zwischensummeNetto = roundEur(mieteNetto + summeReinigungNetto);
  const mwstBetrag = roundEur(zwischensummeNetto * 0.07);
  const gesamtBrutto = roundEur(zwischensummeNetto + mwstBetrag);
  const kopfNetto = roundEur(summeNachtNetto / persCount);
  const isVip = checkIsLongstayVip(persCount, nights, leadJson);
  const cta = getCtaBlock(isVip, true);

  let introPlain = `sehr gerne, wir können Ihr Team in folgenden ${bestKombi.apts.length} Unterkünften gemeinsam unterbringen:\n\n`;
  let introHtml = `sehr gerne, wir können Ihr Team in folgenden ${bestKombi.apts.length} Unterkünften gemeinsam unterbringen:<br><br>`;

  if (altObj && altObj.altStart) {
    const formattedAltStart = formatGermanDate(altObj.altStart);
    introPlain = `vielen Dank für Ihre Anfrage. Zum direkt angefragten Termin haben wir leider keine zusammenhängenden Kapazitäten frei. Wir können Ihr Team jedoch gerne ab dem ${formattedAltStart} in folgenden ${bestKombi.apts.length} Unterkünften unterbringen:\n\n`;
    introHtml = `vielen Dank für Ihre Anfrage. Zum direkt angefragten Termin haben wir leider keine zusammenhängenden Kapazitäten frei. Wir können Ihr Team jedoch gerne <b>ab dem ${formattedAltStart}</b> in folgenden ${bestKombi.apts.length} Unterkünften unterbringen:<br><br>`;
  }

  const plainBody = anredeZeile + "\n\n" +
    introPlain +
    plainWohnungen +
    `Kosten: ${formatEur(kopfNetto)} netto pro Person / Nacht (entspricht ${formatEur(summeNachtNetto)} netto pro Nacht für alle Wohnungen gesamt). Für den gesamten Zeitraum also ${nights} Nächte x ${formatEur(summeNachtNetto)} = ${formatEur(mieteNetto)}, Endreinigung nur Einmalig: ${formatEur(summeReinigungNetto)}, Gesamtbetrag: ${formatEur(gesamtBrutto)} brutto (${formatEur(zwischensummeNetto)} netto zzgl. 7% MwSt.). Keine weiteren Kosten. AGB: https://L8Street.com/agb\n\n` +
    cta.plain + footerText;

  const htmlBody = escapeHtml(anredeZeile) + "<br><br>" +
    introHtml +
    htmlWohnungen +
    `<b>Kosten: ${formatEur(kopfNetto)} netto pro Person / Nacht</b> (entspricht ${formatEur(summeNachtNetto)} netto pro Nacht für alle Wohnungen gesamt). Für den gesamten Zeitraum also ${nights} Nächte x ${formatEur(summeNachtNetto)} = ${formatEur(mieteNetto)}, Endreinigung nur Einmalig: ${formatEur(summeReinigungNetto)}, Gesamtbetrag: <b>${formatEur(gesamtBrutto)} brutto</b> (${formatEur(zwischensummeNetto)} netto zzgl. 7% MwSt.). Keine weiteren Kosten. AGB: <a href="https://L8Street.com/agb">https://L8Street.com/agb</a><br><br>` +
    cta.html + footerTextToHtml(footerText);

  return {
    apartmentIds: bestKombi.apts.map(cleanShortCode),
    plainBody,
    htmlBody,
    gesamtBrutto,
    usedStandardPrice
  };
}

// -------------------------------------------------------------
// BASIS-HILFSFUNKTIONEN & GLOBALE UTILITIES
// -------------------------------------------------------------

function getCompanyFooterText() {
  return "\n\nFreundliche Grüße\n\n" +
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
}

function getValidatedCustomerEmail(direktEmail, leadEmail) {
  const candidates = [String(direktEmail || "").trim(), String(leadEmail || "").trim()];
  for (const email of candidates) {
    if (!email || !email.includes("@")) continue;
    if (email.includes("••") || email.includes("nach Klick")) continue;
    const lower = email.toLowerCase();
    if (
      lower.includes("mein-monteurzimmer.de") ||
      lower.includes("mabach.info") ||
      lower.includes("l8street.com") ||
      lower.startsWith("noreply") ||
      lower.startsWith("no-reply")
    ) {
      continue;
    }
    return email;
  }
  return "";
}

function buildSalutation(ansprechpartner, fallbackName) {
  let name = (ansprechpartner && ansprechpartner.trim().length > 2) ? ansprechpartner.trim() : (fallbackName || "").trim();
  name = name.replace(/\s*\(\d+\)\s*/g, " ")
    .replace(/\b\d{4,}\b/g, "")
    .trim();

  if (!name || name === "Damen und Herren" || name.toLowerCase().startsWith("firma")) {
    return "Guten Tag,";
  }
  if (/^herr\s+/i.test(name)) {
    const parts = name.replace(/^herr\s+/i, "").trim().split(/\s+/).filter(Boolean);
    return `Sehr geehrter Herr ${parts[parts.length - 1]},`;
  }
  if (/^frau\s+/i.test(name)) {
    const parts = name.replace(/^frau\s+/i, "").trim().split(/\s+/).filter(Boolean);
    return `Sehr geehrte Frau ${parts[parts.length - 1]},`;
  }
  return `Guten Tag ${name},`;
}

function footerTextToHtml(footer) {
  return "<br><br>" + escapeHtml(footer.trim())
    .replace(/\n/g, "<br>")
    .replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1">$1</a>');
}

function ensureDraftHeaderMMZ(sheet) {
  const cell = sheet.getRange(1, 8);
  if (!cell.getValue()) {
    cell.setValue("Draft Status (JSON)");
    cell.setFontWeight("bold");
    sheet.getRange("H:H").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

function cleanShortCode(val) {
  if (!val) return "";
  return String(val)
    .trim().toLowerCase()
    .replace(/^https?:\/\/a\.l8street\.com\/a\//i, "")
    .replace(/^\/?a\//i, "")
    .replace(/^\//, "")
    .replace(/\/$/, "");
}

function parseIsoDateString(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT", "yyyy-MM-dd");
  }
  const str = String(val).trim();
  const m = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) {
    return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  }
  return "";
}

function extractPureCityName(ortStr) {
  if (!ortStr) return "";
  let clean = String(ortStr).toLowerCase();
  clean = clean.replace(/\b\d{5}\b/g, " ");
  clean = clean.replace(/umgebung.*$/i, " ");
  clean = clean.replace(/monteur24\s*-\s*/i, " ");
  clean = clean.replace(/l8\s*street\s*-\s*/i, " ");
  clean = clean.replace(/top\s*monteurwohnung\s*/i, " ");
  clean = clean.replace(/monteurwohnung\s*/i, " ");
  clean = clean.replace(/mit\s*parkplatz/i, " ");
  clean = clean.replace(/und\s*rhein-neckar/i, " ");
  clean = clean.replace(/[\/\\#,+()$~%.'":*?<>{}\-_]/g, " ").trim();
  const parts = clean.split(/\s+/).filter(w => w.length >= 3);
  return parts.length > 0 ? parts[0] : "";
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

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}