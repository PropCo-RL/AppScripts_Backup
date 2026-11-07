// =================================================================
// STUFE 3: GMAIL DRAFT GENERATOR FÜR TAB "MMZ"
// Haupt-Ablaufsteuerung & Google Sheets I/O
// Schreibt den Draft-Status in Spalte J (Spalte 10)
// =================================================================
function generateEmailDraftsMMZ() {
  const SHEET_NAME = 'MMZ';
  const LABEL_MANUELL = '02_MMZ_Manuell_Portal';
  const SPREADSHEET_INVENTORY_ID = '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ss.getSheetByName(SHEET_NAME);
  let invSheetM;

  try {
    const invSS = SpreadsheetApp.openById(SPREADSHEET_INVENTORY_ID);
    invSheetM = invSS.getSheetByName('M');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Inventar-Tabelle: " + e.toString());
    return;
  }

  if (!leadsSheet || !invSheetM) {
    Logger.log("FEHLER: Tabellenblatt MMZ oder M fehlt.");
    return;
  }

  const lastRow = leadsSheet.getLastRow();
  if (lastRow <= 1) return;

  // Header für Spalte J (Spalte 10)
  ensureDraftHeaderMMZ(leadsSheet);

  const aptDetailsMap = loadApartmentInventoryMap(invSheetM);

  // 10 Spalten einlesen: A bis J (Index 0 bis 9)
  const numRows = lastRow - 1;
  const dataRange = leadsSheet.getRange(2, 1, numRows, 10);
  const rows = dataRange.getValues();
  const outputColJ = [];
  let updatesMade = false;
  const footerText = getCompanyFooterText();

  for (let i = 0; i < rows.length; i++) {
    const cellB = String(rows[i][1] || "").trim(); // Lead ID / Raw
    const cellD = String(rows[i][3] || "").trim(); // Data JSON
    const cellE = String(rows[i][4] || "").trim(); // Routing JSON
    const cellF = String(rows[i][5] || "").trim(); // Engine A Single
    const cellG = String(rows[i][6] || "").trim(); // Engine B Kombi
    const cellH = String(rows[i][7] || "").trim(); // Alternative Single (<6P)
    const cellI = String(rows[i][8] || "").trim(); // Alternative Kombi (>6P)
    const cellJ = String(rows[i][9] || "").trim(); // Zielspalte: Draft Status

    // Zeile überspringen, wenn unvollständig oder bereits bearbeitet
    if (!cellB || cellB === "-" || !cellD || !cellE || cellJ !== "") {
      outputColJ.push([rows[i][9]]);
      continue;
    }

    let leadJson = {}, valData = {}, routingObj = {};
    try {
      leadJson = JSON.parse(cellB);
      valData = JSON.parse(cellD);
      routingObj = JSON.parse(cellE);
    } catch (e) {
      outputColJ.push([JSON.stringify({ status: "FEHLER_PARSING", message: "JSON unlesbar" })]);
      updatesMade = true;
      continue;
    }

    const routing = routingObj.status || "";
    const spezialData = valData.spezial || {};

    const recipientEmail = getValidatedCustomerEmail(spezialData.direktEmail, leadJson.email);
    if (!recipientEmail) {
      const errEmailObj = handleMissingCustomerEmail(leadJson, LABEL_MANUELL);
      outputColJ.push([JSON.stringify(errEmailObj)]);
      updatesMade = true;
      continue;
    }

    const persData = valData.personen || {};
    const dauerData = valData.dauer || {};
    const persCount = Math.max(1, Number(persData.Pers) || Number(leadJson.personen) || 1);

    let parsedNights = Number(dauerData.tage) || 0;
    if (parsedNights <= 0 && routingObj.kennzahlen && Number(routingObj.kennzahlen.tage) > 0) {
      parsedNights = Number(routingObj.kennzahlen.tage);
    }
    if (parsedNights <= 0 && (routing.includes("LONG") || dauerData.verlaengerung)) {
      parsedNights = 30;
    }
    const nights = Math.max(1, parsedNights);

    const forwardMessageId = leadJson.forwardMessageId || "";
    const forwardThreadId = leadJson.forwardThreadId || "";
    const anredeZeile = buildSalutation(spezialData.ansprechpartner, leadJson.gastName);
    const kundenNummer = leadJson.kundenNummer || "";

    let mailSubject = "Buchungsanfrage - Monteurwohnung L8 Street";
    if (kundenNummer && kundenNummer !== "N/A") {
      mailSubject += " (Anfrage ID: " + kundenNummer + ")";
    }

    let draftResult = null;

    // FALL 1: K.O. ABSAGE
    if (routing === "REJECT_KO" || routing === "REJECT_1PERS") {
      const { plainBody, htmlBody } = buildRejectionContent(anredeZeile, footerText);
      const { draftId, inThread } = createOrReplyDraft(recipientEmail, mailSubject, plainBody, htmlBody, forwardMessageId, forwardThreadId);
      draftResult = {
        status: "DRAFT_ABSAGE_KO",
        draftId: draftId || "",
        to: recipientEmail,
        subject: mailSubject,
        routing: routing,
        imThread: inThread,
        erstelltAm: new Date().toISOString()
      };
    }

    // FALL 2: SINGLE APARTMENT
    else if (routing === "STD_SHORT" || routing === "STD_LONG") {
      let singleList = [];
      try { singleList = JSON.parse(cellF || "[]"); } catch (e) {}

      // Prüfen, ob Standard-Treffer vorliegt
      if (singleList.length > 0) {
        const best = singleList[0];
        const offer = buildSingleOfferContent(best, aptDetailsMap, persCount, nights, anredeZeile, footerText, leadJson, null);
        const { draftId, inThread } = createOrReplyDraft(recipientEmail, mailSubject, offer.plainBody, offer.htmlBody, forwardMessageId, forwardThreadId);
        draftResult = {
          status: "DRAFT_ANGEBOT_SINGLE",
          apartmentId: offer.cleanId,
          standardPreisVerwendet: offer.usedStandardPrice,
          draftId: draftId || "",
          to: recipientEmail,
          subject: mailSubject,
          gesamtBrutto: offer.gesamtBrutto,
          imThread: inThread,
          erstelltAm: new Date().toISOString()
        };
      } else {
        // FALLBACK: Wenn KEIN Standardangebot vorhanden, prüfe Spalte H (Alternative <6P)
        let altSingleList = [];
        try { altSingleList = JSON.parse(cellH || "[]"); } catch (e) {}

        if (altSingleList.length > 0) {
          const bestAlt = altSingleList[0];
          const offer = buildSingleOfferContent(bestAlt, aptDetailsMap, persCount, nights, anredeZeile, footerText, leadJson, bestAlt);
          const { draftId, inThread } = createOrReplyDraft(recipientEmail, mailSubject, offer.plainBody, offer.htmlBody, forwardMessageId, forwardThreadId);
          draftResult = {
            status: "DRAFT_ANGEBOT_SINGLE_ALTERNATIVE",
            apartmentId: offer.cleanId,
            tageVerschoben: bestAlt.tageVerschoben,
            altStart: bestAlt.altStart,
            standardPreisVerwendet: offer.usedStandardPrice,
            draftId: draftId || "",
            to: recipientEmail,
            subject: mailSubject,
            gesamtBrutto: offer.gesamtBrutto,
            imThread: inThread,
            erstelltAm: new Date().toISOString()
          };
        } else {
          // Weder Standard noch Alternative vorhanden -> Absage keine Verfügbarkeit
          const { plainBody, htmlBody } = buildNoAvailabilityContent(anredeZeile, footerText);
          const { draftId, inThread } = createOrReplyDraft(recipientEmail, mailSubject, plainBody, htmlBody, forwardMessageId, forwardThreadId);
          draftResult = {
            status: "DRAFT_KEINE_VERFUEGBARKEIT",
            draftId: draftId || "",
            to: recipientEmail,
            subject: mailSubject,
            imThread: inThread,
            erstelltAm: new Date().toISOString()
          };
        }
      }
    }

    // FALL 3: KOMBI-ANGEBOT
    else if (routing === "BIG_SHORT" || routing === "BIG_LONG") {
      let kombiList = [];
      try { kombiList = JSON.parse(cellG || "[]"); } catch (e) {}

      if (kombiList.length > 0) {
        const bestKombi = kombiList[0];
        const offer = buildKombiOfferContent(bestKombi, aptDetailsMap, persCount, nights, anredeZeile, footerText, leadJson, null);
        const { draftId, inThread } = createOrReplyDraft(recipientEmail, mailSubject, offer.plainBody, offer.htmlBody, forwardMessageId, forwardThreadId);
        draftResult = {
          status: "DRAFT_ANGEBOT_KOMBI",
          apartmentIds: offer.apartmentIds,
          standardPreisVerwendet: offer.usedStandardPrice,
          draftId: draftId || "",
          to: recipientEmail,
          subject: mailSubject,
          gesamtBrutto: offer.gesamtBrutto,
          imThread: inThread,
          erstelltAm: new Date().toISOString()
        };
      } else {
        // FALLBACK: Wenn KEIN Standard-Kombiangebot vorhanden, prüfe Spalte I (Alternative Kombi)
        let altKombiList = [];
        try { altKombiList = JSON.parse(cellI || "[]"); } catch (e) {}

        if (altKombiList.length > 0) {
          const bestAltKombi = altKombiList[0];
          const offer = buildKombiOfferContent(bestAltKombi, aptDetailsMap, persCount, nights, anredeZeile, footerText, leadJson, bestAltKombi);
          const { draftId, inThread } = createOrReplyDraft(recipientEmail, mailSubject, offer.plainBody, offer.htmlBody, forwardMessageId, forwardThreadId);
          draftResult = {
            status: "DRAFT_ANGEBOT_KOMBI_ALTERNATIVE",
            apartmentIds: offer.apartmentIds,
            tageVerschoben: bestAltKombi.tageVerschoben,
            altStart: bestAltKombi.altStart,
            standardPreisVerwendet: offer.usedStandardPrice,
            draftId: draftId || "",
            to: recipientEmail,
            subject: mailSubject,
            gesamtBrutto: offer.gesamtBrutto,
            imThread: inThread,
            erstelltAm: new Date().toISOString()
          };
        } else {
          const { plainBody, htmlBody } = buildNoAvailabilityContent(anredeZeile, footerText);
          const { draftId, inThread } = createOrReplyDraft(recipientEmail, mailSubject, plainBody, htmlBody, forwardMessageId, forwardThreadId);
          draftResult = {
            status: "DRAFT_KEINE_VERFUEGBARKEIT",
            draftId: draftId || "",
            to: recipientEmail,
            subject: mailSubject,
            imThread: inThread,
            erstelltAm: new Date().toISOString()
          };
        }
      }
    }

    // FALL 4: MANUELLE PRÜFUNG
    else {
      draftResult = {
        status: "WAIT_MANUAL_REVIEW",
        routing: routing,
        hinweis: "Kein automatischer Entwurf vorgesehen, erfordert manuelle Prüfung.",
        erfasstAm: new Date().toISOString()
      };
    }

    if (draftResult) {
      outputColJ.push([JSON.stringify(draftResult)]);
      updatesMade = true;
    } else {
      outputColJ.push([rows[i][9]]);
    }
  }

  // Zurückschreiben in Spalte J (Spalte 10)
  if (updatesMade) {
    leadsSheet.getRange(2, 10, numRows, 1).setValues(outputColJ);
    leadsSheet.setRowHeights(2, numRows, 21);
    Logger.log("=== ENTWURF-STATUS ERFOLGREICH IN SPALTE J GESCHRIEBEN ===");
  }
}

function ensureDraftHeaderMMZ(sheet) {
  const cell = sheet.getRange(1, 10); // Spalte J
  if (!cell.getValue() || cell.getValue() !== "Draft Status (JSON)") {
    cell.setValue("Draft Status (JSON)");
    cell.setFontWeight("bold");
    sheet.getRange("J:J").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}