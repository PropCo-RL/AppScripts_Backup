// =================================================================
// STUFE 3: GMAIL DRAFT GENERATOR FÜR TAB "MMZ"
// Liest Spalten A bis J, erzeugt personalisierte Entwürfe
// Schreibt den Draft-Status AUSSCHLIESSLICH in Spalte J (Spalte 10)
// =================================================================

const KONFIGURATION_DRAFT = {
  SHEET_NAME: 'MMZ',
  LABEL_MANUELL: '02_MMZ_Manuell_Portal',
  SPREADSHEET_INVENTORY_ID: '1Aq9doxAkEw435bRCpXafayRwdNeaP9LVMU1YwWG7z70',
  ZEILENHOEHE_PIXEL: 21
};

/**
 * Zentrale Orchestrator-Funktion (Trigger-Einstiegspunkt)
 */
function generateEmailDraftsMMZ() {
  Logger.log("=== START GMAIL DRAFT GENERATOR FÜR MMZ ===");
  const aktiveArbeitsmappe = SpreadsheetApp.getActiveSpreadsheet();
  const leadsTabelle = aktiveArbeitsmappe.getSheetByName(KONFIGURATION_DRAFT.SHEET_NAME);
  if (!leadsTabelle) return;

  const anzahlZeilen = leadsTabelle.getLastRow();
  if (anzahlZeilen <= 1) return;

  let inventarTabelle;
  try {
    const inventarMappe = SpreadsheetApp.openById(KONFIGURATION_DRAFT.SPREADSHEET_INVENTORY_ID);
    inventarTabelle = inventarMappe.getSheetByName('M');
  } catch (e) {
    Logger.log("FEHLER beim Öffnen der Inventar-Tabelle: " + e.toString());
    return;
  }
  if (!inventarTabelle) return;

  stelleDraftHeaderSicher(leadsTabelle);
  const aptDetailsKarte = ladeApartmentInventarKarteSicher(inventarTabelle);
  const firmenFooter = getCompanyFooterText();

  verarbeiteLeadsFuerDrafts(leadsTabelle, anzahlZeilen, aptDetailsKarte, firmenFooter);
  Logger.log("=== GMAIL DRAFT GENERATOR ABGESCHLOSSEN ===");
}

/**
 * Liest Zeilen ein, ermittelt den passenden Entwurf und schreibt Spalte J isoliert zurück
 */
function verarbeiteLeadsFuerDrafts(leadsTabelle, zeilenAnzahl, aptDetailsKarte, firmenFooter) {
  const anzahlDatenZeilen = zeilenAnzahl - 1;
  const datenBereich = leadsTabelle.getRange(2, 1, anzahlDatenZeilen, 10);
  const zeilen = datenBereich.getValues();
  const spalteJWerte = [];
  let aenderungenVorhanden = false;

  for (let index = 0; index < anzahlDatenZeilen; index++) {
    const zeile = zeilen[index];
    const aktuelleZeilennummer = index + 2;
    const spalteJInhalt = String(zeile[9] || "").trim();

    // Bereits bearbeitete Zeilen oder unvollständige Zeilen unverändert beibehalten
    if (!pruefeDraftZeilenVollstaendigkeit(zeile) || spalteJInhalt !== "") {
      spalteJWerte.push([zeile[9]]);
      continue;
    }

    const geparsteDaten = parsiereLeadObjekte(zeile);
    if (!geparsteDaten) {
      spalteJWerte.push([JSON.stringify({ status: "FEHLER_PARSING", message: "JSON unlesbar" })]);
      aenderungenVorhanden = true;
      continue;
    }

    const leadKontext = baueLeadKontext(geparsteDaten, firmenFooter);

    if (!leadKontext.empfaengerEmail) {
      const errEmailObj = handleMissingCustomerEmail(geparsteDaten.leadJson, KONFIGURATION_DRAFT.LABEL_MANUELL);
      spalteJWerte.push([JSON.stringify(errEmailObj)]);
      aenderungenVorhanden = true;
      continue;
    }

    // Warteschleife: Wenn Matching-Engines noch gar nicht gelaufen sind
    if (mussAufEnginesWarten(geparsteDaten.routing, zeile)) {
      spalteJWerte.push([zeile[9]]);
      continue;
    }

    const draftErgebnis = ermittleDraftAktion(zeile, geparsteDaten, leadKontext, aptDetailsKarte);
    if (draftErgebnis) {
      spalteJWerte.push([JSON.stringify(draftErgebnis)]);
      aenderungenVorhanden = true;
      Logger.log(`Zeile ${aktuelleZeilennummer}: Status ${draftErgebnis.status} gesetzt.`);
    } else {
      spalteJWerte.push([zeile[9]]);
    }
  }

  if (aenderungenVorhanden) {
    leadsTabelle.getRange(2, 10, anzahlDatenZeilen, 1).setValues(spalteJWerte);
    leadsTabelle.setRowHeights(2, anzahlDatenZeilen, KONFIGURATION_DRAFT.ZEILENHOEHE_PIXEL);
    Logger.log("=== ENTWURF-STATUS ERFOLGREICH IN SPALTE J GESCHRIEBEN ===");
  }
}

/**
 * Validiert die Mindestanforderungen an eine neue Zeile
 */
function pruefeDraftZeilenVollstaendigkeit(zeile) {
  const cellB = String(zeile[1] || "").trim();
  const cellD = String(zeile[3] || "").trim();
  const cellE = String(zeile[4] || "").trim();
  return cellB && cellB !== "-" && cellD && cellE && !cellE.startsWith("FEHLER");
}

/**
 * Parst JSON aus Spalte B, D und E
 */
function parsiereLeadObjekte(zeile) {
  try {
    return {
      leadJson: JSON.parse(zeile[1]),
      valData: JSON.parse(zeile[3]),
      routingObj: JSON.parse(zeile[4]),
      routing: JSON.parse(zeile[4]).status || ""
    };
  } catch (e) {
    return null;
  }
}

function mussAufEnginesWarten(routing, zeile) {
  const spalteF = String(zeile[5] || "").trim(); // Engine A
  const spalteG = String(zeile[6] || "").trim(); // Engine B
  const spalteH = String(zeile[7] || "").trim(); // Engine A-Alt
  const spalteI = String(zeile[8] || "").trim(); // Engine B-Alt

  if (routing === "STD_SHORT" || routing === "STD_LONG") {
    // 1. Wenn Engine A noch gar nicht lief -> warten
    if (spalteF === "") return true;
    // 2. Wenn Engine A einen Treffer hat, müssen wir nicht auf Alternativen warten
    if (spalteF !== "[]") return false;
    // 3. Engine A war leer ([]): Jetzt müssen Engine A-Alt und Engine B gelaufen sein!
    if (spalteH === "" || spalteG === "") return true;
    // 4. Falls auch Kombi leer war, muss auch B-Alt gelaufen sein
    if (spalteG === "[]" && spalteI === "") return true;
  }

  if (routing === "BIG_SHORT" || routing === "BIG_LONG") {
    // Wenn Engine B noch nicht lief -> warten
    if (spalteG === "") return true;
    // Wenn Engine B leer war ([]), muss auch Engine B-Alt (Spalte I) gelaufen sein
    if (spalteG === "[]" && spalteI === "") return true;
  }

  return false;
}

/**
 * Bereitet E-Mail-Empfänger, Anrede, Nächte und Betreff auf
 */
function baueLeadKontext(geparsteDaten, firmenFooter) {
  const { leadJson, valData, routingObj, routing } = geparsteDaten;
  const spezialData = valData.spezial || {};
  const persData = valData.personen || {};
  const dauerData = valData.dauer || {};

  const empfaengerEmail = getValidatedCustomerEmail(spezialData.direktEmail, leadJson.email);
  const personenAnzahl = Math.max(1, Number(persData.Pers) || Number(leadJson.personen) || 1);

  let naechte = Number(dauerData.tage) || 0;
  if (naechte <= 0 && routingObj.kennzahlen && Number(routingObj.kennzahlen.tage) > 0) {
    naechte = Number(routingObj.kennzahlen.tage);
  }
  if (naechte <= 0 && (routing.includes("LONG") || dauerData.verlaengerung)) {
    naechte = 30;
  }
  naechte = Math.max(1, naechte);

  const kundenNummer = leadJson.kundenNummer || "";
  let mailSubject = "Buchungsanfrage - Monteurwohnung L8 Street";
  if (kundenNummer && kundenNummer !== "N/A") {
    mailSubject += ` (Anfrage ID: ${kundenNummer})`;
  }

  return {
    empfaengerEmail: empfaengerEmail,
    personenAnzahl: personenAnzahl,
    naechte: naechte,
    mailSubject: mailSubject,
    anredeZeile: buildSalutation(spezialData.ansprechpartner, leadJson.gastName),
    forwardMessageId: leadJson.forwardMessageId || "",
    forwardThreadId: leadJson.forwardThreadId || "",
    firmenFooter: firmenFooter
  };
}

/**
 * Entscheidet anhand des Routings und der Kaskade F -> G -> H -> I über den Entwurf
 */
function ermittleDraftAktion(zeile, geparsteDaten, ctx, aptDetailsKarte) {
  const { routing, leadJson } = geparsteDaten;

  // 1. K.O.-Absagen
  if (routing === "REJECT_KO" || routing === "REJECT_1PERS") {
    return erstelleAbsageDraft(ctx, "DRAFT_ABSAGE_KO", routing, buildRejectionContent);
  }

  // 2. Standard-Leads (Kaskade: F -> H -> Fallback auf G -> Fallback auf I -> Absage)
  if (routing === "STD_SHORT" || routing === "STD_LONG") {
    return verarbeiteStandardLeadKaskade(zeile, ctx, aptDetailsKarte, leadJson);
  }

  // 3. Großgruppen (Kaskade: G -> I -> Absage)
  if (routing === "BIG_SHORT" || routing === "BIG_LONG") {
    return verarbeiteKombiLeadKaskade(zeile, ctx, aptDetailsKarte, leadJson);
  }

  // 4. Manuelle Prüfung
  return {
    status: "WAIT_MANUAL_REVIEW",
    routing: routing,
    hinweis: "Kein automatischer Entwurf vorgesehen, erfordert manuelle Prüfung.",
    erfasstAm: new Date().toISOString()
  };
}

/**
 * Behandelt Standard-Leads mit lückenloser Kaskade: Single -> Alt-Single -> Kombi -> Alt-Kombi
 */
function verarbeiteStandardLeadKaskade(zeile, ctx, aptDetailsKarte, leadJson) {
  const singleTreffer = parseArraySicher(zeile[5]);
  if (singleTreffer.length > 0) {
    return erstelleSingleAngebotDraft(singleTreffer[0], null, ctx, aptDetailsKarte, leadJson, false);
  }

  const altSingleTreffer = parseArraySicher(zeile[7]);
  if (altSingleTreffer.length > 0) {
    return erstelleSingleAngebotDraft(altSingleTreffer[0], altSingleTreffer[0], ctx, aptDetailsKarte, leadJson, true);
  }

  // Fallback: Prüfen, ob Engine B (Kombi) oder B-Alt für diesen Standardlead gegriffen hat
  const kombiTreffer = parseArraySicher(zeile[6]);
  if (kombiTreffer.length > 0) {
    return erstelleKombiAngebotDraft(kombiTreffer[0], null, ctx, aptDetailsKarte, leadJson, false);
  }

  const altKombiTreffer = parseArraySicher(zeile[8]);
  if (altKombiTreffer.length > 0) {
    return erstelleKombiAngebotDraft(altKombiTreffer[0], altKombiTreffer[0], ctx, aptDetailsKarte, leadJson, true);
  }

  return erstelleAbsageDraft(ctx, "DRAFT_KEINE_VERFUEGBARKEIT", null, buildNoAvailabilityContent);
}

/**
 * Behandelt Großgruppen-Leads: Kombi -> Alt-Kombi -> Absage
 */
function verarbeiteKombiLeadKaskade(zeile, ctx, aptDetailsKarte, leadJson) {
  const kombiTreffer = parseArraySicher(zeile[6]);
  if (kombiTreffer.length > 0) {
    return erstelleKombiAngebotDraft(kombiTreffer[0], null, ctx, aptDetailsKarte, leadJson, false);
  }

  const altKombiTreffer = parseArraySicher(zeile[8]);
  if (altKombiTreffer.length > 0) {
    return erstelleKombiAngebotDraft(altKombiTreffer[0], altKombiTreffer[0], ctx, aptDetailsKarte, leadJson, true);
  }

  return erstelleAbsageDraft(ctx, "DRAFT_KEINE_VERFUEGBARKEIT", null, buildNoAvailabilityContent);
}

/**
 * Baut und sendet Single-Apartment Entwürfe
 */
function erstelleSingleAngebotDraft(angebot, altMeta, ctx, aptDetailsKarte, leadJson, istAlternative) {
  const offer = buildSingleOfferContent(angebot, aptDetailsKarte, ctx.personenAnzahl, ctx.naechte, ctx.anredeZeile, ctx.firmenFooter, leadJson, altMeta);
  const { draftId, inThread } = createOrReplyDraft(ctx.empfaengerEmail, ctx.mailSubject, offer.plainBody, offer.htmlBody, ctx.forwardMessageId, ctx.forwardThreadId);

  return {
    status: istAlternative ? "DRAFT_ANGEBOT_SINGLE_ALTERNATIVE" : "DRAFT_ANGEBOT_SINGLE",
    apartmentId: offer.cleanId,
    tageVerschoben: altMeta ? altMeta.tageVerschoben : 0,
    altStart: altMeta ? altMeta.altStart : null,
    standardPreisVerwendet: offer.usedStandardPrice,
    draftId: draftId || "",
    to: ctx.empfaengerEmail,
    subject: ctx.mailSubject,
    gesamtBrutto: offer.gesamtBrutto,
    imThread: inThread,
    erstelltAm: new Date().toISOString()
  };
}

/**
 * Baut und sendet Kombi-Apartment Entwürfe
 */
function erstelleKombiAngebotDraft(kombiAngebot, altMeta, ctx, aptDetailsKarte, leadJson, istAlternative) {
  const offer = buildKombiOfferContent(kombiAngebot, aptDetailsKarte, ctx.personenAnzahl, ctx.naechte, ctx.anredeZeile, ctx.firmenFooter, leadJson, altMeta);
  const { draftId, inThread } = createOrReplyDraft(ctx.empfaengerEmail, ctx.mailSubject, offer.plainBody, offer.htmlBody, ctx.forwardMessageId, ctx.forwardThreadId);

  return {
    status: istAlternative ? "DRAFT_ANGEBOT_KOMBI_ALTERNATIVE" : "DRAFT_ANGEBOT_KOMBI",
    apartmentIds: offer.apartmentIds,
    tageVerschoben: altMeta ? altMeta.tageVerschoben : 0,
    altStart: altMeta ? altMeta.altStart : null,
    standardPreisVerwendet: offer.usedStandardPrice,
    draftId: draftId || "",
    to: ctx.empfaengerEmail,
    subject: ctx.mailSubject,
    gesamtBrutto: offer.gesamtBrutto,
    imThread: inThread,
    erstelltAm: new Date().toISOString()
  };
}

/**
 * Baut und sendet Absage-Entwürfe
 */
function erstelleAbsageDraft(ctx, statusWert, routing, generatorFn) {
  const { plainBody, htmlBody } = generatorFn(ctx.anredeZeile, ctx.firmenFooter);
  const { draftId, inThread } = createOrReplyDraft(ctx.empfaengerEmail, ctx.mailSubject, plainBody, htmlBody, ctx.forwardMessageId, ctx.forwardThreadId);

  return {
    status: statusWert,
    draftId: draftId || "",
    to: ctx.empfaengerEmail,
    subject: ctx.mailSubject,
    routing: routing || undefined,
    imThread: inThread,
    erstelltAm: new Date().toISOString()
  };
}

/**
 * Lädt Inventardaten aus Tab 'M' mit Spalte AJ Ausschluss
 */
function ladeApartmentInventarKarteSicher(tabelle) {
  if (typeof loadApartmentInventoryMap === "function") {
    return loadApartmentInventoryMap(tabelle);
  }
  const rohWerte = tabelle.getDataRange().getValues();
  const inventar = {};
  for (let r = 1; r < rohWerte.length; r++) {
    const name = String(rohWerte[r][0] || "").trim();
    if (!name) continue;
    // Spalte AJ Ausschluss prüfen
    if (String(rohWerte[r][35] || "").trim().toLowerCase() === "x") continue;
    inventar[name.toLowerCase()] = rohWerte[r];
  }
  return inventar;
}

/**
 * Header für Spalte J sicherstellen
 */
function stelleDraftHeaderSicher(tabelle) {
  const kopfZelle = tabelle.getRange(1, 10);
  if (!kopfZelle.getValue() || kopfZelle.getValue() !== "Draft Status (JSON)") {
    kopfZelle.setValue("Draft Status (JSON)");
    kopfZelle.setFontWeight("bold");
    tabelle.getRange("J:J").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

/**
 * Hilfsfunktion zum sicheren Parsen von Treffer-Arrays
 */
function parseArraySicher(wert) {
  try {
    const parsed = JSON.parse(String(wert || "[]"));
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}