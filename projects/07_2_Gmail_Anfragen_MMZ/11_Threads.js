// =================================================================
// 11_THREAD SYNC (MMZ & THREAD ARCHIVE) - MODULAR ARCHITECTURE
// - Orchestrator synchronisiert E-Mails zum 'Thread Archive'
// - Schreibt JEDE Nachricht zeilenweise ins Archiv (Höhe: 21, CLIP)
// - Übergibt in MMZ Spalte P gezielt den letzten Kunden-Inbound ('in')
// - Duplikatsperre über Gmail Message-ID
// =================================================================

const THREAD_SYNC_KONFIGURATION = {
  TABELLE_HAUPT: 'MMZ',
  TABELLE_ARCHIV: 'Thread Archive',
  SPALTE_LEAD: 2,       // B: Lead JSON
  SPALTE_DRAFT: 10,     // J: Draft Status JSON
  SPALTE_SIGNAL: 16     // P: Letzter Inbound Signal
};

// -----------------------------------------------------------------
// 1. ZENTRALER ORCHESTRATOR
// -----------------------------------------------------------------
function syncEmailThreadsMMZ() {
  const arbeitsmappe = SpreadsheetApp.getActiveSpreadsheet();
  const hauptBlatt = arbeitsmappe.getSheetByName(THREAD_SYNC_KONFIGURATION.TABELLE_HAUPT);
  if (!hauptBlatt) return;

  const anzahlZeilen = hauptBlatt.getLastRow() - 1;
  if (anzahlZeilen <= 0) return;

  const archivBlatt = sicherstelleArchivBlatt(arbeitsmappe);
  sicherstelleSignalHeader(hauptBlatt);

  const batchDaten = ladeHauptblattSyncDaten(hauptBlatt, anzahlZeilen);
  const bekannteNachrichtIds = ladeBekannteNachrichtIds(archivBlatt);

  const verarbeitungsErgebnis = verarbeiteThreadsBatch(batchDaten, bekannteNachrichtIds);

  if (verarbeitungsErgebnis.archivZeilen.length > 0) {
    haengeZeilenAnArchivAn(archivBlatt, verarbeitungsErgebnis.archivZeilen);
  }

  if (verarbeitungsErgebnis.signalAktualisiert) {
    schreibeSignalBatchZurueck(hauptBlatt, verarbeitungsErgebnis.signalWerte, anzahlZeilen);
  }
}

// -----------------------------------------------------------------
// 2. SHEET-IO & BATCHING
// -----------------------------------------------------------------
function ladeHauptblattSyncDaten(hauptBlatt, anzahlZeilen) {
  return {
    leads: hauptBlatt.getRange(2, THREAD_SYNC_KONFIGURATION.SPALTE_LEAD, anzahlZeilen, 1).getValues(),
    drafts: hauptBlatt.getRange(2, THREAD_SYNC_KONFIGURATION.SPALTE_DRAFT, anzahlZeilen, 1).getValues(),
    signale: hauptBlatt.getRange(2, THREAD_SYNC_KONFIGURATION.SPALTE_SIGNAL, anzahlZeilen, 1).getValues()
  };
}

function haengeZeilenAnArchivAn(archivBlatt, neueZeilen) {
  const startZeile = archivBlatt.getLastRow() + 1;
  const anzahl = neueZeilen.length;

  archivBlatt.getRange(startZeile, 1, anzahl, 8).setValues(neueZeilen);
  archivBlatt.getRange(startZeile, 1, anzahl, 8).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  archivBlatt.setRowHeights(startZeile, anzahl, 21);

  Logger.log(`${anzahl} Nachricht(en) ins 'Thread Archive' weggeschrieben.`);
}

function schreibeSignalBatchZurueck(hauptBlatt, signalWerte, anzahlZeilen) {
  hauptBlatt.getRange(2, THREAD_SYNC_KONFIGURATION.SPALTE_SIGNAL, anzahlZeilen, 1).setValues(signalWerte);
  hauptBlatt.setRowHeights(2, anzahlZeilen, 21);
  SpreadsheetApp.flush();
  Logger.log("MMZ Spalte P erfolgreich aktualisiert.");
}

function sicherstelleArchivBlatt(arbeitsmappe) {
  let blatt = arbeitsmappe.getSheetByName(THREAD_SYNC_KONFIGURATION.TABELLE_ARCHIV);
  if (!blatt) {
    blatt = arbeitsmappe.insertSheet(THREAD_SYNC_KONFIGURATION.TABELLE_ARCHIV);
    const spaltenKöpfe = [
      "Zeit", "InOut", "Sender / Empfänger", "Art",
      "Nachricht", "Nachricht ID", "Nachricht JSON", "State JSON"
    ];
    blatt.getRange(1, 1, 1, spaltenKöpfe.length).setValues([spaltenKöpfe]).setFontWeight("bold");
    blatt.setFrozenRows(1);
    blatt.getRange("A:H").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  return blatt;
}

function sicherstelleSignalHeader(hauptBlatt) {
  const kopfZelle = hauptBlatt.getRange(1, THREAD_SYNC_KONFIGURATION.SPALTE_SIGNAL);
  if (!kopfZelle.getValue() || kopfZelle.getValue() !== "Letzter Inbound (Signal)") {
    kopfZelle.setValue("Letzter Inbound (Signal)").setFontWeight("bold");
    hauptBlatt.getRange(1, THREAD_SYNC_KONFIGURATION.SPALTE_SIGNAL, hauptBlatt.getMaxRows(), 1)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

// -----------------------------------------------------------------
// 3. THREAD-VERARBEITUNG & PARSING
// -----------------------------------------------------------------
function verarbeiteThreadsBatch(batchDaten, bekannteIds) {
  const anzahl = batchDaten.leads.length;
  const archivZeilen = [];
  const signalWerte = batchDaten.signale.slice();
  let signalAktualisiert = false;

  for (let i = 0; i < anzahl; i++) {
    const leadText = String(batchDaten.leads[i][0] || "").trim();
    const draftText = String(batchDaten.drafts[i][0] || "").trim();
    if (!leadText || leadText === "-") continue;

    const leadInfo = parseLeadUndDraft(leadText, draftText);
    if (!leadInfo.gastEmail && !leadInfo.draftId) continue;

    const thread = sucheGmailThread(leadInfo.draftId, leadInfo.gastEmail, leadInfo.erfasstAm);
    if (!thread) continue;

    const nachrichtenErgebnis = filtereUndFormatiereNachrichten(thread.getMessages(), bekannteIds, leadInfo.gastEmail);
    nachrichtenErgebnis.neueArchivEintraege.forEach(zeile => archivZeilen.push(zeile));

    if (nachrichtenErgebnis.letzterKundenInbound) {
      signalWerte[i][0] = JSON.stringify({
        letzteMsgId: nachrichtenErgebnis.letzterKundenInbound.id,
        t: "in",
        date: nachrichtenErgebnis.letzterKundenInbound.date,
        textSnippet: nachrichtenErgebnis.letzterKundenInbound.text.substring(0, 300),
        gastEmail: leadInfo.gastEmail
      });
      signalAktualisiert = true;
    }
  }

  return { archivZeilen: archivZeilen, signalWerte: signalWerte, signalAktualisiert: signalAktualisiert };
}

function filtereUndFormatiereNachrichten(nachrichten, bekannteIds, gastEmail) {
  const neueArchivEintraege = [];
  let letzterKundenInbound = null;

  for (const msg of nachrichten) {
    const nachrichtId = msg.getId();

    if (bekannteIds.has(nachrichtId)) {
      const istOut = String(msg.getFrom() || "").toLowerCase().includes("l8street.com");
      if (!istOut) {
        letzterKundenInbound = { id: nachrichtId, date: msg.getDate().toISOString(), text: msg.getPlainBody() || "" };
      }
      continue;
    }

    if (istSystemOderSpamMail(msg)) continue;

    const absender = String(msg.getFrom() || "").toLowerCase();
    const istOutbound = absender.includes("l8street.com");
    const inOut = istOutbound ? "out" : "in";
    const bereinigterText = bereinigeNachrichtenText(msg.getPlainBody() || "", inOut);
    if (!bereinigterText) continue;

    const einzeiligerText = bereinigterText.replace(/[\r\n]+/g, " ").trim();
    const zeitpunktIso = msg.getDate().toISOString();
    const kontaktKey = gastEmail || absender;

    const roheNutzdaten = { from: msg.getFrom(), to: msg.getTo(), subject: msg.getSubject(), date: zeitpunktIso };

    neueArchivEintraege.push([
      zeitpunktIso, inOut, kontaktKey, "text", einzeiligerText,
      nachrichtId, JSON.stringify(roheNutzdaten), ""
    ]);

    bekannteIds.add(nachrichtId);

    if (inOut === "in") {
      letzterKundenInbound = { id: nachrichtId, date: zeitpunktIso, text: einzeiligerText };
    }
  }

  return { neueArchivEintraege: neueArchivEintraege, letzterKundenInbound: letzterKundenInbound };
}

function parseLeadUndDraft(leadText, draftText) {
  let leadDaten = {}, draftDaten = {};
  try {
    leadDaten = JSON.parse(leadText);
    if (draftText && !draftText.startsWith("Fehler")) {
      draftDaten = JSON.parse(draftText);
    }
  } catch (e) {}

  return {
    gastEmail: String(leadDaten.email || draftDaten.to || "").toLowerCase().trim(),
    draftId: draftDaten.draftId || null,
    erfasstAm: leadDaten.erfasstAm || null
  };
}

function sucheGmailThread(draftId, gastEmail, erfasstAmText) {
  if (draftId) {
    try {
      const draft = GmailApp.getDraft(draftId);
      if (draft) return draft.getMessage().getThread();
    } catch (e) {}
  }
  if (!gastEmail) return null;

  try {
    let bezugsDatum = erfasstAmText ? new Date(erfasstAmText) : new Date(Date.now() - (30 * 86400000));
    if (isNaN(bezugsDatum.getTime())) bezugsDatum = new Date(Date.now() - (30 * 86400000));
    const sekundenSeitEpoch = Math.floor(bezugsDatum.getTime() / 1000);
    const suchanfrage = `(to:${gastEmail} OR from:${gastEmail}) after:${sekundenSeitEpoch}`;
    const threads = GmailApp.search(suchanfrage, 0, 1);
    if (threads && threads.length > 0) return threads[0];
  } catch (e) {}
  return null;
}

function ladeBekannteNachrichtIds(archivBlatt) {
  const letzteZeile = archivBlatt.getLastRow();
  if (letzteZeile <= 1) return new Set();

  const idWerte = archivBlatt.getRange(2, 6, letzteZeile - 1, 1).getValues();
  const idMenge = new Set();
  for (let r = 0; r < idWerte.length; r++) {
    const eintrag = String(idWerte[r][0] || "").trim();
    if (eintrag) idMenge.add(eintrag);
  }
  return idMenge;
}

// -----------------------------------------------------------------
// 4. FILTERUNG & TEXTBEREINIGUNG
// -----------------------------------------------------------------
function istSystemOderSpamMail(nachricht) {
  const absender = String(nachricht.getFrom() || "").toLowerCase();
  const empfaenger = String(nachricht.getTo() || "").toLowerCase();
  const betreff = String(nachricht.getSubject() || "").toLowerCase();

  const domainAusschluss = [
    "lodgify.com", "no-reply", "noreply", "support@wix-forms.com",
    "wix-forms.com", "sendgrid.net", "twilio.com", "zendesk.com",
    "mein-monteurzimmer.de"
  ];
  for (const domain of domainAusschluss) {
    if (absender.includes(domain)) return true;
  }
  if (empfaenger.includes("cc@l8street.com")) return true;

  const betreffAusschluss = [
    "automatische antwort", "out of office", "abwesenheitsnotiz",
    "automatic reply", "your report has been generated"
  ];
  for (const begriff of betreffAusschluss) {
    if (betreff.includes(begriff)) return true;
  }
  return false;
}

function bereinigeNachrichtenText(roherText, inOut) {
  if (!roherText) return "";
  let text = roherText;

  const zitatKopfzeilen = [
    /[-*_\s]*Ursprüngliche Nachricht[-*_\s]*/i,
    /[-*_\s]*Original Message[-*_\s]*/i,
    /Am\s+.+schrieb\s+.+:/i,
    /On\s+.+wrote:/i
  ];
  for (const regex of zitatKopfzeilen) {
    const teile = text.split(regex);
    if (teile.length > 1 && teile[0].trim().length > 10) {
      text = teile[0];
    }
  }

  text = text.split('\n')
    .filter(zeile => !zeile.trim().startsWith('>'))
    .join('\n');

  const appSignaturen = [
    /Gesendet mit der GMX Mail App/i,
    /Von meinem iPhone gesendet/i,
    /Von meinem Mobiltelefon gesendet/i,
    /Get Outlook for/i
  ];
  for (const fRegex of appSignaturen) {
    const fTeile = text.split(fRegex);
    if (fTeile.length > 1) text = fTeile[0];
  }

  if (inOut === "out") {
    const firmenSignaturen = [
      /Freundliche Grüße/i, /Mit freundlichen Grüßen/i, /Besten Gruß/i,
      /Viele Grüße/i, /©\s*2017-\d{4}\s*L8\s*Street\s*GmbH/i,
      /Geschäftsführer:\s*Raul\s*Leneweit/i, /Amtsgericht\s*Mannheim\s*HRB/i
    ];
    for (const fRegex of firmenSignaturen) {
      const fTeile = text.split(fRegex);
      if (fTeile.length > 1) text = fTeile[0];
    }
  }

  return text.trim();
}