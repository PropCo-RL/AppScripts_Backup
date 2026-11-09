/**
 * 01_ImportUndFilter.gs
 * Lädt Gmail-Rechnungen, gleicht Absender gegen die neuesten 200 Kunden-Mails aus dem 
 * Buchungs-Sheet (Tab RE, Spalte I) ab, blockiert Kunden-Zahlungsbelege, 
 * dedupliziert Caya-Eingänge und speichert nur echte Rechnungs-PDFs in Drive & Sheet.
 */
function importiereUndFiltereRechnungen() {
  Logger.log("=== START: IMPORT & DIREKTFILTERUNG ===");

  const konfiguration = {
    suchLabel: "support@-rechnung@",
    folgeLabel: "support@-rechnung@/weitergeleitetAnLx",
    driveOrdnerId: "1j8xPlrOZTFP205HG9EPpGxiQYV-UI6rm",
    tabellenName: "API Mail Eingang",
    eigeneDomain: "l8street.com",
    minKb: 15,
    buchungsSpreadsheetId: "1mMEVsVFIx5VxIV3EyJYlba_G0fgEX9t0tR2KbFkm-nU",
    buchungsTabName: "RE"
  };

  const gesperrteAbsender = [
    "sirato31215@gmail.com",
    "chanleneweit@gmail.com",
    "inkasso@pforzheim.creditreform.de"
  ];

  let folgeLabelObj = GmailApp.getUserLabelByName(konfiguration.folgeLabel) || GmailApp.createLabel(konfiguration.folgeLabel);
  const zielOrdner = DriveApp.getFolderById(konfiguration.driveOrdnerId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tabellenBlatt = ss.getSheetByName(konfiguration.tabellenName) || ss.insertSheet(konfiguration.tabellenName);

  // 1. Tabellen-Header initialisieren (Spalte 5 und 6 ohne Versatz)
  initialisiereSheetHeader(tabellenBlatt);

  // 2. Kunden-E-Mails aus den letzten 200 Zeilen des Buchungs-Sheets laden
  const kundenEmailSet = ladeKundenEmailSet(konfiguration.buchungsSpreadsheetId, konfiguration.buchungsTabName);

  // 3. Vorhandene Keys (inkl. Caya-Normalisierung) einlesen
  const vorhandeneSchluessel = ladeVorhandeneSchluessel(tabellenBlatt);

  // 4. Gmail Threads abfragen
  const suchQuery = `label:${konfiguration.suchLabel} -label:"${konfiguration.folgeLabel}" has:attachment`;
  const emailThreads = GmailApp.search(suchQuery, 0, 50);
  const neueZeilen = [];

  for (let i = 0; i < emailThreads.length; i++) {
    const nachrichten = emailThreads[i].getMessages();

    for (let j = 0; j < nachrichten.length; j++) {
      const nachricht = nachrichten[j];
      const rawSender = nachricht.getFrom();
      const rawReplyTo = nachricht.getReplyTo();
      const cleanSender = bereinigeEmail(rawSender);
      const cleanReplyTo = bereinigeEmail(rawReplyTo);
      const betreff = nachricht.getSubject().toLowerCase();
      const datumIso = formatiereIsoDatum(nachricht.getDate());

      // Empfänger-Strings direkt prüfen (fängt Mehrfachempfänger zuverlässig ab)
      const toEmpfaengerRaw = String(nachricht.getTo() || "").toLowerCase();
      const ccEmpfaengerRaw = String(nachricht.getCc() || "").toLowerCase();

      // Schutz 1: Eigene Domain & explizit gesperrte Absender
      if (cleanSender.includes(konfiguration.eigeneDomain) || gesperrteAbsender.includes(cleanSender)) continue;
      if (/^(accepted:|declined:|tentative:|invitation:|aktualisiert:|abgesagt:)/i.test(betreff)) continue;

      // Schutz 2: System- & reine Info-Mails aussortieren
      const istRechnungsDienst = cleanSender.includes("lexware.de") || cleanSender.includes("sevdesk") || cleanSender.includes("fastbill");
      const istSystemReport = cleanSender.includes("dmarc") || cleanSender.includes("noreply") || betreff.includes("report domain");
      
      const hatInfoEmpfaenger = toEmpfaengerRaw.includes("info@") || ccEmpfaengerRaw.includes("info@");
      const hatRechnungEmpfaenger = toEmpfaengerRaw.includes("rechnung@") || ccEmpfaengerRaw.includes("rechnung@");
      const istNurInfo = hatInfoEmpfaenger && !hatRechnungEmpfaenger;

      if (!istRechnungsDienst && (istSystemReport || istNurInfo)) continue;

      const effectiveSender = (cleanReplyTo.includes("@") && istRechnungsDienst) ? cleanReplyTo : cleanSender;

      // Schutz 3: Kunden-Zahlungsbelege aus Buchungs-Sheet ignorieren
      if (kundenEmailSet.has(cleanSender) || kundenEmailSet.has(effectiveSender)) {
        Logger.log(`[KUNDE ERKANNT] Zahlungsbeleg von Buchungskunde ignoriert: ${effectiveSender}`);
        continue;
      }

      const anhaenge = nachricht.getAttachments();
      for (let k = 0; k < anhaenge.length; k++) {
        const anhang = anhaenge[k];
        const dateiName = anhang.getName().toLowerCase();
        const mimeTyp = anhang.getContentType().toLowerCase();
        const dateiGroesseKb = anhang.getSize() / 1024;

        // Schutz 4: Strikt nur PDFs ab 15 KB (Videos, MP4, Bilder fliegen direkt raus)
        if (dateiGroesseKb < konfiguration.minKb) continue;
        if (!istZulaessigesPdf(dateiName, mimeTyp)) continue;

        // Schutz 5: Eindeutiger Duplikat-Schlüssel mit Caya-Sonderbehandlung
        const istCaya = cleanSender.includes("getcaya.com");
        const absenderKey = istCaya ? "caya" : effectiveSender;
        const eindeutigerKey = `${absenderKey}|${datumIso}|${dateiName}|${Math.round(dateiGroesseKb)}`;
        const fallbackKey = `${absenderKey}|${datumIso}|${dateiName}`;

        if (vorhandeneSchluessel.has(eindeutigerKey) || vorhandeneSchluessel.has(fallbackKey)) {
          continue;
        }

        // Schutz 6: Typische Schlagworte für Nicht-Rechnungen
        if (dateiName.includes("zahlungsbeleg") || dateiName.includes("potwierdzenie") || dateiName.includes("ausgang bestellung")) {
          continue;
        }

        // In Google Drive ablegen
        const driveDatei = zielOrdner.createFile(anhang);
        const anzeigeSender = (rawReplyTo && cleanReplyTo !== cleanSender) 
          ? `${rawSender} (Antwort an: ${rawReplyTo})` 
          : rawSender;

        neueZeilen.push([
          nachricht.getDate(),
          anzeigeSender,
          driveDatei.getName(),
          driveDatei.getUrl(),
          eindeutigerKey,
          "Freigegeben für Versand"
        ]);

        vorhandeneSchluessel.add(eindeutigerKey);
        vorhandeneSchluessel.add(fallbackKey);
      }
    }
    emailThreads[i].addLabel(folgeLabelObj);
  }

  // 5. Batch-Insert am echten Zeilenende
  if (neueZeilen.length > 0) {
    const echteLetzteZeile = ermittleEchteLetzteZeile(tabellenBlatt);
    tabellenBlatt.getRange(echteLetzteZeile + 1, 1, neueZeilen.length, 6).setValues(neueZeilen);
    Logger.log(`${neueZeilen.length} geprüfte Rechnungen gespeichert.`);
  }
}

/**
 * Liest Kunden-E-Mails aus den letzten 200 Zeilen im Tab "RE", Spalte I ein.
 * Parst alle E-Mail-Adressen per RegEx (trennt Kommas, Semikolons, Leerzeichen sauber auf).
 */
function ladeKundenEmailSet(spreadsheetId, tabName) {
  const kundenSet = new Set();
  const maxZeilen = 200;

  try {
    const buchungsSs = SpreadsheetApp.openById(spreadsheetId);
    const reTab = buchungsSs.getSheetByName(tabName);
    if (!reTab) {
      Logger.log(`FEHLER: Tab '${tabName}' nicht gefunden.`);
      return kundenSet;
    }

    const letzteZeile = reTab.getLastRow();
    if (letzteZeile < 2) return kundenSet;

    const startZeile = Math.max(2, letzteZeile - maxZeilen + 1);
    const anzahlZeilen = letzteZeile - startZeile + 1;

    // Spalte I entspricht Spalte 9
    const daten = reTab.getRange(startZeile, 9, anzahlZeilen, 1).getValues();
    const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

    for (let i = 0; i < daten.length; i++) {
      const zellenInhalt = String(daten[i][0] || "").toLowerCase();
      if (!zellenInhalt) continue;

      const treffer = zellenInhalt.match(emailRegex);
      if (treffer) {
        for (let m = 0; m < treffer.length; m++) {
          const mail = treffer[m].trim();
          // Eigene Domain nicht als Kundenadresse in die Sperrliste aufnehmen
          if (!mail.includes("l8street.com")) {
            kundenSet.add(mail);
          }
        }
      }
    }

    Logger.log(`Kunden-Scan: Zeile ${startZeile} bis ${letzteZeile} durchsucht. ${kundenSet.size} Kunden-Mails geladen.`);
  } catch (err) {
    Logger.log("Fehler beim Laden des Kunden-Sheets: " + err.toString());
  }

  return kundenSet;
}

/**
 * Lässt AUSSCHLIESSLICH PDFs durch und filtert Stördateien
 */
function istZulaessigesPdf(dateiName, mimeTyp) {
  if (mimeTyp.startsWith("video/") || mimeTyp.startsWith("image/") || mimeTyp.startsWith("audio/")) return false;
  if (!dateiName.endsWith(".pdf") || !mimeTyp.includes("pdf")) return false;
  if (dateiName.match(/(datenschutz|widerruf|belehrung|wr-belehrung|agb|anleitung|merkblatt|rechnung_re-|mahnung_re-|anfrage\.xml)/i)) return false;
  return true;
}

/**
 * Initialisiert Tabellen-Header ohne Versatz
 */
function initialisiereSheetHeader(blatt) {
  if (blatt.getLastRow() === 0) {
    blatt.appendRow(["Datum / Zeit", "Absender", "Dateiname", "Link zur Datei", "UniqueKey", "Status"]);
    return;
  }
  if (!blatt.getRange(1, 5).getValue()) {
    blatt.getRange(1, 5).setValue("UniqueKey");
  }
  if (!blatt.getRange(1, 6).getValue()) {
    blatt.getRange(1, 6).setValue("Status");
  }
}

/**
 * Lädt vorhandene Schlüssel und normalisiert Altdaten von Caya konsistent
 */
function ladeVorhandeneSchluessel(blatt) {
  const menge = new Set();
  const anzahlZeilen = blatt.getLastRow();
  if (anzahlZeilen <= 1) return menge;

  const werte = blatt.getRange(2, 1, anzahlZeilen - 1, 5).getValues();
  for (let i = 0; i < werte.length; i++) {
    if (werte[i][4]) {
      menge.add(String(werte[i][4]).toLowerCase());
    }

    const absenderRoh = String(werte[i][1] || "").toLowerCase();
    const istCayaAltbestand = absenderRoh.includes("getcaya.com");
    const absenderFallback = istCayaAltbestand ? "caya" : bereinigeEmail(werte[i][1]);
    
    const fallback = `${absenderFallback}|${formatiereIsoDatum(werte[i][0])}|${String(werte[i][2]).trim().toLowerCase()}`;
    menge.add(fallback);
  }
  return menge;
}

/**
 * Ermittelt die echte letzte Zeile anhand der befüllten Spalte A
 */
function ermittleEchteLetzteZeile(blatt) {
  const datenA = blatt.getRange("A:A").getValues();
  for (let i = datenA.length - 1; i >= 0; i--) {
    if (datenA[i][0] !== "" && datenA[i][0] !== null && datenA[i][0] !== undefined) {
      return i + 1;
    }
  }
  return 1;
}

function bereinigeEmail(text) {
  if (!text) return "";
  const match = String(text).match(/<([^>]+)>/);
  return (match ? match[1] : String(text)).trim().toLowerCase();
}

function formatiereIsoDatum(datum) {
  try {
    const d = (datum instanceof Date) ? datum : new Date(datum);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  } catch (e) {}
  return String(datum).slice(0, 10);
}