/**
 * Datei: 03_TransferService.gs
 * Verantwortlich für: Prüfung der aktiven Tabellenzeile und Erstellung des Transfers
 */

/**
 * Startet die Einzelüberweisung für die aktuell markierte Zeile
 */
function fuehreWiseEinzelueberweisungFuerAktiveZeileAus() {
  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet();
  const belegeBlatt = tabellenBlatt.getActiveSheet();

  if (belegeBlatt.getName() !== 'Belege') {
    Logger.log('Abbruch: Das aktive Blatt ist nicht "Belege".');
    return;
  }

  const zeilenNummer = belegeBlatt.getActiveCell().getRow();
  if (zeilenNummer <= 1) {
    Logger.log('Abbruch: Kopfzeile ausgewählt.');
    return;
  }

  verarbeiteEinzelneBelegZeile(zeilenNummer);
}

/**
 * Validiert die Zeile (Freigabe Raul Leneweit, leeres Q, Betrag, IBAN) und stößt Transfer an
 */
function verarbeiteEinzelneBelegZeile(zeilenNummer) {
  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet();
  const belegeBlatt = tabellenBlatt.getSheetByName('Belege');
  const kontakteBlatt = tabellenBlatt.getSheetByName('Kontakte');

  // Spalten A bis R (18 Spalten) gebatcht einlesen
  const zeilenWerte = belegeBlatt.getRange(zeilenNummer, 1, 1, 18).getValues()[0];

  // 1. Zwingende manuelle Freigabe in Spalte R prüfen
  const manuelleFreigabe = String(zeilenWerte[17] || '').trim();
  if (manuelleFreigabe !== 'Raul Leneweit') {
    Logger.log(`Zeile ${zeilenNummer} abgebrochen: Keine Freigabe durch "Raul Leneweit" in Spalte R.`);
    return;
  }

  // 2. Prüfen, ob Spalte Q noch leer ist
  const aktuellerLog = String(zeilenWerte[16] || '').trim();
  if (aktuellerLog !== '') {
    Logger.log(`Zeile ${zeilenNummer} übersprungen: Spalte Q nicht leer ("${aktuellerLog}").`);
    return;
  }

  const firmenName = String(zeilenWerte[1] || '').trim();
  const rechnungsNr = String(zeilenWerte[2] || '').trim();
  const bruttoBetrag = bereinigeBetrag(zeilenWerte[7]);

  if (!bruttoBetrag || bruttoBetrag <= 0) {
    schreibeLogInSpalteQ(belegeBlatt, zeilenNummer, 'Fehler: Ungültiger Bruttobetrag');
    return;
  }

  const kontakteLookup = erstelleKontakteLookup(kontakteBlatt);
  const zielIban = ermittleGueltigeIban(zeilenWerte, firmenName, kontakteLookup);

  if (!zielIban) {
    schreibeLogInSpalteQ(belegeBlatt, zeilenNummer, 'Fehler: Keine gültige IBAN (kein Match & nicht in Kontakten)');
    return;
  }

  const profilId = holeWiseKonfigurationsWert('WISE_PROFILE_ID');
  const apiKey = holeWiseKonfigurationsWert('wiseKey');

  const ergebnis = fuehreVollstaendigeTransaktionAus(apiKey, profilId, firmenName, zielIban, bruttoBetrag, rechnungsNr);
  schreibeLogInSpalteQ(belegeBlatt, zeilenNummer, ergebnis);
}

/**
 * Führt Quote, Empfänger, Transfer und Statusabfrage nacheinander aus
 */
function fuehreVollstaendigeTransaktionAus(apiKey, profilId, firmenName, iban, betrag, rechnungsNr) {
  try {
    // 1. Authentifizierte Quote anlegen
    const quoteId = erstelleAuthentifizierteWiseQuote({
      profileId: profilId,
      quellWaehrung: 'EUR',
      zielWaehrung: 'EUR',
      betrag: betrag,
      betragstyp: 'sourceAmount'
    });
    if (!quoteId) return 'Fehler: Quote-Erstellung fehlgeschlagen';

    // 2. Empfänger (SEPA IBAN) anlegen / abrufen
    const recipientId = holeOderErstelleEurEmpfaenger(profilId, firmenName, iban);
    if (!recipientId) return 'Fehler: Empfänger-Erstellung fehlgeschlagen';

    // 3. Transfer anlegen
    const saubereReferenz = bereinigeWiseReferenz(rechnungsNr);
    const transferId = erstelleStandardTransferDatensatz(apiKey, recipientId, quoteId, saubereReferenz);
    if (!transferId) return 'Fehler: Transfer-Erstellung fehlgeschlagen';

    // 4. Status ermitteln und Log für manuelle Freigabe setzen (Aufruf von Datei 04)
    const statusErgebnis = pruefeUndAktualisiereZahlung(profilId, transferId);
    return statusErgebnis.meldung;
  } catch (fehler) {
    return `Fehler: ${fehler.message}`;
  }
}

/**
 * Erstellt den Standard-Transfer nach Wise-Spezifikation (POST /v1/transfers)
 */
function erstelleStandardTransferDatensatz(apiKey, recipientId, quoteId, referenz) {
  const basisUrl = holeWiseBasisUrl();
  const url = `${basisUrl}/v1/transfers`;

  const nutzdaten = {
    targetAccount: Number(recipientId),
    quoteUuid: String(quoteId),
    customerTransactionId: Utilities.getUuid(),
    details: {
      reference: referenz
    }
  };

  const antwort = fuehreWiseApiAufrufAus(url, 'POST', apiKey, nutzdaten);
  if (!antwort || !antwort.id) {
    Logger.log('Fehler beim Anlegen des Transfers.');
    return null;
  }

  Logger.log(`Transfer erfolgreich erstellt. ID: ${antwort.id}, Status: ${antwort.status}`);
  return antwort.id;
}

/**
 * Bereinigt den Verwendungszweck gemäß Wise SEPA-Vorgaben [a-zA-Z0-9- ]* (max. 30 Zeichen)
 */
function bereinigeWiseReferenz(rechnungsNr) {
  if (!rechnungsNr) return 'Rechnung';

  const gesaeubert = String(rechnungsNr)
    .replace(/ä/g, 'ae')
    .replace(/ö/g, 'oe')
    .replace(/ü/g, 'ue')
    .replace(/ß/g, 'ss')
    .replace(/[^a-zA-Z0-9\- ]/g, '')
    .trim();

  const praefix = `RE ${gesaeubert}`.trim();
  return praefix.substring(0, 30);
}

/**
 * Ermittelt die IBAN: Spalte O nur bei "Match", sonst Kontakte Spalte G
 */
function ermittleGueltigeIban(zeile, firmenName, kontakteLookup) {
  const ibanGemini = String(zeile[14] || '').replace(/\s+/g, '').toUpperCase();
  const ibanMatchStatus = String(zeile[15] || '').trim().toLowerCase();

  // Nur wenn "Match" oder "IBAN Match" in Spalte P steht
  if (ibanMatchStatus.indexOf('match') !== -1 && ibanGemini.length >= 15) {
    return ibanGemini;
  }

  // Fallback: Kontakte Spalte G über Firmenname
  const firmaKey = firmenName.toLowerCase();
  if (kontakteLookup[firmaKey]) {
    return kontakteLookup[firmaKey];
  }

  return null;
}

/**
 * Liest Tab Kontakte gebatcht ein (Spalte A: Firma, Spalte G: IBAN)
 */
function erstelleKontakteLookup(kontakteBlatt) {
  const lookup = {};
  if (!kontakteBlatt) return lookup;
  const letzteZeile = kontakteBlatt.getLastRow();
  if (letzteZeile <= 1) return lookup;

  const kontakteDaten = kontakteBlatt.getRange(2, 1, letzteZeile - 1, 7).getValues();
  for (let i = 0; i < kontakteDaten.length; i++) {
    const firma = String(kontakteDaten[i][0] || '').trim().toLowerCase();
    const iban = String(kontakteDaten[i][6] || '').replace(/\s+/g, '').toUpperCase();
    if (firma && iban) {
      lookup[firma] = iban;
    }
  }
  return lookup;
}

/**
 * Bereinigt Währungs-Strings (z. B. "189,21 €") in Fließkommazahlen
 */
function bereinigeBetrag(rohWert) {
  if (typeof rohWert === 'number') return rohWert;
  if (!rohWert) return 0;

  const gesaeubert = String(rohWert)
    .replace('€', '')
    .replace(/\s+/g, '')
    .replace(/\./g, '')
    .replace(',', '.');

  const zahl = parseFloat(gesaeubert);
  return isNaN(zahl) ? 0 : zahl;
}

/**
 * Schreibt den Status gebatcht in Spalte Q (Spalte 17)
 */
function schreibeLogInSpalteQ(blatt, zeilenNummer, text) {
  blatt.getRange(zeilenNummer, 17, 1, 1).setValues([[text]]);
}