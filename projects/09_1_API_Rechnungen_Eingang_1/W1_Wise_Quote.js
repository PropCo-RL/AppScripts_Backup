/**
 * Datei: 01_QuoteService.gs
 * Verantwortlich für: Authentifizierte Quotes (Kurse, Gebühren, Quote-ID)
 */

/**
 * Zentrale Test- und Trigger-Funktion für Datei 1
 */
function testeErstelleWiseQuote() {
  const profilId = holeWiseKonfigurationsWert('WISE_PROFILE_ID');
  if (!profilId) {
    Logger.log('Abbruch: WISE_PROFILE_ID fehlt in den Script Properties.');
    return;
  }

  const anfrageParameter = {
    profileId: profilId,
    quellWaehrung: 'EUR',
    zielWaehrung: 'EUR',
    betrag: 10.00,
    betragstyp: 'sourceAmount' // 'sourceAmount' ODER 'targetAmount'
  };

  const quoteId = erstelleAuthentifizierteWiseQuote(anfrageParameter);
  if (quoteId) {
    Logger.log(`Quote erfolgreich angelegt! Gültige Quote-ID: ${quoteId}`);
  }
}

/**
 * Erstellt eine verbindliche, authentifizierte Quote für das angegebene Profil.
 * Liefert die quoteId zurück oder null bei einem Fehler.
 */
function erstelleAuthentifizierteWiseQuote(parameter) {
  if (!parameter || !parameter.profileId || !parameter.betrag) {
    Logger.log('Fehler: Ungültige Parameter für Quote-Erstellung.');
    return null;
  }

  const apiKey = holeWiseKonfigurationsWert('WISE_API_KEY');
  if (!apiKey) {
    Logger.log('Fehler: WISE_API_KEY nicht in den Script Properties gefunden.');
    return null;
  }

  const basisUrl = holeWiseBasisUrl();
  const url = `${basisUrl}/v3/profiles/${parameter.profileId}/quotes`;

  // Nutzdaten exakt nach Wise Balance-Funding Spezifikation
  const nutzdaten = {
    sourceCurrency: parameter.quellWaehrung,
    targetCurrency: parameter.zielWaehrung,
    payInMethod: 'BALANCE',
    preferredPayIn: 'BALANCE',
    sourceAmount: parameter.betragstyp === 'sourceAmount' ? parameter.betrag : null,
    targetAmount: parameter.betragstyp === 'targetAmount' ? parameter.betrag : null
  };

  const antwort = fuehreWiseApiAufrufAus(url, 'POST', apiKey, nutzdaten);
  if (!antwort || !antwort.id) {
    Logger.log('Fehler: Keine gültige Quote-ID in der Antwort erhalten.');
    return null;
  }

  Logger.log(`Status der Quote: ${antwort.status}`);
  Logger.log(`Gesperrter Kurs: ${antwort.rate}`);
  return antwort.id;
}

/**
 * Führt den eigentlichen HTTP-Aufruf via UrlFetchApp aus
 */
function fuehreWiseApiAufrufAus(url, methode, apiKey, bodyDaten) {
  const kopfzeilen = {
    'Authorization': `Bearer ${apiKey}`,
    'Content-Type': 'application/json',
    'X-External-Correlation-Id': Utilities.getUuid()
  };

  const optionen = {
    method: methode,
    headers: kopfzeilen,
    muteHttpExceptions: true
  };

  if (bodyDaten && (methode === 'POST' || methode === 'PUT')) {
    optionen.payload = JSON.stringify(bodyDaten);
  }

  try {
    const antwort = UrlFetchApp.fetch(url, optionen);
    const statusKode = antwort.getResponseCode();
    const textInhalt = antwort.getContentText();

    let jsonDaten = null;
    if (textInhalt) {
      try {
        jsonDaten = JSON.parse(textInhalt);
      } catch (parseFehler) {
        Logger.log(`JSON-Parse-Fehler: ${parseFehler.message}`);
      }
    }

    if (statusKode < 200 || statusKode >= 300) {
      Logger.log(`Wise API Fehler (${statusKode}) bei ${url}: ${textInhalt}`);
      return null;
    }

    return jsonDaten;
  } catch (netzwerkFehler) {
    Logger.log(`Netzwerkfehler bei Aufruf von ${url}: ${netzwerkFehler.message}`);
    return null;
  }
}

/**
 * Liest einen Wert aus den Script Properties aus
 */
function holeWiseKonfigurationsWert(schluessel) {
  try {
    return PropertiesService.getScriptProperties().getProperty(schluessel);
  } catch (fehler) {
    Logger.log(`Fehler beim Lesen von ${schluessel}: ${fehler.message}`);
    return null;
  }
}

/**
 * Gibt die Basis-URL zurück (LIVE oder SANDBOX)
 */
function holeWiseBasisUrl() {
  const umgebung = holeWiseKonfigurationsWert('WISE_ENVIRONMENT');
  if (umgebung === 'LIVE') {
    return 'https://api.wise.com';
  }
  return 'https://api.sandbox.transferwise.tech';
}
