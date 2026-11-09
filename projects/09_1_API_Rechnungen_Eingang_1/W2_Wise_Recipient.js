/**
 * Datei: 02_RecipientService.gs
 * Verantwortlich für: Anlegen, Prüfen und Verifizieren (VoP/CoP) von EUR-Empfängern
 */

/**
 * Zentrale Orchestrator-Funktion für Datei 2:
 * Findet bestehenden Empfänger oder legt einen neuen an (inkl. VoP-Freigabe).
 */
function holeOderErstelleEurEmpfaenger(profilId, firmenName, iban) {
  if (!profilId || !firmenName || !iban) {
    Logger.log('Fehler: Unvollständige Empfängerdaten übergeben.');
    return null;
  }

  const bereinigteIban = iban.replace(/\s+/g, '').toUpperCase();
  const apiKey = holeWiseKonfigurationsWert('wiseKey');
  if (!apiKey) {
    Logger.log('Fehler: "wiseKey" fehlt in den Script Properties.');
    return null;
  }

  // 1. Prüfen, ob IBAN bereits im Wise-Profil registriert ist
  const bestehendeEmpfaengerId = findeBestehendeEmpfaengerId(apiKey, profilId, bereinigteIban);
  if (bestehendeEmpfaengerId) {
    Logger.log(`Vorhandener Empfänger wiederverwendet (ID: ${bestehendeEmpfaengerId}).`);
    return bestehendeEmpfaengerId;
  }

  // 2. Neu anlegen und ggf. Namensüberprüfung (VoP) akzeptieren
  return erstelleUndVerifiziereEurEmpfaenger(apiKey, profilId, firmenName, bereinigteIban);
}

/**
 * Legt einen neuen SEPA-IBAN-Empfänger bei Wise an (POST /v1/accounts)
 */
function erstelleUndVerifiziereEurEmpfaenger(apiKey, profilId, firmenName, iban) {
  const basisUrl = holeWiseBasisUrl();
  const url = `${basisUrl}/v1/accounts`;

  const nutzdaten = {
    currency: 'EUR',
    type: 'iban',
    profile: Number(profilId),
    ownedByCustomer: false,
    accountHolderName: firmenName.substring(0, 140),
    details: {
      legalType: 'BUSINESS',
      iban: iban
    }
  };

  const antwort = fuehreWiseApiAufrufAus(url, 'POST', apiKey, nutzdaten);
  if (!antwort || !antwort.id) {
    Logger.log(`Fehler: Empfänger "${firmenName}" konnte nicht angelegt werden.`);
    return null;
  }

  // VoP / Confirmation of Payee Check auswerten
  const bestaetigungErforderlich = pruefeObAkzeptanzErforderlich(antwort);
  if (bestaetigungErforderlich) {
    const freigabeErfolg = akzeptiereVoPNamensAbweichung(apiKey, antwort.id);
    if (!freigabeErfolg) {
      Logger.log(`Warnung: VoP-Akzeptanz für Empfänger ${antwort.id} fehlgeschlagen.`);
      return null;
    }
  }

  Logger.log(`Empfänger bereit für Transfer (ID: ${antwort.id}).`);
  return antwort.id;
}

/**
 * Prüft das confirmations.outcomes-Objekt auf requiresCustomerAcceptance
 */
function pruefeObAkzeptanzErforderlich(antwort) {
  if (!antwort.confirmations || !Array.isArray(antwort.confirmations.outcomes)) {
    return false;
  }

  for (let i = 0; i < antwort.confirmations.outcomes.length; i++) {
    const outcome = antwort.confirmations.outcomes[i];
    if (outcome.requiresCustomerAcceptance === true) {
      Logger.log(`VoP-Hinweis: ${outcome.message || outcome.outcome}. Manuelle Akzeptanz wird gesendet.`);
      return true;
    }
  }
  return false;
}

/**
 * Schaltet den Empfänger per PATCH /accounts/{id}/confirmations frei
 */
function akzeptiereVoPNamensAbweichung(apiKey, empfaengerId) {
  const basisUrl = holeWiseBasisUrl();
  const url = `${basisUrl}/2026Q4/accounts/${empfaengerId}/confirmations`;
  const nutzdaten = {
    acceptedOutcomes: true
  };

  const antwort = fuehreWiseApiAufrufAus(url, 'PATCH', apiKey, nutzdaten);
  if (!antwort || !antwort.confirmations || !antwort.confirmations.acceptedOutcomes) {
    return false;
  }

  Logger.log(`VoP-Ergebnis für Empfänger ${empfaengerId} erfolgreich bestätigt.`);
  return true;
}

/**
 * Durchsucht bestehende Empfänger nach übereinstimmender IBAN
 */
function findeBestehendeEmpfaengerId(apiKey, profilId, gesuchteIban) {
  const basisUrl = holeWiseBasisUrl();
  const url = `${basisUrl}/v2/accounts?profile=${profilId}&currency=EUR`;

  const antwort = fuehreWiseApiAufrufAus(url, 'GET', apiKey, null);
  if (!antwort) return null;

  const empfaengerListe = Array.isArray(antwort) ? antwort : (antwort.content || []);
  if (!Array.isArray(empfaengerListe) || empfaengerListe.length === 0) {
    return null;
  }

  for (let i = 0; i < empfaengerListe.length; i++) {
    const eintrag = empfaengerListe[i];
    const eintragIban = (eintrag.details && eintrag.details.iban)
      ? eintrag.details.iban.replace(/\s+/g, '').toUpperCase()
      : '';

    if (eintragIban === gesuchteIban) {
      return eintrag.id;
    }
  }

  return null;
}

/**
 * Führt HTTP-Aufrufe an die Wise API aus
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

  if (bodyDaten && (methode === 'POST' || methode === 'PUT' || methode === 'PATCH')) {
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
        Logger.log(`Warnung: Antwort kein JSON: ${textInhalt}`);
      }
    }

    if (statusKode < 200 || statusKode >= 300) {
      Logger.log(`Wise API Fehler (${statusKode}) an ${url}: ${textInhalt}`);
      return null;
    }

    return jsonDaten;
  } catch (netzwerkFehler) {
    Logger.log(`Netzwerkfehler an ${url}: ${netzwerkFehler.message}`);
    return null;
  }
}

/**
 * Liest Konfigurationswerte aus den Script Properties
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
 * Ermittelt die Basis-URL für Wise API-Aufrufe
 */
function holeWiseBasisUrl() {
  const umgebung = holeWiseKonfigurationsWert('WISE_ENVIRONMENT');
  if (umgebung === 'LIVE') {
    return 'https://api.wise.com';
  }
  return 'https://api.sandbox.transferwise.tech';
}