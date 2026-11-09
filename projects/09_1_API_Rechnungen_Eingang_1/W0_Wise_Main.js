/**
 * Datei: 00_MainOrchestrator.gs
 * Verantwortlich für: Zentraler Einstiegspunkt, Menüsteuerung und Vorab-Validierung
 */

/**
 * Erstellt beim Öffnen des Sheets das Aktionsmenü in der Menüleiste
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Wise Überweisung')
    .addItem('Aktive Zeile überweisen (Einzel-Transaktion)', 'starteWiseEinzelTransaktionOrchestrator')
    .addToUi();
}

/**
 * Zentraler Haupt-Orchestrator
 * Prüft Voraussetzungen und leitet die Überweisung der aktiven Zeile ein
 */
function starteWiseEinzelTransaktionOrchestrator() {
  const konfigurationStatus = pruefeWiseKonfiguration();
  if (!konfigurationStatus.gueltig) {
    Logger.log(`Abbruch Orchestrator: ${konfigurationStatus.meldung}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(konfigurationStatus.meldung, 'Konfigurationsfehler', 5);
    return;
  }

  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet();
  const aktivesBlatt = tabellenBlatt.getActiveSheet();

  if (aktivesBlatt.getName() !== 'Belege') {
    SpreadsheetApp.getActiveSpreadsheet().toast('Bitte wechsle zuerst in das Blatt "Belege".', 'Hinweis', 4);
    return;
  }

  const aktiveZeile = aktivesBlatt.getActiveCell().getRow();
  if (aktiveZeile <= 1) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Die Kopfzeile kann nicht verarbeitet werden.', 'Hinweis', 4);
    return;
  }

  Logger.log(`Orchestrator startet Überweisungsprozess für Zeile ${aktiveZeile}...`);
  verarbeiteEinzelneBelegZeile(aktiveZeile);
}

/**
 * Prüft, ob alle erforderlichen Schlüssel in den Script Properties existieren
 */
function pruefeWiseKonfiguration() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('wiseKey');
  if (!apiKey) {
    return { gueltig: false, meldung: 'Property "wiseKey" fehlt in den Script Properties.' };
  }

  const profilId = PropertiesService.getScriptProperties().getProperty('WISE_PROFILE_ID');
  if (!profilId) {
    return { gueltig: false, meldung: 'Property "WISE_PROFILE_ID" fehlt in den Script Properties.' };
  }

  return { gueltig: true, meldung: 'Konfiguration vollständig.' };
}