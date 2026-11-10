/**
 * Haupt-Trigger: Verarbeitet Webhook-Daten von Twilio (SMS) und Vapi (Anruf/SMS).
 */
function doPost(e) {
  try {
    // 1. Pfad: Direkte Twilio-SMS (x-www-form-urlencoded)
    if (e.parameter && e.parameter.Body) {
      const zeilenDaten = extrahiereTwilioSms(e.parameter);
      speichereInTabelle(zeilenDaten);
      return erstelleXmlAntwort();
    }

    // 2. Pfad: Vapi JSON-Payloads (Anrufe oder Vapi-SMS)
    if (e.postData && e.postData.contents) {
      const payload = JSON.parse(e.postData.contents);
      const nachricht = payload.message || payload;
      const zeilenDaten = verarbeiteVapiPayload(nachricht);

      if (zeilenDaten) {
        speichereInTabelle(zeilenDaten);
      }
    }

    return erstelleJsonAntwort({ status: 'success' });
  } catch (fehler) {
    return erstelleJsonAntwort({ status: 'error', nachricht: fehler.message });
  }
}

/**
 * Bewertet den Nachrichtentyp aus Vapi und extrahiert die Daten einzeilig.
 */
function verarbeiteVapiPayload(nachricht) {
  const nachrichtenTyp = nachricht.type;
  const zeitstempel = new Date();

  // Vapi End-of-Call Report
  if (nachrichtenTyp === 'end-of-call-report') {
    const anruferNummer = nachricht.call?.customer?.number || nachricht.customer?.number || 'Unbekannt';
    const zusammenfassung = entferneZeilenumbrueche(nachricht.summary || nachricht.analysis?.summary || 'Keine Zusammenfassung');
    const transkript = entferneZeilenumbrueche(nachricht.transcript || 'Kein Transkript vorhanden');
    const audioLink = nachricht.recordingUrl || '';
    return [zeitstempel, anruferNummer, 'Sprachanruf', zusammenfassung, transkript, audioLink];
  }

  // Vapi Inbound SMS (falls über Vapi geroutet)
  if (nachrichtenTyp === 'inbound-sms' || nachrichtenTyp === 'sms') {
    const absender = nachricht.from || nachricht.customer?.number || 'Unbekannt';
    const inhalt = entferneZeilenumbrueche(nachricht.text || nachricht.body || nachricht.content || '');
    return [zeitstempel, absender, 'SMS (via Vapi)', 'Eingehende SMS', inhalt, '-'];
  }

  return null;
}

/**
 * Filtert Standard-Twilio-SMS-Parameter ohne Zeilenumbrüche.
 */
function extrahiereTwilioSms(parameter) {
  const zeitstempel = new Date();
  const absender = parameter.From || 'Unbekannt';
  const textInhalt = entferneZeilenumbrueche(parameter.Body || '');
  return [zeitstempel, absender, 'SMS (via Twilio)', 'Eingehende SMS', textInhalt, '-'];
}

/**
 * Ersetzt Zeilenumbrüche durch Trennzeichen, um Zeilenausdehnung zu verhindern.
 */
function entferneZeilenumbrueche(text) {
  if (typeof text !== 'string') return text;
  return text.replace(/(\r\n|\n|\r)/gm, ' | ').trim();
}

/**
 * Schreibt gebatcht eine Zeile in das aktive Sheet und fixiert die Höhe auf 21 Pixel.
 */
function speichereInTabelle(zeilenDaten) {
  if (!zeilenDaten || zeilenDaten.length === 0) return;

  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (tabellenBlatt.getLastRow() === 0) {
    const kopfzeile = [[
      'Datum & Uhrzeit',
      'Absender / Telefonnummer',
      'Kanal',
      'KI-Zusammenfassung',
      'Vollständiger Text / Transkript',
      'Aufnahme-Link'
    ]];
    tabellenBlatt.getRange(1, 1, 1, kopfzeile[0].length).setValues(kopfzeile);
    tabellenBlatt.setRowHeight(1, 25);
  }

  const naechsteZeile = tabellenBlatt.getLastRow() + 1;
  
  // 1. Daten gebatcht schreiben
  tabellenBlatt.getRange(naechsteZeile, 1, 1, zeilenDaten.length).setValues([zeilenDaten]);
  
  // 2. Zeilenhöhe explizit auf 21 festsetzen
  tabellenBlatt.setRowHeight(naechsteZeile, 21);
}

/**
 * Einmalige Hilfsfunktion: Setzt bestehende Zeilen manuell auf Höhe 21 zurück.
 */
function repariereZeilenhoehe() {
  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const maxZeilen = tabellenBlatt.getLastRow();
  if (maxZeilen > 1) {
    for (let zeile = 2; zeile <= maxZeilen; zeile++) {
      tabellenBlatt.setRowHeight(zeile, 21);
    }
  }
}

/**
 * Erzeugt eine leere TwiML-XML-Antwort.
 */
function erstelleXmlAntwort() {
  return ContentService.createTextOutput('<Response></Response>')
    .setMimeType(ContentService.MimeType.XML);
}

/**
 * Erzeugt eine standardisierte JSON-Antwort.
 */
function erstelleJsonAntwort(daten) {
  return ContentService.createTextOutput(JSON.stringify(daten))
    .setMimeType(ContentService.MimeType.JSON);
}