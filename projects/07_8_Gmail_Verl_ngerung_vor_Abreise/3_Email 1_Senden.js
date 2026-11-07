/**
* - SCRIPT 3: Versendet E-Mails direkt in Gmail und schreibt Status + Zeitstempel
* - in Spalte Q sowie einen Status-Vermerk in Spalte R.
*/
function sendEmailsDirectly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName('output');
  if (!outputSheet) {
    Logger.log('Fehler: Tabellenblatt "output" nicht gefunden.');
    return;
  }

  const lastRow = outputSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log('Keine Daten im Tab "output" vorhanden.');
    return;
  }

  // Liest alle Spalten bis R ein (18 Spalten)
  const dataRange = outputSheet.getRange(3, 1, lastRow - 2, 18);
  const data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 3;

    const name = row[0]; // Spalte A: Name
    const apartment = row[1]; // Spalte B: Apartment
    const departureRaw = row[2]; // Spalte C: Abreisedatum
    const email = row[3]; // Spalte D: E-Mail
    const nextBookingRaw = row[8]; // Spalte I: Anschlussbuchung Datum
    const templateType = row[10]; // Spalte K: Template Typ
    const statusQ = String(row[16]).trim(); // Spalte Q: Gmail 1 Gesendet / Status

    // Nur verarbeiten, wenn in Spalte Q noch NICHTS steht
    if (statusQ !== '') continue;

    // Prüfen, ob eine gültige E-Mail-Adresse vorliegt
    if (!email || !String(email).includes('@')) {
      outputSheet.getRange(rowIndex, 17).setValue('Fehler: Keine E-Mail'); // Spalte Q
      continue;
    }

    const depDateObj = parseDate_(departureRaw);
    if (!depDateObj) continue;

    const departureDayName = getGermanDayName_(depDateObj);
    const departureDateFormatted = Utilities.formatDate(depDateObj, Session.getScriptTimeZone(), 'dd.MM.yyyy');

    let nextDayName = '';
    let nextDateFormatted = '';

    if (nextBookingRaw) {
      const nextDateObj = parseDate_(nextBookingRaw);
      if (nextDateObj) {
        nextDayName = getGermanDayName_(nextDateObj);
        nextDateFormatted = Utilities.formatDate(nextDateObj, Session.getScriptTimeZone(), 'dd.MM.yyyy');
      }
    }

    let subject = '';
    let body = '';

    const footer = `Freundliche Grüße\n\n` +
      `L8 Street\n\n` +
      `© 2017-2026\n` +
      `L8 Street GmbH\n` +
      `Postfach CAYA595616\n` +
      `96035 Bamberg\n\n` +
      `WhatsApp: +4917684801295\n` +
      `E-Mail: support@L8Street.com\n` +
      `Web: https://www.L8Street.com\n` +
      `Geschäftsführer: Raul Leneweit\n` +
      `Sitz der Gesellschaft: Forststraße 65,\n` +
      `75223 Niefern-Öschelbronn\n` +
      `Amtsgericht Mannheim HRB: 72 85 52\n` +
      `USt-IdNr.: DE 314 603 427`;

    if (templateType === 'Option B (Anschlussbuchung)' && nextDateFormatted) {
      subject = `WICHTIG: Abreise am ${departureDayName}, ${departureDateFormatted} – Apartment ${apartment}`;
      body = `Hallo ${name},\n\n` +
        `Ihr Apartment (${apartment}) ist ab ${nextDayName}, den ${nextDateFormatted} an den nächsten Gast vermietet.\n\n` +
        `Die Wohnung muss daher spätestens am ${departureDayName}, den ${departureDateFormatted} bis 10:00 Uhr geräumt sein.\n\n` +
        `Falls Sie verlängern möchten, wäre dies nur durch einen Umzug in ein anderes Apartment möglich. Bitte kontaktieren Sie uns dafür dringend direkt per E-Mail oder WhatsApp.\n\n` +
        `Hinweis: Für im Apartment hinterlassene Wertsachen übernehmen wir keine Haftung.\n\n` +
        footer;
    } else {
      subject = `Abreise am ${departureDayName}, ${departureDateFormatted} – Apartment ${apartment}`;
      body = `Hallo ${name},\n\n` +
        `bleibt es bei der Abreise am ${departureDayName}, den ${departureDateFormatted} (bis 10:00 Uhr Vormittag)?\n\n` +
        `Sofern wir nichts anderes hören gehen wir von der Abreise aus und planen die Reinigung sowie Neuvermietung. Für im Apartment hinterlassene Wertsachen können wir keine Haftung übernehmen. \n\n` +
        footer;
    }

    try {
      // E-Mail direkt über Gmail senden
      GmailApp.sendEmail(email, subject, body);

      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');

      // Status in Spalte Q und Vermerk in Spalte R eintragen
      outputSheet.getRange(rowIndex, 17).setValue(`Gesendet (${timestamp})`); // Spalte Q
      outputSheet.getRange(rowIndex, 18).setValue('E-Mail versendet'); // Spalte R
      Logger.log(`E-Mail gesendet an ${email} (Zeile ${rowIndex})`);
    } catch (e) {
      outputSheet.getRange(rowIndex, 17).setValue('Fehler: ' + e.message); // Spalte Q
      Logger.log(`Fehler bei Zeile ${rowIndex}: ${e.message}`);
    }
  }
}

function getGermanDayName_(dateObj) {
  const days = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
  return days[dateObj.getDay()];
}

function parseDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return new Date(val.getTime());
  if (typeof val === 'string') {
    const parts = val.trim().split('.');
    if (parts.length === 3) {
      return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}