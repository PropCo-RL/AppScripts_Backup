/**
 * SCRIPT 3: Erstellt E-Mail-Entwürfe in Gmail mit dynamischer Spaltensuche
 * und baut einen klickbaren Web-App-Button für die Abreise/Verlängerung ein.
 */
function createEmailDraftswithWebApp() {
  // ---------------------------------------------------------------------------
  // Ihre finale Web-App-URL:
  // ---------------------------------------------------------------------------
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyALzy6TguW4rMDjqGzfab0WK5Jsg9vfJTAVJ_EEVNsOKTGjiyww59aePaIwFAArBv3aw/exec";

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

  // --- Dynamische Spaltensuche nach Überschriften ---
  const colMap = getColumnMapping_(outputSheet);
  
  if (colMap.name === -1 || colMap.apartment === -1 || colMap.abreise === -1 || colMap.email === -1 || colMap.statusQ === -1 || colMap.gmailId === -1) {
    Logger.log('Fehler: Mindestens eine der erforderlichen Spaltenüberschriften wurde im Tab "output" nicht gefunden.');
    return;
  }

  const lastCol = outputSheet.getLastColumn();
  const dataRange = outputSheet.getRange(3, 1, lastRow - 2, lastCol);
  const data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 3;

    const name = row[colMap.name - 1];
    const apartment = row[colMap.apartment - 1];
    const departureRaw = row[colMap.abreise - 1];
    const email = row[colMap.email - 1];
    const nextBookingRaw = colMap.anschlussbuchung !== -1 ? row[colMap.anschlussbuchung - 1] : null;
    const templateType = colMap.templateType !== -1 ? row[colMap.templateType - 1] : null;
    const statusQ = String(row[colMap.statusQ - 1]).trim();

    if (statusQ !== '') continue;

    if (!email || !String(email).includes('@')) {
      outputSheet.getRange(rowIndex, colMap.statusQ).setValue('Fehler: Keine E-Mail');
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

    const guestLink = `${WEB_APP_URL}?row=${rowIndex}&name=${encodeURIComponent(name)}&apt=${encodeURIComponent(apartment)}&depDate=${encodeURIComponent(departureDateFormatted)}`;

    let subject = '';
    let plainBody = '';
    let htmlBody = '';

    const textFooter = `\n\nFreundliche Grüße\n\n` +
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

    const htmlFooter = `<br><br>Freundliche Grüße<br><br>` +
      `<strong>L8 Street</strong><br><br>` +
      `© 2017-2026<br>` +
      `L8 Street GmbH<br>` +
      `Postfach CAYA595616<br>` +
      `96035 Bamberg<br><br>` +
      `WhatsApp: +4917684801295<br>` +
      `E-Mail: <a href="mailto:support@L8Street.com">support@L8Street.com</a><br>` +
      `Web: <a href="https://www.L8Street.com">https://www.L8Street.com</a><br>` +
      `Geschäftsführer: Raul Leneweit<br>` +
      `Sitz der Gesellschaft: Forststraße 65,<br>` +
      `75223 Niefern-Öschelbronn<br>` +
      `Amtsgericht Mannheim HRB: 72 85 52<br>` +
      `USt-IdNr.: DE 314 603 427`;

    // Dezentere, genau passende Button-Größe (Höhe 38px, Schrift 14px)
    const buttonHtml = `
      <div style="margin: 18px 0 10px 0;">
        <!--[if mso]>
        <v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" href="${guestLink}" style="height:38px;v-text-anchor:middle;width:280px;" arcsize="12%" stroke="f" fillcolor="#3182ce">
          <w:anchorlock/>
          <center style="color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:bold;">➔ Abreise / Verlängerung bestätigen</center>
        </v:roundrect>
        <![style]-->
        <a href="${guestLink}" target="_blank" style="background-color:#3182ce; border-radius:6px; color:#ffffff; display:inline-block; font-family:Arial, sans-serif; font-size:14px; font-weight:bold; line-height:38px; text-align:center; text-decoration:none; padding: 0 18px; -webkit-text-size-adjust:none;">➔ Abreise / Verlängerung bestätigen</a>
      </div>
    `;

    if (templateType === 'Option B (Anschlussbuchung)' && nextDateFormatted) {
      subject = `WICHTIG: Abreise am ${departureDayName}, ${departureDateFormatted} – Apartment ${apartment}`;
      
      plainBody = `Hallo ${name},\n\n` +
        `Ihr Apartment (${apartment}) ist ab ${nextDayName}, den ${nextDateFormatted} an den nächsten Gast vermietet.\n\n` +
        `Die Wohnung muss daher spätestens am ${departureDayName}, den ${departureDateFormatted} bis 10:00 Uhr geräumt sein.\n\n` +
        `Falls Sie verlängern möchten, wäre dies nur durch einen Umzug in ein anderes Apartment möglich. Bitte kontaktieren Sie uns dafür dringend direkt per E-Mail oder WhatsApp.\n\n` +
        `Hinweis: Für im Apartment hinterlassene Wertsachen übernehmen wir keine Haftung.\n\n` +
        `Alternativ können Sie uns auch per WhatsApp kontaktieren an 017684801295` +
        textFooter;

      htmlBody = `<p>Hallo ${name},</p>` +
        `<p>Ihr Apartment (${apartment}) ist ab ${nextDayName}, den ${nextDateFormatted} an den nächsten Gast vermietet.</p>` +
        `<p>Die Wohnung muss daher spätestens am <strong>${departureDayName}, den ${departureDateFormatted} bis 10:00 Uhr</strong> geräumt sein.</p>` +
        `<p>Falls Sie verlängern möchten, wäre dies nur durch einen Umzug in ein anderes Apartment möglich. Bitte kontaktieren Sie uns dafür dringend direkt per E-Mail oder WhatsApp.</p>` +
        `<p><em>Hinweis: Für im Apartment hinterlassene Wertsachen übernehmen wir keine Haftung.</em></p>` +
        `<p>Alternativ können Sie uns auch per WhatsApp kontaktieren an <strong>017684801295</strong></p>` +
        htmlFooter;

    } else {
      subject = `Abreise am ${departureDayName}, ${departureDateFormatted} – Apartment ${apartment}`;
      
      plainBody = `Hallo ${name},\n\n` +
        `bleibt es bei der Abreise am ${departureDayName}, den ${departureDateFormatted} (bis 10:00 Uhr Vormittag)?\n\n` +
        `Sofern wir nichts von Ihnen hören, gehen wir von der Abreise aus und planen die Reinigung sowie Neuvermietung. Für im Apartment hinterlassene Wertsachen können wir keine Haftung übernehmen.\n\n` +
        `Bitte bestätigen Sie Ihre Abreise oder eine gewünschte Verlängerung ganz einfach mit einem Klick:\n${guestLink}\n\n` +
        `Alternativ können Sie uns auch per WhatsApp kontaktieren an 017684801295` +
        textFooter;

      htmlBody = `<p>Hallo ${name},</p>` +
        `<p>bleibt es bei der Abreise am <strong>${departureDayName}, den ${departureDateFormatted}</strong> (bis 10:00 Uhr Vormittag)?</p>` +
        `<p>Sofern wir nichts von Ihnen hören, gehen wir von der Abreise aus und planen die Reinigung sowie Neuvermietung. Für im Apartment hinterlassene Wertsachen können wir keine Haftung übernehmen.</p>` +
        buttonHtml +
        `<p style="margin-top: 8px; margin-bottom: 0;">Alternativ können Sie uns auch per WhatsApp kontaktieren an <strong>017684801295</strong></p>` +
        htmlFooter;
    }

    try {
      const draft = GmailApp.createDraft(email, subject, plainBody, {
        htmlBody: htmlBody
      });
      
      const draftId = draft.getId();
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');

      outputSheet.getRange(rowIndex, colMap.statusQ).setValue(`Entwurf erstellt (${timestamp})`);
      outputSheet.getRange(rowIndex, colMap.gmailId).setValue(draftId);

      Logger.log(`Entwurf erstellt für ${email} (Zeile ${rowIndex})`);
    } catch (e) {
      outputSheet.getRange(rowIndex, colMap.statusQ).setValue('Fehler: ' + e.message);
      Logger.log(`Fehler bei Zeile ${rowIndex}: ${e.message}`);
    }
  }
}

/**
 * Hilfsfunktion: Ermittelt dynamisch die Spaltenindizes anhand der Überschriften in Zeile 1 oder 2
 */
function getColumnMapping_(sheet) {
  const mapping = {
    name: -1,
    apartment: -1,
    abreise: -1,
    email: -1,
    anschlussbuchung: -1,
    templateType: -1,
    statusQ: -1,
    gmailId: -1
  };

  const headers = sheet.getRange(1, 1, 2, sheet.getLastColumn()).getValues();

  for (let r = 0; r < headers.length; r++) {
    for (let c = 0; c < headers[r].length; c++) {
      const headerText = String(headers[r][c]).trim().toLowerCase();

      if (headerText === 'name') mapping.name = c + 1;
      else if (headerText === 'apartment') mapping.apartment = c + 1;
      else if (headerText === 'abreise') mapping.abreise = c + 1;
      else if (headerText === 'email') mapping.email = c + 1;
      else if (headerText === 'anschlussbuchung') mapping.anschlussbuchung = c + 1;
      else if (headerText.startsWith('gmail 1 gesendet')) mapping.statusQ = c + 1;
      else if (headerText === 'gmail 1 id') mapping.gmailId = c + 1;
    }
  }

  return mapping;
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