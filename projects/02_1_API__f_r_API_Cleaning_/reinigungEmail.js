function sendCleaningMails() {
  const ss = SpreadsheetApp.getActive();
  const s = ss.getSheetByName("wCl"); 
  if (!s) throw new Error('Sheet "wCl" nicht gefunden');

  const startRow = 5;
  const last = s.getLastRow();
  if (last < startRow) return;

  const data = s.getRange(startRow, 1, last - (startRow - 1), 4).getValues();
  const tz = ss.getSpreadsheetTimeZone();

  const today = new Date(); today.setHours(0,0,0,0);
  const end = new Date(today); end.setDate(end.getDate() + 7);

  const firms = {};
  const masterContent = { today: [], upcoming: {} };

  data.forEach((r) => {
    let rawDate = r[0]; // Spalte A
    const apt   = r[1]; // Spalte B
    const email = r[3]; // Spalte D

    if (!rawDate || !email) return;

    let clnDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (isNaN(clnDate.getTime())) {
      if (typeof rawDate === 'string' && rawDate.includes('.')) {
        const parts = rawDate.split('.');
        clnDate = new Date(parts[2], parts[1] - 1, parts[0]);
      } else {
        return; 
      }
    }

    const d = new Date(clnDate); d.setHours(0,0,0,0);
    if (d < today || d > end) return;

    const emailClean = email.toString().trim();
    if (!emailClean.includes("@")) return;

    if (!firms[emailClean]) {
      firms[emailClean] = { today: [], upcoming: {} };
    }

    const aptInfo = apt || "Apartment ohne Namen";
    const dateKey = Utilities.formatDate(d, tz, "dd.MM.yyyy");

    if (d.getTime() === today.getTime()) {
      firms[emailClean].today.push(aptInfo);
      masterContent.today.push(`${aptInfo} (${emailClean})`);
    } else {
      if (!firms[emailClean].upcoming[dateKey]) firms[emailClean].upcoming[dateKey] = [];
      firms[emailClean].upcoming[dateKey].push(aptInfo);
      
      if (!masterContent.upcoming[dateKey]) masterContent.upcoming[dateKey] = [];
      masterContent.upcoming[dateKey].push(`${aptInfo} (${emailClean})`);
    }
  });

  // --- 1. SEND INDIVIDUAL MAILS TO CLEANERS (WITH BCC) ---
  for (const email in firms) {
    const f = firms[email];
    const body = buildEmailText(f, "Guten Morgen,\n\nhier ist dein persönlicher Reinigungsplan.\n\n", tz);

    try {
      GmailApp.sendEmail(email, `Tägliche Übersicht L8 Termine - ${Utilities.formatDate(today, tz, "dd.MM.yyyy")}`, body, {
        from: "support@l8street.com",
        name: "L8 Street Support",
        bcc: "info@L8Street.com" // You will receive a hidden copy of every cleaner email
      });
    } catch (e) {
      console.error("Fehler bei " + email + ": " + e.message);
    }
  }

  // --- 2. SEND ONE MASTER MAIL TO MANAGEMENT ---
  const masterRecipients = "info@L8Street.com, rechnung@L8Street.com, sirato31215@gmail.com";
  const masterBody = buildEmailText(masterContent, "Guten Morgen Team,\n\nhier ist die Übersicht aller Reinigungen.\n\n", tz);

  try {
    GmailApp.sendEmail(masterRecipients, `Alle L8 Termine - ${Utilities.formatDate(today, tz, "dd.MM.yyyy")}`, masterBody, {
      from: "support@l8street.com",
      name: "L8 Street Support"
    });
  } catch (e) {
    console.error("Fehler beim Senden der Master-Liste: " + e.message);
  }
}

function buildEmailText(dataObj, greeting, tz) {
  let body = greeting;
  body += "========================================\n";
  body += "TERMINE HEUTE\n";
  body += "========================================\n";
  
  if (dataObj.today.length) {
    body += dataObj.today.map(a => "[ ] " + a).join("\n");
  } else {
    body += "- Keine Reinigungen für heute geplant -";
  }
  
  body += "\n\n\n";
  body += "========================================\n";
  body += "VORSCHAU 7 TAGE\n";
  body += "========================================\n";
  
  const today = new Date();
  const sortedDays = Object.keys(dataObj.upcoming).sort((a, b) => {
    const pA = a.split('.'); const pB = b.split('.');
    return new Date(pA[2], pA[1]-1, pA[0]) - new Date(pB[2], pB[1]-1, pB[0]);
  });

  if (sortedDays.length > 0) {
    sortedDays.forEach(day => {
      body += `\n${day}:\n`;
      body += dataObj.upcoming[day].map(a => "  > " + a).join("\n");
      body += "\n";
    });
  } else {
    body += "- Keine weiteren geplanten Reinigungen -";
  }

  body += "\n\n________________________________________\n";
  body += "Viele Grüße\n";
  body += "L8 Street Team";
  return body;
}