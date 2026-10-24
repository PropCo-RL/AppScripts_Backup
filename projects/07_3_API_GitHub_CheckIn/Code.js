// ==========================================
// GOOGLE APPS SCRIPT BACKEND (CHECKIN & DRIVE UPLOAD)
// ==========================================

const TARGET_DRIVE_FOLDER_ID = "1Xh84O1T8fiGCcToOk-zszLzjBGcgfAHv";

function doGet(e) {
  const action = e.parameter.action;

  if (action === "getPreCacheJson") {
    return ContentService.createTextOutput(JSON.stringify(generatePreCacheData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "requestOTP") {
    return jsonResponse(requestOTP(e.parameter.email));
  }

  if (action === "verifyOTP") {
    return jsonResponse(verifyOTP(e.parameter.email, e.parameter.otp));
  }

  return jsonResponse({ success: false, message: "Ungültige Aktion" });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === "confirmCheckin") {
      const res = recordCheckin(data.email, data.lodgifyId, data.aptCode, data.guestName, "Check-in Erfolgreich");
      return jsonResponse(res);
    }

    if (data.action === "uploadPaymentProof") {
      const res = recordPaymentProof(data.email, data.lodgifyId, data.aptCode, data.fileName, data.fileBase64);
      return jsonResponse(res);
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function generatePreCacheData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Nachweise & Gemini-Status aus Tab 'checkin' einlesen
  const proofStatusMap = {};
  const sheetCheckin = ss.getSheetByName("checkin");

  if (sheetCheckin && sheetCheckin.getLastRow() >= 2) {
    const checkinData = sheetCheckin.getRange(2, 1, sheetCheckin.getLastRow() - 1, 8).getValues();
    checkinData.forEach(row => {
      const timestamp = row[0];
      const actionType = String(row[2] || "").trim();
      const lodgifyId = String(row[3] || "").trim();
      const geminiStatus = String(row[6] || "").trim();

      if (lodgifyId && actionType === "Zahlungsnachweis Upload") {
        proofStatusMap[lodgifyId] = {
          status: geminiStatus || "IN_PRUEFUNG",
          uploadedAt: formatDateStrWithTime(timestamp),
          fileName: String(row[5] || "").replace("Nachweis hochgeladen: ", "").trim()
        };
      }
    });
  }

  // 2. Tab d (Apartments) einlesen
  const sheetD = ss.getSheetByName("d");
  const aptMap = {};
  if (sheetD && sheetD.getLastRow() >= 2) {
    const dData = sheetD.getRange(2, 1, sheetD.getLastRow() - 1, 19).getValues();
    dData.forEach(row => {
      const link = String(row[3] || "").trim();
      if (!link) return;
      const match = link.match(/\/a\/([^\/\?]+)/i);
      if (!match || !match[1]) return;

      const code = match[1].toLowerCase().trim();
      const strasse = String(row[7] || "").trim();
      const nr = String(row[8] || "").trim();
      const plz = String(row[9] || "").trim();
      const stadt = String(row[10] || "").trim();

      aptMap[code] = {
        address: [strasse + " " + nr, plz + " " + stadt].filter(x => x.trim().length > 1).join(", ") || "Adresse auf Anfrage",
        etage: String(row[11] || "").trim() || "-",
        seite: String(row[12] || "").trim() || "-",
        checkinCode: String(row[18] || "").trim() || "Kein Code hinterlegt"
      };
    });
  }

  // 3. Tab RE (Buchungen) einlesen
  const sheetRE = ss.getSheetByName("RE");
  const bookingsByEmail = {};

  if (sheetRE && sheetRE.getLastRow() >= 2) {
    const reData = sheetRE.getRange(2, 1, sheetRE.getLastRow() - 1, 37).getValues();
    
    // Stichtag für alte Buchungen (90 Tage zurück)
    const now = new Date();
    const ninetyDaysAgo = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000));
    ninetyDaysAgo.setHours(0, 0, 0, 0);

    reData.forEach(row => {
      const rawEmailCell = String(row[8] || "").trim().toLowerCase(); // Spalte I (Index 8)
      if (!rawEmailCell) return;

      // E-Mails an Kommas oder Zeilenumbrüchen trennen und ungültige filtern
      const emailList = rawEmailCell.split(/[,;\n]+/).map(e => e.trim()).filter(e => e.length > 0 && e.includes("@"));
      if (emailList.length === 0) return;

      const anreiseDateObj = parseGermanDate(row[10]); // Spalte K (Index 10) ist Anreise!
      const statusRaw = String(row[31] || "").trim().toLowerCase(); // Spalte AF (Index 31)
      const lodgifyId = String(row[18] || "-").trim(); // Spalte S (Index 18)
      const pdfUrl = String(row[36] || "").trim(); // Spalte AK (Index 36)

      const proofInfo = proofStatusMap[lodgifyId] || null;

      const isPaidSheet = statusRaw.includes("bezahlt");
      const isPaidGemini = proofInfo && proofInfo.status === "ERFOLGREICH";
      const isPaid = isPaidSheet || isPaidGemini;

      const isOffen = statusRaw.includes("offen") || statusRaw === "";
      
      // Nur Buchungen anzeigen, die max. 90 Tage in der Vergangenheit liegen
      const isRecent = anreiseDateObj ? (anreiseDateObj >= ninetyDaysAgo) : true;

      if ((isPaid || isOffen) && isRecent) {
        const aptCode = String(row[13] || "").toLowerCase().trim(); // Spalte N (Index 13)
        const aptDetails = aptMap[aptCode] || {
          address: "Adresse auf Anfrage",
          etage: "-",
          seite: "-",
          checkinCode: "Nicht verfügbar"
        };

        // Buchung für JEDE E-Mail in der Zelle separat hinterlegen
        emailList.forEach(email => {
          if (!bookingsByEmail[email]) {
            bookingsByEmail[email] = [];
          }

          bookingsByEmail[email].push({
            lodgifyId: lodgifyId,
            gastName: String(row[1] || "Gast"), // Spalte B
            aptCode: aptCode,
            anreise: formatDateStr(row[10]), // Spalte K (Anreise)
            abreise: formatDateStr(row[11]), // Spalte L (Abreise)
            address: aptDetails.address,
            etage: aptDetails.etage,
            seite: aptDetails.seite,
            checkinCode: aptDetails.checkinCode,
            isPaid: isPaid,
            pdfUrl: pdfUrl,
            proofInfo: proofInfo
          });
        });
      }
    });
  }

  return {
    updatedAt: new Date().toISOString(),
    data: bookingsByEmail
  };
}

function requestOTP(inputVal) {
  const cleanEmail = String(inputVal || "").trim().toLowerCase();
  if (!cleanEmail) return { success: false, message: "Bitte E-Mail eingeben." };

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  PropertiesService.getScriptProperties().setProperty("CHECKIN_OTP_" + cleanEmail, otp);

  GmailApp.sendEmail(
    cleanEmail,
    "Ihr Check-in Sicherheitscode – L8 Street",
    "Ihr Sicherheitscode lautet: " + otp,
    {
      htmlBody: `
        <div style="font-family: Arial, sans-serif; max-width: 500px; padding: 20px; border: 1px solid #e5e7eb; border-radius: 8px;">
          <h3 style="color: #111827; margin-top: 0;">L8 Street Check-in Portal</h3>
          <p>Ihr Sicherheitscode für den Online Check-in lautet:</p>
          <div style="background: #f3f4f6; padding: 12px; font-size: 24px; font-weight: bold; letter-spacing: 4px; text-align: center; border-radius: 6px; margin: 15px 0;">
            ${otp}
          </div>
        </div>
      `
    }
  );

  return { success: true };
}

function verifyOTP(inputVal, otp) {
  const cleanEmail = String(inputVal || "").trim().toLowerCase();
  const storedOTP = PropertiesService.getScriptProperties().getProperty("CHECKIN_OTP_" + cleanEmail);

  if (!storedOTP || storedOTP !== String(otp).trim()) {
    return { found: false, message: "Ungültiger oder abgelaufener Sicherheitscode." };
  }

  PropertiesService.getScriptProperties().deleteProperty("CHECKIN_OTP_" + cleanEmail);
  return { success: true, email: cleanEmail };
}

function recordCheckin(email, lodgifyId, aptCode, guestName, statusTxt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let checkinSheet = ss.getSheetByName("checkin");

  if (!checkinSheet) {
    checkinSheet = ss.insertSheet("checkin");
    checkinSheet.appendRow(["Zeitstempel", "E-Mail", "Gast / Firma", "Lodgify ID / Buchung", "Apartment Kürzel", "Status / Nachweis Link", "Gemini_Status", "Gemini_Interne_Notiz"]);
    checkinSheet.getRange("1:1").setFontWeight("bold");
  }

  checkinSheet.appendRow([
    new Date(),
    email,
    guestName || "-",
    lodgifyId || "-",
    aptCode || "-",
    statusTxt || "Aktion",
    "-",
    "-"
  ]);

  return { success: true, message: "Aktion registriert!" };
}

function recordPaymentProof(email, lodgifyId, aptCode, fileName, fileBase64) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let checkinSheet = ss.getSheetByName("checkin");

  if (!checkinSheet) {
    checkinSheet = ss.insertSheet("checkin");
    checkinSheet.appendRow(["Zeitstempel", "E-Mail", "Gast / Firma", "Lodgify ID / Buchung", "Apartment Kürzel", "Status / Nachweis Link", "Gemini_Status", "Gemini_Interne_Notiz"]);
    checkinSheet.getRange("1:1").setFontWeight("bold");
  }

  let fileUrl = "Kein Link erzeugt";

  if (fileName && fileBase64) {
    try {
      const targetFolder = DriveApp.getFolderById(TARGET_DRIVE_FOLDER_ID);
      const bytes = Utilities.base64Decode(fileBase64);
      const blob = Utilities.newBlob(bytes, null, fileName);

      const newFile = targetFolder.createFile(blob);
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = newFile.getUrl();
    } catch (err) {
      fileUrl = "Fehler beim Drive-Upload: " + err.toString();
    }
  }

  checkinSheet.appendRow([
    new Date(),
    email,
    "Zahlungsnachweis Upload",
    lodgifyId || "-",
    aptCode || "-",
    fileUrl,
    "IN_PRUEFUNG",
    "Upload empfangen, wartet auf Gemini-Prüfung"
  ]);

  return { success: true, message: "Upload erfolgreich" };
}

function parseGermanDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'string') {
    const p = val.trim().split('.');
    if (p.length === 3) {
      return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
    }
  }
  return null;
}

function formatDateStr(dateVal) {
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd.MM.yyyy");
  }
  return dateVal || "-";
}

function formatDateStrWithTime(dateVal) {
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
  }
  return dateVal || "-";
}