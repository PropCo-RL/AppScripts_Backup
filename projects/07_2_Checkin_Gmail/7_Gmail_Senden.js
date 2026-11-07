function processEmailDrafts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName("Output");
  var dSheet = ss.getSheetByName("D");
  if (!outputSheet || !dSheet) {
    Logger.log("Fehler: 'Output' oder 'D' Sheet existiert nicht.");
    return;
  }

  var outputData = outputSheet.getDataRange().getValues();
  if (outputData.length <= 1) return; // Keine Daten
  var outputHeaders = outputData[0];

  // 1. Spalten-Indizes dynamisch anhand der Header im Output-Sheet ermitteln
  var nameColIdx = outputHeaders.indexOf("Name"); // Spalte B
  var emailColIdx = outputHeaders.indexOf("Email"); // Spalte C
  var aptIdColIdx = outputHeaders.indexOf("Apartment id"); // Spalte F
  var rechnungStatusColIdx = outputHeaders.indexOf("Rechnung Status"); // Spalte L
  var codeGesendetColIdx = outputHeaders.indexOf("Code Gesendet (Gmail)"); // Spalte O
  var emailStatusColIdx = outputHeaders.indexOf("Email Status") !== -1 ? outputHeaders.indexOf("Email Status") : outputHeaders.indexOf("Email status");
  var anreiseColIdx = outputHeaders.indexOf("Anreise"); // Spalte D
  var erinnerungColIdx = outputHeaders.indexOf("Erinnerung Gesendet"); // Spalte P

  // Index für Spalte S (Gemini Status) ermitteln
  var geminiStatusColIdx = outputHeaders.indexOf("Gemini Status");
  if (geminiStatusColIdx === -1) geminiStatusColIdx = 18; // Spalte S (Index 18)

  // Absicherung per Index (0-basiert):
  if (nameColIdx === -1) nameColIdx = 1;
  if (emailColIdx === -1) emailColIdx = 2;
  if (aptIdColIdx === -1) aptIdColIdx = 5;
  if (rechnungStatusColIdx === -1) rechnungStatusColIdx = 11;
  if (codeGesendetColIdx === -1) codeGesendetColIdx = 14;
  if (erinnerungColIdx === -1) erinnerungColIdx = 15;

  // 2. Tab "D" einlesen und Apartment-Daten in einer Map indizieren
  var dData = dSheet.getDataRange().getValues();
  var aptMap = {};
  for (var i = 1; i < dData.length; i++) {
    var fullUrl = String(dData[i][3] || "").trim();
    var shortId = extractAptId(fullUrl);
    if (shortId !== "") {
      aptMap[shortId] = {
        strasse: dData[i][7] || "",
        hausnr: dData[i][8] || "",
        plz: dData[i][9] || "",
        stadt: dData[i][10] || "",
        etage: dData[i][11] || "",
        seite: dData[i][12] || "",
        code: dData[i][23] || ""
      };
    }
  }

  var now = new Date();
  var bccEmail = "cc@l8street.com";
  var footerText = "\n\nFreundliche Grüße\n\nL8 Street\n\n© 2017-2026\nL8 Street GmbH\nPostfach CAYA595616\n96035 Bamberg\n\nWhatsApp: +4917684801295\nE-Mail: support@L8Street.com\nWeb: https://www.L8Street.com\nGeschäftsführer: Raul Leneweit\nSitz der Gesellschaft: Forststraße 65,\n75223 Niefern-Öschelbronn\nAmtsgericht Mannheim HRB: 72 85 52\nUSt-IdNr.: DE 314 603 427";

  // 3. Zeilen im Output-Sheet iterieren
  for (var r = 1; r < outputData.length; r++) {
    var row = outputData[r];
    var email = String(row[emailColIdx] || "").trim();
    var aptId = String(row[aptIdColIdx] || "").trim();
    var rechnungStatus = String(row[rechnungStatusColIdx] || "").trim();
    var codeGesendetVal = String(row[codeGesendetColIdx] || "").trim();
    var erinnerungVal = String(row[erinnerungColIdx] || "").trim();
    var geminiStatusVal = String(row[geminiStatusColIdx] || "").trim(); // Status aus Spalte S
    var name = String(row[nameColIdx] || "").trim();
    if (!name) name = "Gast";
    if (!email) continue;

    // Gesamtprüfbedingung: Entweder in Spalte L "bezahlt" ODER Gemini-Status in Spalte S "ERFOLGREICH"
    var isBezahltOderGeminiErfolgreich = (rechnungStatus.toLowerCase() === "bezahlt" || geminiStatusVal === "ERFOLGREICH");

    // --- FALL 1: Code an den Kunden senden (Check-In Mail) ---
    if (codeGesendetVal === "" && isBezahltOderGeminiErfolgreich) {
      var aptDetails = aptMap[aptId];
      var codeText = aptDetails ? aptDetails.code : "[Code nicht gefunden]";
      var strasse = aptDetails ? aptDetails.strasse : "";
      var hausnr = aptDetails ? aptDetails.hausnr : "";
      var plz = aptDetails ? aptDetails.plz : "";
      var stadt = aptDetails ? aptDetails.stadt : "";
      var etage = aptDetails ? aptDetails.etage : "";
      var seite = aptDetails ? aptDetails.seite : "";

      var mailSubject = "Ihre Check-in Daten für Ihren Aufenthalt - L8 Street";
      var mailBody = "Hallo " + name + ",\n\n" +
        "vielen Dank für Ihre Zahlung! Hier sind Ihre Check-in Informationen:\n\n" +
        "Zugangscode: " + codeText + "\n\n" +
        "Adresse der Unterkunft:\n" +
        strasse + " " + hausnr + "\n" +
        plz + " " + stadt + "\n" +
        "Etage: " + etage + " | Seite: " + seite + "\n\n" +
        "Am Tag der Anreise können Sie ab 16 Uhr anreisen (Abreise bis 10 Uhr Vormittag)." +
        footerText;

      // Direkt versenden via Draft-Send-Workflow (garantiert eine Message-ID ohne Absturz)
      var draft = GmailApp.createDraft(email, mailSubject, mailBody, { bcc: bccEmail });
      var sentMessage = draft.send();

      var msgId = "";
      try {
        msgId = sentMessage ? sentMessage.getId() : "";
      } catch (eId) {
        msgId = "nicht_ermittelbar";
      }

      var statusJson = JSON.stringify({
        status: "Gesendet",
        timestamp: formatDateIso(now),
        messageId: msgId,
        versendetAn: email
      });

      outputSheet.getRange(r + 1, codeGesendetColIdx + 1).setValue(statusJson);
      SpreadsheetApp.flush();
      continue;
    }

    // --- FALL 2: Erinnerungs-E-Mail senden ---
    // Nur senden, wenn WEDER bezahlt NOCH ein erfolgreicher Gemini-Beleg vorliegt!
    if (codeGesendetVal === "" && erinnerungVal === "" && !isBezahltOderGeminiErfolgreich) {
      var emailStatusVal = (emailStatusColIdx !== -1) ? String(row[emailStatusColIdx]) : "";
      var anreiseVal = (anreiseColIdx !== -1) ? row[anreiseColIdx] : null;

      var isOlderThan6Hours = checkEmailStatusTimestamp(emailStatusVal, now);
      var isTodayAndAfter13 = checkIsAnreiseTodayAndAfter13(anreiseVal, now);

      if (isOlderThan6Hours && isTodayAndAfter13) {
        var errSubject = "Buchung und Zugangsdaten - L8 Street";
        var errBody = "Hallo " + name + ",\n\n" +
          "bitte denken Sie noch an den Beleg für die ausgeführte Überweisung. Sobald wir diesen erhalten haben dürfen wir sofort den Code für die Anreise zusenden." +
          footerText;

        // Direkt versenden via Draft-Send-Workflow
        var errDraft = GmailApp.createDraft(email, errSubject, errBody, { bcc: bccEmail });
        var errSentMessage = errDraft.send();

        var errMsgId = "";
        try {
          errMsgId = errSentMessage ? errSentMessage.getId() : "";
        } catch (eId2) {
          errMsgId = "nicht_ermittelbar";
        }

        var errStatusJson = JSON.stringify({
          status: "Erinnerung Gesendet",
          timestamp: formatDateIso(now),
          messageId: errMsgId,
          versendetAn: email
        });

        outputSheet.getRange(r + 1, erinnerungColIdx + 1).setValue(errStatusJson);
        SpreadsheetApp.flush();
      }
    }
  }
}

// ==========================================
// HILFSFUNKTIONEN
// ==========================================

function extractAptId(url) {
  if (!url) return "";
  var match = url.match(/\/a\/([^\/\?#]+)/i);
  if (match && match[1]) {
    return match[1].trim();
  }
  return url.trim();
}

function checkEmailStatusTimestamp(statusText, now) {
  if (!statusText) return false;
  var match = statusText.match(/\((\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})\)/);
  if (!match) return false;

  var day = parseInt(match[1], 10);
  var month = parseInt(match[2], 10) - 1;
  var year = parseInt(match[3], 10);
  var hours = parseInt(match[4], 10);
  var minutes = parseInt(match[5], 10);

  var statusDate = new Date(year, month, day, hours, minutes);
  var sixHoursInMs = 6 * 60 * 60 * 1000;

  return (now.getTime() - statusDate.getTime()) >= sixHoursInMs;
}

function checkIsAnreiseTodayAndAfter13(anreiseVal, now) {
  var anreiseDate = parseDateSafe(anreiseVal);
  if (!anreiseDate || isNaN(anreiseDate.getTime())) return false;

  // Bricht ab, wenn es vor 13:00 Uhr ist
  if (now.getHours() < 13) return false;

  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var checkDate = new Date(anreiseDate.getFullYear(), anreiseDate.getMonth(), anreiseDate.getDate());

  return checkDate.getTime() === today.getTime();
}

function formatDateIso(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
}

function parseDateSafe(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") return value;
  if (typeof value === "string") {
    value = value.trim();
    var deMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (deMatch) return new Date(parseInt(deMatch[3], 10), parseInt(deMatch[2], 10) - 1, parseInt(deMatch[1], 10));

    var usMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (usMatch) return new Date(parseInt(usMatch[3], 10), parseInt(usMatch[2], 10) - 1, parseInt(usMatch[1], 10));

    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}