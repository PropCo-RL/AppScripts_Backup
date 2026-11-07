const TARGET_IBAN = "DE41666500850008979332"; // L8 Street GmbH Ziel-IBAN (ohne Leerzeichen)

function processPaymentProofsWithGemini() {
  Logger.log("Starte Gemini-Prüfung für E-Mail- & WhatsApp-Zahlungsbelege...");
  const apiKey = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  if (!apiKey) {
    Logger.log("FEHLER: 'Gemini-Key' nicht in ScriptProperties gefunden!");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName("Output");
  if (!outputSheet) return;

  const data = outputSheet.getDataRange().getValues();
  if (data.length <= 1) return;
  const headers = data[0];

  // Spaltenzuordnungen
  const driveBelegColIdx = headers.indexOf("Drive Zahlungsbeleg") !== -1 ? headers.indexOf("Drive Zahlungsbeleg") : 17; // Spalte R (Index 17)
  const rechnungUrlColIdx = headers.indexOf("Google Drive Location") !== -1 ? headers.indexOf("Google Drive Location") : (headers.indexOf("Rechng Google") !== -1 ? headers.indexOf("Rechng Google") : 12); // Spalte M (Index 12)

  // Spalten S & T für Ergebnisse anlegen, falls noch nicht vorhanden
  let statusColIdx = headers.indexOf("Gemini Status");
  if (statusColIdx === -1) {
    statusColIdx = 18; // Spalte S
    outputSheet.getRange(1, 19).setValue("Gemini Status");
  }

  let noteColIdx = headers.indexOf("Gemini Notiz");
  if (noteColIdx === -1) {
    noteColIdx = 19; // Spalte T
    outputSheet.getRange(1, 20).setValue("Gemini Notiz");
  }

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const fileUrlString = String(row[driveBelegColIdx] || "").trim();
    const currentGeminiStatus = String(row[statusColIdx] || "").trim();
    const invoiceUrl = String(row[rechnungUrlColIdx] || "").trim();

    // Nur Zeilen mit Drive-Links prüfen, die noch nicht abgeschlossen oder erfolgreich sind
    if (fileUrlString.includes("drive.google.com") && (currentGeminiStatus === "" || currentGeminiStatus === "IN_PRUEFUNG")) {
      // 🎯 ALLER BELEGE-LINKS SPLITTEN
      const rawUrls = fileUrlString.split(",").map(u => u.trim()).filter(u => u.length > 0);
      const proofBlobs = [];

      // 🎯 VORFILTERUNG & DRIVE-DOWNLOAD (KOSTEN SPAREN)
      for (let u = 0; u < rawUrls.length; u++) {
        const fileId = extractDriveId(rawUrls[u]);
        if (!fileId) continue;
        try {
          const file = DriveApp.getFileById(fileId);
          const fileSizeKb = file.getSize() / 1024;

          // Spar-Filter: Sehr kleine Dateien (< 15 KB, z.B. Icons, OTP-Mini-Screenshots) überspringen
          if (fileSizeKb < 15) {
            Logger.log(`Zeile ${r + 1}: Datei [${file.getName()}] übersprungen (${fileSizeKb.toFixed(1)} KB < 15 KB).`);
            continue;
          }

          const blob = file.getBlob();
          proofBlobs.push(blob);
          Logger.log(`Zeile ${r + 1}: Beleg [${file.getName()}] (${fileSizeKb.toFixed(1)} KB) für Analyse geladen.`);
        } catch (e) {
          Logger.log(`Zeile ${r + 1}: Datei [${rawUrls[u]}] konnte nicht geladen werden: ` + e.toString());
        }
      }

      // Falls keine verwertbaren Belege übrig bleiben
      if (proofBlobs.length === 0) {
        outputSheet.getRange(r + 1, statusColIdx + 1).setValue("ABGELEHNT_SONSTIGES");
        outputSheet.getRange(r + 1, noteColIdx + 1).setValue("Begründung: Keine verwertbaren Belege (> 15 KB) gefunden.");
        continue;
      }

      try {
        // Zugehörige Original-Rechnung aus Spalte M abrufen
        let invoiceBlob = null;
        const invoiceFileId = extractDriveId(invoiceUrl);
        if (invoiceFileId) {
          try {
            invoiceBlob = DriveApp.getFileById(invoiceFileId).getBlob();
            Logger.log(`Rechnungs-PDF [${invoiceFileId}] geladen.`);
          } catch (e) {
            Logger.log("Konnte Rechnungs-PDF nicht laden: " + e.toString());
          }
        }

        Logger.log(`Analysiere ${proofBlobs.length} Beleg-Datei(en) für Zeile ${r + 1} mit Gemini...`);
        outputSheet.getRange(r + 1, statusColIdx + 1).setValue("IN_PRUEFUNG");

        // 🎯 GEMINI-ANALYSE MIT ALLER GEFILTERTEN BELEGEN + ORIGINAL-RECHNUNG
        const analysisObj = callGeminiForMultiplePaymentProofs(proofBlobs, invoiceBlob, apiKey);
        const finalStatus = analysisObj.status || "ABGELEHNT_SONSTIGES";
        const internalNote = `Betrag: ${analysisObj.betrag || '-'} | Empfänger: ${analysisObj.empfaenger || '-'} | Begründung: ${analysisObj.begruendung || '-'}`;

        outputSheet.getRange(r + 1, statusColIdx + 1).setValue(finalStatus);
        outputSheet.getRange(r + 1, noteColIdx + 1).setValue(internalNote);
        Logger.log(`--> Ergebnis Zeile ${r + 1}: ${finalStatus} (${internalNote})`);
      } catch (err) {
        Logger.log("Fehler bei Verarbeitung: " + err.toString());
        outputSheet.getRange(r + 1, statusColIdx + 1).setValue("ABGELEHNT_SONSTIGES");
        outputSheet.getRange(r + 1, noteColIdx + 1).setValue("Fehler beim Verarbeiten: " + err.toString());
      }
    }
  }
  Logger.log("Gemini-Prüfung beendet.");
}

function callGeminiForMultiplePaymentProofs(proofBlobs, invoiceBlob, apiKey) {
  const fallback = { status: "ABGELEHNT_SONSTIGES", betrag: "-", empfaenger: "-", begruendung: "Analyse unvollständig" };
  if (!apiKey || !proofBlobs || proofBlobs.length === 0) return fallback;

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
  const parts = [];

  // 1. Alle gefilterten Überweisungsbelege anhängen
  proofBlobs.forEach((blob) => {
    parts.push({
      inlineData: {
        mimeType: blob.getContentType() || (blob.getName().toLowerCase().endsWith(".pdf") ? "application/pdf" : "image/jpeg"),
        data: Utilities.base64Encode(blob.getBytes())
      }
    });
  });

  // 2. Original-Rechnung anhängen (falls vorhanden)
  let invoiceInstruction = "";
  if (invoiceBlob) {
    parts.push({
      inlineData: {
        mimeType: invoiceBlob.getContentType() || "application/pdf",
        data: Utilities.base64Encode(invoiceBlob.getBytes())
      }
    });

    invoiceInstruction = `
ABGLEICH MIT ORIGINAL-RECHNUNG:
- Das LETZTE angehängte Dokument ist die ORIGINAL-RECHNUNG.
- Alle davor angehängten Dokumente sind ÜBERWEISUNGSBELEGE / SCREENSHOTS vom Kunden.
- Prüfe, ob mindestens einer der Belege einen überwiesenen Betrag zeigt, der mindestens dem Gesamtrechnungsbetrag entspricht.
- Verwendungszweck-Prüfung IST NICHT ERFORDERLICH. Vergleiche NUR den Gesamtbetrag.
- Falls kein Beleg den vollen Rechnungsbetrag abdeckt, wähle Status "ABGELEHNT_SONSTIGES".
`;
  }

  const promptText = `
Du bist ein Experte für die automatisierte Sicherheitsprüfung von Überweisungsbelegen für L8 Street.
Das HEUTIGE DATUM lautet: **${todayStr}**.
DIE KORREKTE ZIEL-IBAN VON L8 STREET LAUTET: **DE41 6665 0085 0008 9793 32** (DE41666500850008979332).
WICHTIGE ANWEISUNG FÜR MEHRERE DOKUMENTE:
Es wurden insgesamt ${proofBlobs.length} Kundendokumente/Bilder angehängt.
Wenn MINDESTENS EIN DOKUMENT davon einen gültigen, ausgeführten Überweisungsbeleg darstellt, der alle Kriterien erfüllt, gilt die Gesamtprüfung als "ERFOLGREICH". Ignoriere unvollständige Screenshots oder OTP-Codes, sofern ein echter Beleg dabei ist.
PRÜF-KRITERIEN & SICHERHEITSREGELN:
1. ZIEL-IBAN PRÜFUNG:
- Die Ziel-IBAN auf dem Beleg MUSS mit der L8 Street IBAN (${TARGET_IBAN}) übereinstimmen.
- Wir stellen strikt auf die IBAN ab.
${invoiceInstruction}
2. KATEGORIEN FÜR "status":
a) "ERFOLGREICH":
- Mindestens ein Beleg zeigt eine ECHT ausgeführte, gebuchte oder abgeschlossene Überweisung.
- Das Ausführungsdatum entspricht dem heutigen Datum (${todayStr}) oder liegt in der VERGANGENHEIT.
- Das Geld ging an die IBAN ${TARGET_IBAN}.
- Der Betrag deckt die Rechnung vollständig ab (falls Rechnung angehängt).
b) "ABGELEHNT_GEPLANT":
- Das Ausführungsdatum liegt strikt in der ZUKUNFT (nach dem ${todayStr}).
- Die Belege weisen explizit nur auf "geplante", "terminierte" Überweisungen oder unbestätigte Entwürfe hin.
c) "ABGELEHNT_SONSTIGES":
- Keines der Dokumente enthält einen gültigen Beleg für die IBAN ${TARGET_IBAN}.
- Der überwiesene Betrag ist geringer als der Rechnungsbetrag.
- Fälschungsanzeichen: Marketing-Zusätze in IBAN-Feldern, fehlerhaftes Layout, ungleiche Schriftarten.
- Status ist rein "Autorisiert" oder "Vorgemerkt", aber noch NICHT gebucht/ausgeführt.
ANTWORTE AUSSCHLIESSLICH ALS VALIDES JSON:
{"status": "ERFOLGREICH" oder "ABGELEHNT_GEPLANT" oder "ABGELEHNT_SONSTIGES",
"betrag": "Gefundener Betrag mit Währung (vom erfolgreichen Beleg)",
"empfaenger": "Gefundener Empfänger / IBAN",
"begruendung": "Kurze präzise Begründung der Entscheidung"
}`;

  parts.push({ text: promptText });
  const result = executeGeminiRequest(parts, apiKey);
  return result ? { ...fallback, ...result } : fallback;
}

function executeGeminiRequest(parts, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent`;
  const payload = {
    contents: [{ parts: parts }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.1 }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) return null;
    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts[0].text) {
      let rawText = json.candidates[0].content.parts[0].text.trim();
      const jsonMatch = rawText.match(/\[[\s\S]*\]|\{[\s\S]*\}/);
      if (jsonMatch) rawText = jsonMatch[0];
      return safeParseJSON(rawText);
    }
  } catch (e) {
    Logger.log("Fehler bei Gemini API Aufruf: " + e.toString());
  }
  return null;
}

function safeParseJSON(rawText) {
  try {
    return JSON.parse(rawText);
  } catch (e) {
    try {
      let cleanText = rawText
        .replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
        .replace(/\r?\n/g, " ")
        .replace(/",\s*}/g, '"}')
        .replace(/",\s*]/g, '"]');
      return JSON.parse(cleanText);
    } catch (e2) {
      Logger.log("🚨 Sicheres JSON-Parsing fehlgeschlagen: " + e2.toString());
      return null;
    }
  }
}

function extractDriveId(url) {
  if (!url) return null;
  const match = url.match(/(?:id=|\/d\/)([-\w]{25,})/);
  if (match && match[1]) {
    return match[1];
  }
  const fallbackMatch = url.match(/[-\w]{25,}/);
  return fallbackMatch ? fallbackMatch[0] : null;
}