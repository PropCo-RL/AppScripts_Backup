// ==========================================
// SEPARATER GEMINI AI CHECKER FÜR ZAHLUNGSBELEGE (INKL. RECHNUNGSABGLEICH)
// ==========================================

const TARGET_SPREADSHEET_ID = "1VAw-KR7KlxXjjaj-0-iNZ0g-VvKzwg9uKaUIuop71sE";
const TARGET_IBAN = "DE41666500850008979332"; // L8 Street GmbH Ziel-IBAN (ohne Leerzeichen)

function processNewPaymentProofsWithGemini() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('Gemini-Key');
  if (!apiKey) {
    Logger.log("FEHLER: 'Gemini-Key' nicht in ScriptProperties gefunden!");
    return;
  }

  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const sheetCheckin = ss.getSheetByName("checkin");
  const sheetRE = ss.getSheetByName("RE");
  if (!sheetCheckin) return;

  // Rechnungs-PDFs aus Tab 'RE' indizieren (Lodgify-ID -> PDF URL in Spalte AK / Index 36)
  const invoicePdfMap = {};
  if (sheetRE && sheetRE.getLastRow() >= 2) {
    const reData = sheetRE.getRange(2, 1, sheetRE.getLastRow() - 1, 37).getValues();
    reData.forEach(row => {
      const lodgifyId = String(row[18] || "").trim(); // Spalte S
      const pdfUrl = String(row[36] || "").trim();    // Spalte AK
      if (lodgifyId && pdfUrl) {
        invoicePdfMap[lodgifyId] = pdfUrl;
      }
    });
  }

  const lastRow = sheetCheckin.getLastRow();
  if (lastRow < 2) return;

  const range = sheetCheckin.getRange(2, 1, lastRow - 1, 8);
  const data = range.getValues();

  for (let i = 0; i < data.length; i++) {
    const fileUrl = String(data[i][5] || "").trim(); // Spalte F (Upload Beleg Link)
    const lodgifyId = String(data[i][3] || "").trim(); // Spalte D (Lodgify ID)
    const currentStatus = String(data[i][6] || "").trim(); // Spalte G

    if (fileUrl.includes("drive.google.com") && (currentStatus === "IN_PRUEFUNG" || currentStatus === "Ausstehend (Prüfung läuft)" || currentStatus === "" || currentStatus === "-")) {
      const proofFileId = extractDriveId(fileUrl);
      if (!proofFileId) continue;

      try {
        const proofFile = DriveApp.getFileById(proofFileId);
        const proofBlob = proofFile.getBlob();

        // 1. Zugehörige Original-Rechnung aus Tab RE abrufen
        let invoiceBlob = null;
        const invoiceUrl = invoicePdfMap[lodgifyId] || "";
        const invoiceFileId = extractDriveId(invoiceUrl);

        if (invoiceFileId) {
          try {
            invoiceBlob = DriveApp.getFileById(invoiceFileId).getBlob();
            Logger.log(`Rechnungs-PDF [${invoiceFileId}] für Buchung ${lodgifyId} geladen.`);
          } catch (e) {
            Logger.log("Konnte Rechnungs-PDF nicht laden: " + e.toString());
          }
        }

        Logger.log(`Analysiere Überweisungsbeleg [${proofFile.getName()}] mit Gemini...`);

        // 2. Gemini-Analyse mit Beleg UND Original-Rechnung
        const analysisObj = callGeminiForPaymentProof(proofBlob, invoiceBlob, apiKey);

        const finalStatus = analysisObj.status || "ABGELEHNT_SONSTIGES";
        const internalNote = `Betrag: ${analysisObj.betrag || '-'} | Empfänger: ${analysisObj.empfaenger || '-'} | Begründung: ${analysisObj.begruendung || '-'}`;

        sheetCheckin.getRange(i + 2, 7).setValue(finalStatus);
        sheetCheckin.getRange(i + 2, 8).setValue(internalNote);

        Logger.log(`--> Ergebnis für ${proofFile.getName()}: ${finalStatus} (${internalNote})`);

      } catch (err) {
        Logger.log("Fehler bei Verarbeitung: " + err.toString());
        sheetCheckin.getRange(i + 2, 7).setValue("ABGELEHNT_SONSTIGES");
        sheetCheckin.getRange(i + 2, 8).setValue("Fehler beim Verarbeiten: " + err.toString());
      }

      Utilities.sleep(300);
    }
  }
}

function callGeminiForPaymentProof(proofBlob, invoiceBlob, apiKey) {
  const fallback = { status: "ABGELEHNT_SONSTIGES", betrag: "-", empfaenger: "-", begruendung: "Analyse unvollständig" };
  if (!apiKey || !proofBlob) return fallback;

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
  const parts = [];

  // 1. Überweisungsbeleg anhängen
  parts.push({
    inlineData: {
      mimeType: proofBlob.getContentType() || (proofBlob.getName().endsWith(".pdf") ? "application/pdf" : "image/jpeg"),
      data: Utilities.base64Encode(proofBlob.getBytes())
    }
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
- Es wurden ZWEI Dokumente angehängt: Das erste ist der ÜBERWEISUNGSBELEG, das zweite ist die ORIGINAL-RECHNUNG.
- Prüfe, ob der überwiesene Betrag auf dem Beleg mindestens dem Gesamtrechnungsbetrag entspricht.
- Verwendungszweck-Prüfung IST NICHT ERFORDERLICH. Vergleiche NUR den Gesamtbetrag.
- Falls der Betrag auf dem Beleg geringer ist als der Rechnungsbetrag, wähle Status "ABGELEHNT_SONSTIGES".
`;
  }

  const promptText = `
Du bist ein Experte für die automatisierte Sicherheitsprüfung von Überweisungsbelegen für L8 Street.
Das HEUTIGE DATUM lautet: **${todayStr}**.
DIE KORREKTE ZIEL-IBAN VON L8 STREET LAUTET: **DE41 6665 0085 0008 9793 32** (DE41666500850008979332).

PRÜF-KRITERIEN & SICHERHEITSREGELN:

1. ZIEL-IBAN PRÜFUNG:
   - Die Ziel-IBAN auf dem Beleg MUSS mit der L8 Street IBAN (${TARGET_IBAN}) übereinstimmen.
   - Wir stellen strikt auf die IBAN ab.

${invoiceInstruction}

2. KATEGORIEN FÜR "status":
   a) "ERFOLGREICH":
      - Der Beleg zeigt eine ECHT ausgeführte, gebuchte oder abgeschlossene Überweisung.
      - Das Ausführungsdatum entspricht dem heutigen Datum (${todayStr}) oder liegt in der VERGANGENHEIT.
      - Das Geld ging an die IBAN ${TARGET_IBAN}.
      - Der Betrag deckt die Rechnung vollständig ab (falls Rechnung angehängt).

   b) "ABGELEHNT_GEPLANT":
      - Das Ausführungsdatum liegt strikt in der ZUKUNFT (nach dem ${todayStr}).
      - Der Beleg weist explizit auf eine "geplante", "terminierte" Überweisung oder einen unbestätigten Entwurf hin.

   c) "ABGELEHNT_SONSTIGES":
      - Die Empfänger-IBAN passt NICHT zu ${TARGET_IBAN}.
      - Der überwiesene Betrag ist geringer als der Rechnungsbetrag.
      - Fälschungsanzeichen: Marketing-Zusätze in IBAN-Feldern, fehlerhaftes Layout, ungleiche Schriftarten.
      - Status ist rein "Autorisiert" oder "Vorgemerkt", aber noch NICHT gebucht/ausgeführt.

ANTWORTE AUSSCHLIESSLICH ALS VALIDES JSON:
{
  "status": "ERFOLGREICH" oder "ABGELEHNT_GEPLANT" oder "ABGELEHNT_SONSTIGES",
  "betrag": "Gefundener Betrag mit Währung",
  "empfaenger": "Gefundener Empfänger / IBAN",
  "begruendung": "Kurze präzise Begründung der Entscheidung"
}
  `;

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
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}