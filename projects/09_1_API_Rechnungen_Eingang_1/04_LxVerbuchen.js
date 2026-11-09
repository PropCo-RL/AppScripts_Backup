const LX_KEY_VERBUCHUNG = PropertiesService.getScriptProperties().getProperty('lxKey');

function postCheckedVouchersToLexoffice() {
  const TEST_MODE = false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLog = ss.getSheetByName("Belege");
  const sheetAmz = ss.getSheetByName("AMZ");
  const sheetContacts = ss.getSheetByName("Kontakte");

  // Kontakte-Map aus dem Tab 'Kontakte' laden (Spalte A = Name, Spalte B = UUID)
  const contactsMap = verbuchung_getContactsMap(sheetContacts);

  const sheets = [sheetLog, sheetAmz];

  sheets.forEach(sheet => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    if (sheet.getRange(1, 14).getValue() === "") {
      sheet.getRange(1, 14).setValue("Zahlungsstatus").setFontWeight("bold");
    }

    const range = sheet.getRange(2, 1, lastRow - 1, 14);
    const data = range.getValues();

    let processedCountInThisSheet = 0;

    for (let index = 0; index < data.length; index++) {
      const row = data[index];
      const rowIndex = index + 2;
      const voucherId = String(row[6]).trim();
      const statusM = String(row[12]).trim();

      if (!voucherId) continue;

      // STEP 1: Noch nicht verbuchte Belege nach Lexoffice übertragen
      if (statusM === "") {
        Logger.log(`[TEST-MODE: ${TEST_MODE}] Starte Verbuchung für Zeile ${rowIndex} im Tab '${sheet.getName()}': ${row[1]} (${row[2]})`);

        // Rate-Limit-Schutz: 500 ms Pause vor API-Anfrage
        Utilities.sleep(500);

        const voucherDetails = verbuchung_getVoucherDetails(voucherId);
        if (!voucherDetails) {
          Logger.log(`Konnte Details für Beleg ${voucherId} nicht abrufen. Überspringe.`);
          continue;
        }

        const companyName = String(row[1]).trim();
        const contactKey = companyName.toLowerCase();

        let useCollective = true;
        let contactIdParam = null;

        if (contactsMap[contactKey]) {
          useCollective = false;
          contactIdParam = contactsMap[contactKey];
        }

        // Dynamischer Belegtyp aus Spalte E (purchaseinvoice oder purchasecreditnote)
        const voucherTypeFromSheet = String(row[4]).trim() || "purchaseinvoice";

        const gross = Math.abs(Number(row[7])) || 0;
        const taxPercent = Number(row[5]) || 0;
        const categoryId = String(row[11]).trim();

        // GEÄNDERT: MwSt. und Netto werden jetzt exakt aus Brutto & Steuersatz berechnet (verhindert invalid_taxamount)
        let net = taxPercent > 0 ? gross / (1 + taxPercent / 100) : gross;
        let taxAmount = Math.round((gross - net) * 100) / 100;

        // FIX FÜR INNERGEMEINSCHAFTLICHEN ERWERB:
        if (categoryId === "dba64b05-85f7-4359-8e83-a2556a62eeac") {
          taxAmount = 0;
        }

        const payload = {
          version: voucherDetails.version,
          type: voucherTypeFromSheet,
          voucherStatus: "open",
          voucherNumber: String(row[2]) || "OHNE-NR",
          voucherDate: verbuchung_formatDateToISO(row[0]),
          dueDate: verbuchung_formatDateToISO(row[10]),
          totalGrossAmount: gross,
          totalTaxAmount: taxAmount,
          taxType: "gross",
          useCollectiveContact: useCollective,
          files: voucherDetails.files || [],
          voucherItems: [
            {
              amount: gross,
              taxAmount: taxAmount,
              taxRatePercent: taxPercent,
              categoryId: categoryId
            }
          ]
        };

        if (!useCollective && contactIdParam) {
          payload.contactId = contactIdParam;
        } else {
          payload.contactName = companyName;
        }

        // Rate-Limit-Schutz: Pause vor PUT-Anfrage
        Utilities.sleep(500);

        const url = `https://api.lexware.io/v1/vouchers/${voucherId}`;
        const response = UrlFetchApp.fetch(url, {
          method: "put",
          contentType: "application/json",
          headers: { "Authorization": "Bearer " + LX_KEY_VERBUCHUNG, "Accept": "application/json" },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
          const todayFormatted = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
          sheet.getRange(rowIndex, 13).setValue(`X (${todayFormatted})`);
          Logger.log(`SUCCESS: Zeile ${rowIndex} im Tab '${sheet.getName()}' (${companyName}) als ${voucherTypeFromSheet} gebucht!`);
        } else {
          Logger.log(`FEHLER bei Zeile ${rowIndex} (${sheet.getName()}): ` + response.getContentText());
        }

        processedCountInThisSheet++;
        if (TEST_MODE && processedCountInThisSheet >= 1) break;
      }

      // STEP 2: Zahlungsstatus aus Lexoffice abfragen
      else if (statusM.startsWith("X")) {
        Utilities.sleep(300);
        const paymentInfo = verbuchung_getPaymentStatus(voucherId);
        if (paymentInfo) {
          sheet.getRange(rowIndex, 14).setValue(paymentInfo);
        }
      }
    }
  });
}

// --- Hilfsfunktionen für Skript 2 ---

function verbuchung_getContactsMap(sheetContacts) {
  let map = {};
  if (!sheetContacts) return map;
  const lastRow = sheetContacts.getLastRow();
  if (lastRow > 1) {
    const data = sheetContacts.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(row => {
      const name = String(row[0]).trim().toLowerCase();
      const uuid = String(row[1]).trim();
      if (name && uuid) {
        map[name] = uuid;
      }
    });
  }
  return map;
}

function verbuchung_getPaymentStatus(voucherId) {
  try {
    const url = `https://api.lexware.io/v1/payments/${voucherId}`;
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": "Bearer " + LX_KEY_VERBUCHUNG, "Accept": "application/json" },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.paymentStatus === "balanced") {
        const paidDate = data.paidDate ? data.paidDate.substring(0, 10) : "";
        return `Bezahlt ${paidDate ? "(" + paidDate + ")" : ""}`;
      } else if (data.paymentStatus === "openExpense") {
        return `Offen (${data.openAmount} €)`;
      }
    }
  } catch (e) {}
  return "Offen";
}

function verbuchung_getVoucherDetails(voucherId) {
  try {
    const url = `https://api.lexware.io/v1/vouchers/${voucherId}`;
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": "Bearer " + LX_KEY_VERBUCHUNG, "Accept": "application/json" },
      muteHttpExceptions: true
    });
    return JSON.parse(response.getContentText());
  } catch (e) { return null; }
}

function verbuchung_formatDateToISO(dateInput) {
  if (!dateInput) return "";
  try {
    const d = new Date(dateInput);
    if (isNaN(d.getTime())) return String(dateInput).substring(0, 10);
    return Utilities.formatDate(d, "GMT", "yyyy-MM-dd");
  } catch (e) { return ""; }
}