function loadWiseHistory2026() {
  const token = (typeof wiseKey !== 'undefined') 
    ? wiseKey 
    : PropertiesService.getScriptProperties().getProperty('wiseKey');

  if (!token) {
    SpreadsheetApp.getUi().alert("Fehler: Kein 'wiseKey' gefunden. Bitte prüfe die Speicherung in Google Apps Script.");
    return;
  }

  const START_DATE = "2026-01-15T00:00:00.000Z";

  try {
    // 1. Profil-ID abrufen
    const profileUrl = "https://api.wise.com/v1/profiles";
    const profileResponse = UrlFetchApp.fetch(profileUrl, {
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });

    const profiles = JSON.parse(profileResponse.getContentText());
    if (!Array.isArray(profiles) || profiles.length === 0 || !profiles[0].id) {
      SpreadsheetApp.getUi().alert("Fehler beim Abrufen des Profils. Bitte prüfe den wiseKey.");
      return;
    }

    const profileId = profiles[0].id;

    // 2. Alle Konten / Guthaben (Balances) abrufen
    const balancesUrl = `https://api.wise.com/v4/profiles/${profileId}/balances?types=STANDARD`;
    const balancesResponse = UrlFetchApp.fetch(balancesUrl, {
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });

    const balances = JSON.parse(balancesResponse.getContentText());

    if (!Array.isArray(balances) || balances.length === 0) {
      SpreadsheetApp.getUi().alert("Keine Guthaben-Konten im Wise-Profil gefunden.");
      return;
    }

    let allTransactions = [];

    // 3. Jedes Währungskonto durchgehen und Kontoauszug abfragen
    balances.forEach(b => {
      const statementUrl = `https://api.wise.com/v3/profiles/${profileId}/borderless-accounts/${b.id}/statement.json?currency=${b.currency}&intervalStart=${START_DATE}`;
      const statementResponse = UrlFetchApp.fetch(statementUrl, {
        headers: { "Authorization": "Bearer " + token },
        muteHttpExceptions: true
      });

      const statementData = JSON.parse(statementResponse.getContentText());

      if (statementData && Array.isArray(statementData.transactions)) {
        statementData.transactions.forEach(t => {
          allTransactions.push({
            date: t.date || "",
            type: t.type || "",
            amount: (t.amount && t.amount.value) !== undefined ? t.amount.value : 0,
            currency: (t.amount && t.amount.currency) || b.currency,
            description: (t.details && t.details.description) || t.title || "",
            payerName: (t.details && t.details.payer && t.details.payer.name) || "",
            merchantName: (t.details && t.details.merchant && t.details.merchant.name) || "",
            postBalance: (t.runningBalance && t.runningBalance.value) !== undefined ? t.runningBalance.value : "",
            reference: (t.details && t.details.paymentReference) || ""
          });
        });
      }
    });

    // 4. Tabellenblatt "wise 26" anlegen oder zurücksetzen
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("wise 26");
    if (!sheet) {
      sheet = ss.insertSheet("wise 26");
    } else {
      sheet.clear();
    }

    // Header-Zeile schreiben
    sheet.appendRow([
      "Datum",
      "Typ",
      "Betrag",
      "Währung",
      "Beschreibung / Händler / Absender",
      "Referenz",
      "Kontostand danach"
    ]);
    sheet.getRange("A1:G1").setFontWeight("bold");

    // 5. Transaktionen chronologisch sortieren und eintragen
    if (allTransactions.length > 0) {
      allTransactions.sort((a, b) => new Date(a.date) - new Date(b.date));

      allTransactions.forEach(t => {
        const details = t.merchantName || t.payerName || t.description;
        sheet.appendRow([
          t.date,
          t.type,
          t.amount,
          t.currency,
          details,
          t.reference,
          t.postBalance
        ]);
      });

      SpreadsheetApp.getUi().alert(`Erfolgreich! Insgesamte Kontobewegungen geladen: ${allTransactions.length}`);
    } else {
      SpreadsheetApp.getUi().alert("Keine Kontobewegungen seit dem 15.01.2026 auf allen Konten gefunden.");
    }

  } catch (err) {
    SpreadsheetApp.getUi().alert("Fehler beim Importieren: " + err.toString());
  }
}