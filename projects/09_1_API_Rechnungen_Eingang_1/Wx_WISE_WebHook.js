function doPost(e) {
  try {
    const SECRET_KEY = "MeinSicheresPasswort123";
    if (e.parameter && e.parameter.secret && e.parameter.secret !== SECRET_KEY) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "unauthorized" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (e && e.postData && e.postData.contents) {
      const rawJson = e.postData.contents;
      const payload = JSON.parse(rawJson);

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName("Wise");

      if (!sheet) {
        sheet = ss.insertSheet("Wise");
        sheet.appendRow([
          "Datum", 
          "Transfer ID", 
          "Status", 
          "Betrag", 
          "Währung", 
          "Empfänger Name", 
          "Referenz / Verwendungszweck", 
          "Gesamtes JSON (Raw)"
        ]);
        sheet.getRange("A1:H1").setFontWeight("bold");
      }

      // 1. Datum
      let datum = "";
      if (payload.data && payload.data.occurred_at) {
        datum = payload.data.occurred_at.substring(0, 19).replace('T', ' ');
      } else if (payload.sent_at) {
        datum = payload.sent_at.substring(0, 19).replace('T', ' ');
      }

      // 2. Status & Transfer ID
      let status = (payload.data && payload.data.current_state) ? payload.data.current_state : (payload.event_type || "");
      let transferId = "";
      if (payload.data) {
        if (payload.data.resource && payload.data.resource.id) {
          transferId = payload.data.resource.id;
        } else if (payload.data.transfer_id) {
          transferId = payload.data.transfer_id;
        }
      }

      // 3. Wenn es ein Transfer ist: Fehlende Details (Betrag, Name, Referenz) via API nachladen
      let betrag = "";
      let waehrung = "";
      let empfaengerName = "";
      let referenz = "";

      const token = PropertiesService.getScriptProperties().getProperty('wiseKey') || (typeof wiseKey !== 'undefined' ? wiseKey : null);

      if (transferId && token && payload.event_type && payload.event_type.startsWith("transfers#")) {
        try {
          const detailRes = UrlFetchApp.fetch(`https://api.wise.com/v1/transfers/${transferId}`, {
            headers: { "Authorization": "Bearer " + token },
            muteHttpExceptions: true
          });

          if (detailRes.getResponseCode() === 200) {
            const details = JSON.parse(detailRes.getContentText());
            betrag = details.targetValue || details.sourceValue || "";
            waehrung = details.targetCurrency || details.sourceCurrency || "";
            referenz = (details.details && details.details.reference) || details.reference || "";
            
            // Empfängernamen über targetAccount abrufen
            if (details.targetAccount) {
              const accRes = UrlFetchApp.fetch(`https://api.wise.com/v2/accounts/${details.targetAccount}`, {
                headers: { "Authorization": "Bearer " + token },
                muteHttpExceptions: true
              });
              if (accRes.getResponseCode() === 200) {
                const accDetails = JSON.parse(accRes.getContentText());
                empfaengerName = accDetails.accountHolderName || "";
              }
            }
          }
        } catch (apiErr) {
          // Falls API-Call fehlschlägt, läuft Skript trotzdem durch
        }
      }

      // 4. Zeile im Sheet anfügen
      sheet.appendRow([
        datum,
        transferId,
        status,
        betrag,
        waehrung,
        empfaengerName,
        referenz,
        rawJson
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({ "status": "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "ok", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}