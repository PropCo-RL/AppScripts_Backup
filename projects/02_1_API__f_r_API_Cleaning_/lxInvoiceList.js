function syncLexwareToLxLog() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("lxLog") || ss.insertSheet("lxLog");

  // 1. Setup Headers if sheet is blank
  if (sh.getLastRow() === 0) {
    sh.appendRow(["RE uuid", "KD uuid", "RE-Nr", "RE-Date", "Created", "Erfolg", "Einleitung", "Name", "Start", "Ende", "Total", "Status", "", "Pos.", "Item", "Menge", "Preis", "Info"]);
    sh.getRange("1:1").setFontWeight("bold");
  }

  const options = {
    headers: { 
      "Authorization": `Bearer ${props.lxKey}`, 
      "Accept": "application/json" 
    },
    muteHttpExceptions: true
  };

  // 2. Fetch the list of invoices (Page 0 and 1 = Top 200 newest)
  let invoiceRefs = [];
  for (let p = 0; p < 2; p++) {
    const listUrl = `https://api.lexware.io/v1/voucherlist?voucherType=invoice&voucherStatus=any&size=100&page=${p}&sort=voucherdate,DESC`;
    const resp = JSON.parse(UrlFetchApp.fetch(listUrl, options).getContentText());
    if (!resp.content?.length) break;
    
    // Filter out drafts and voided to save on detail calls
    const valid = resp.content.filter(v => !["draft", "voided"].includes(v.voucherStatus));
    invoiceRefs = invoiceRefs.concat(valid);
    if (resp.last) break;
  }

  // 3. Process details one by one (Defensive approach to avoid Bandwidth errors)
  const rows = [];

  for (let i = 0; i < invoiceRefs.length; i++) {
    const v = invoiceRefs[i];
    const url = `https://api.lexware.io/v1/invoices/${v.id}`;
    
    try {
      const res = UrlFetchApp.fetch(url, options);
      
      if (res.getResponseCode() !== 200) {
        rows.push([v.id, v.contactId || "", "", "", "", "FAIL", "", "", "", "", 0, v.voucherStatus, "", "", "", "", "", `ERR: ${res.getResponseCode()}`]);
      } else {
        const inv = JSON.parse(res.getContentText());
        const sc = inv.shippingConditions || {};
        
        const base = [
          inv.id, v.contactId || "", inv.voucherNumber || "",
          new Date(inv.voucherDate), new Date(inv.createdDate),
          "IMPORT", inv.introduction || "", inv.address?.name || "",
          sc.shippingDate ? new Date(sc.shippingDate) : "",
          sc.shippingEndDate ? new Date(sc.shippingEndDate) : "",
          inv.totalPrice?.totalGrossAmount || 0, inv.voucherStatus, ""
        ];

        (inv.lineItems || []).forEach((item, pos) => {
          rows.push([...base, pos + 1, item.name || "", item.quantity || 0, item.unitPrice?.netAmount || 0, ""]);
        });
      }
    } catch (e) {
      // Catch network or bandwidth errors specifically
      rows.push([v.id, v.contactId || "", "", "", "", "FAIL", "", "", "", "", 0, v.voucherStatus, "", "", "", "", "", `API ERROR: ${e.message}`]);
    }

    // Pause for 800ms after every single request (Safe rate: ~1.2 requests/sec)
    Utilities.sleep(800); 
  }

  // 4. Overwrite Sheet
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 18).clearContent();
  }

  if (rows.length) {
    sh.getRange(2, 1, rows.length, 18).setValues(rows);
    // Format Date Columns (D, E and I, J)
    sh.getRange(2, 4, rows.length, 2).setNumberFormat("dd.MM.yyyy");
    sh.getRange(2, 9, rows.length, 2).setNumberFormat("dd.MM.yyyy");
  }
  
  console.log("Sync abgeschlossen. Verarbeitete Rechnungen: " + invoiceRefs.length);
}