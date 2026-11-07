function executeAllTasks() {
  const s = SpreadsheetApp.getActive().getSheetByName("e");
  const ss = SpreadsheetApp.getActive();
  const prop = PropertiesService.getScriptProperties();
  const [lxK, loK] = [prop.getProperty("lxKey"), prop.getProperty("loKey")];
  
  // Optimierte Hilfsfunktion für Lexoffice und Lodgify
  const call = (url, key, p, isLx = true) => {
    const headers = {
      "Content-Type": "application/json",
      "Accept": "application/json"
    };
    
    if (isLx) {
      headers["Authorization"] = "Bearer " + key;
    } else {
      headers["X-ApiKey"] = key;
    }

    const res = UrlFetchApp.fetch(url, {
      method: "post",
      headers: headers,
      payload: (typeof p === 'string') ? p : JSON.stringify(p), 
      muteHttpExceptions: true
    });

    const resText = res.getContentText();
    const resCode = res.getResponseCode();
    
    // Fehlerbehandlung: Lexware gibt 200 oder 201 bei Erfolg zurück
    if (resCode > 201) throw "Fehler bei " + url + " (Status " + resCode + "): " + resText;
    
    return JSON.parse(resText).id;
  };

  try {
    // 1. RECHNUNG (C2) - Nutzt Einmal-Adresse aus deiner neuen Formel
    // ?finalize=true sorgt dafür, dass direkt ein PDF generiert wird
    const reId = call("https://api.lexoffice.io/v1/invoices?finalize=true", lxK, s.getRange("C2").getValue());

    // 2. KONTAKT (C3) - Falls du dennoch parallel einen Kontakt in Lexware anlegen willst
    let kdId = "ONE-TIME"; // Standardwert, falls kein Kontakt erstellt wird
    const c3Value = s.getRange("C3").getValue().toString();
    if (c3Value.includes("?payload=")) {
      const [lxUrl, lxP] = c3Value.split("?payload=");
      kdId = call(lxUrl.replace("lexware.io", "lexoffice.io"), lxK, lxP);
    }

    // 3. LODGIFY (C4)
    const c4Value = s.getRange("C4").getValue().toString();
    if (c4Value.includes("?payload=")) {
      const [loUrl, loP] = c4Value.split("?payload=");
      call(loUrl, loK, loP, false);
    }

    // 4. LOG lxC (C5) -> Spalten A, D, H-M
    const c5Parts = s.getRange("C5").getValue().toString().split("|");
    let rowC = new Array(15).fill(""); 
    rowC[0] = kdId; 
    c5Parts.forEach(p => {
      const val = p.includes(": ") ? p.split(": ")[1].trim() : "";
      if (p.includes("Name:"))   rowC[3] = val;  // D
      if (p.includes("Zusatz:")) rowC[7] = val;  // H
      if (p.includes("Straße:")) rowC[8] = val;  // I
      if (p.includes("PLZ:"))    rowC[9] = val;  // J
      if (p.includes("Stadt:"))  rowC[10] = val; // K
      if (p.includes("Land:"))   rowC[11] = val; // L
      if (p.includes("Email:"))  rowC[12] = val; // M
    });
    ss.getSheetByName("lxC").appendRow(rowC);

    // 5. LOG lxLog (C6) -> Spalten 1 bis 17 (A bis Q)
    const c6Parts = s.getRange("C6").getValue().toString().split("|");
    let rowLog = new Array(17).fill(""); 
    rowLog[0] = reId; // A: RE uuid
    rowLog[1] = kdId; // B: KD uuid (oder "ONE-TIME")
    
    c6Parts.forEach(p => {
      const val = p.includes(": ") ? p.split(": ")[1].trim() : "";
      if (p.includes("RE-Nr:"))      rowLog[2] = val;  
      if (p.includes("RE-Date:"))    rowLog[3] = val;  
      if (p.includes("Einleitung:")) rowLog[6] = val;  
      if (p.includes("Name:"))       rowLog[7] = val;  
      if (p.includes("Start:"))      rowLog[8] = val;  
      if (p.includes("Ende:"))       rowLog[9] = val;  
      if (p.includes("Total:"))      rowLog[10] = val; 
      if (p.includes("Status:"))     rowLog[11] = val; 
      if (p.includes("Item 1:"))     rowLog[14] = val; 
      if (p.includes("Menge:"))      rowLog[15] = val; 
      if (p.includes("Preis:"))      rowLog[16] = val; 
    });
    ss.getSheetByName("lxLog").appendRow(rowLog);

    s.getRange("G2").setValue("Erfolg!").setFontColor("green");

  } catch (e) {
    s.getRange("G2").setValue("Fehler: " + e).setFontColor("red");
    Logger.log(e);
  }
}