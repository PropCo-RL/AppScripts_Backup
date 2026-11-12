/**
 * Hauptfunktion: Aktualisiert das Blatt 'Heizung_Status'.
 * Fragt alle hinterlegten Accounts ab, listet KH100-Hubs und deren KE100-Thermostate
 * nach der python-kasa SMART-Spezifikation auf.
 */
function aktualisiereKasaDaten() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Heizung_Status");
  if (!sheet) {
    sheet = ss.insertSheet("Heizung_Status");
  }

  const headers = [
    "Zuletzt geprüft",
    "Account (E-Mail)",
    "Gerätename / Standort",
    "Modell",
    "Status",
    "Ist-Temp (°C)",
    "Soll-Temp (°C)",
    "Batterie"
  ];

  sheet.clearContents();
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);

  // 1. Script Properties auslesen
  const scriptProperties = PropertiesService.getScriptProperties().getProperties();
  const accounts = [];

  for (const [key, password] of Object.entries(scriptProperties)) {
    if (key.startsWith("Kasa-")) {
      accounts.push({ email: key.substring(5).trim(), pass: password });
    }
  }

  if (accounts.length === 0) {
    Logger.log("Keine Properties mit dem Präfix 'Kasa-' gefunden.");
    return;
  }

  const timestamp = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm:ss");
  const rowsToWrite = [];

  // 2. Alle hinterlegten Accounts abfragen
  accounts.forEach(acc => {
    Logger.log(`Frage Account ab: ${acc.email}...`);
    const auth = getKasaAuth(acc.email, acc.pass);
    if (!auth || !auth.token) {
      Logger.log(`Überspringe ${acc.email} wegen Login-Fehler.`);
      return;
    }

    const token = auth.token;
    const baseUrl = auth.serverUrl || "https://eu-wap.tplinkcloud.com";

    // Geräteliste abrufen
    const deviceRes = UrlFetchApp.fetch(baseUrl + "?token=" + token, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ method: "getDeviceList" }),
      muteHttpExceptions: true
    });

    const resJson = JSON.parse(deviceRes.getContentText());
    const devices = resJson.result?.deviceList || [];

    devices.forEach(dev => {
      const alias = decodeName(dev.alias);
      const model = dev.deviceModel || dev.deviceType || "Kasa-Gerät";
      const deviceId = dev.deviceId;
      const deviceServer = dev.appServerUrl || baseUrl;

      // Wenn es ein KH100 Hub ist
      if (model.includes("KH100") || dev.deviceType?.includes("HUB")) {
        const subDevices = getHubChildDevicesDirect(deviceServer, token, deviceId);
        
        // Hub ist online, wenn er Untergeräte liefert oder das Status-Flag gesetzt ist
        const isHubOnline = subDevices.length > 0 || dev.status === 1 || dev.deviceState === 1;

        rowsToWrite.push([
          timestamp,
          acc.email,
          alias + " (HUB)",
          model,
          isHubOnline ? "🟢 ONLINE" : "🔴 OFFLINE",
          "-",
          "-",
          "-"
        ]);

        // Thermostate und Sensoren unter dem Hub eintragen
        subDevices.forEach(child => {
          // Name aus nickname (Base64) oder alias
          const childAlias = decodeName(child.nickname || child.alias || child.name || "Thermostat");
          const childModel = child.device_model || child.model || child.type || "KE100";
          
          // Logik aus python-kasa: is_available = status != "offline"
          const isOffline = child.status === "offline" || child.status === 0 || child.status === "0";
          const childOnline = !isOffline;

          let curTemp = "-";
          let trgTemp = "-";
          let batt = "OK";

          // Ist-Temperatur (current_temp)
          if (child.current_temp !== undefined) curTemp = child.current_temp;
          else if (child.current_temperature !== undefined) curTemp = child.current_temperature;

          // Soll-Temperatur (target_temp)
          if (child.target_temp !== undefined) trgTemp = child.target_temp;
          else if (child.target_temperature !== undefined) trgTemp = child.target_temperature;

          // Batterie-Status
          if (child.low_battery !== undefined) {
            batt = child.low_battery ? "⚠️ SCHWACH" : "OK";
          } else if (child.battery_level !== undefined) {
            batt = child.battery_level;
          }

          rowsToWrite.push([
            timestamp,
            acc.email,
            "  ↳ " + childAlias,
            childModel,
            childOnline ? "🟢 ONLINE" : "🔴 OFFLINE",
            curTemp,
            trgTemp,
            batt
          ]);
        });

      } else {
        // Eigenständige Geräte
        const isOnline = dev.status === 1 || dev.deviceState === 1;
        rowsToWrite.push([
          timestamp,
          acc.email,
          alias,
          model,
          isOnline ? "🟢 ONLINE" : "🔴 OFFLINE",
          "-",
          "-",
          "-"
        ]);
      }
    });
  });

  // Alles in das Sheet übertragen
  if (rowsToWrite.length > 0) {
    sheet.getRange(2, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
  }

  Logger.log(`Aktualisierung beendet: ${rowsToWrite.length} Zeilen geschrieben.`);
}

/**
 * Fragt die Untergeräte (KE100) vom KH100 Hub exakt nach der python-kasa Struktur ab
 */
function getHubChildDevicesDirect(serverUrl, token, hubDeviceId) {
  const payload = {
    method: "passthrough",
    params: {
      deviceId: hubDeviceId,
      requestData: JSON.stringify({
        "get_child_device_component_list": null,
        "get_child_device_list": null
      })
    }
  };

  try {
    const res = UrlFetchApp.fetch(serverUrl + "?token=" + token, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const resJson = JSON.parse(res.getContentText());
    if (resJson.result && resJson.result.responseData) {
      let parsed = resJson.result.responseData;
      if (typeof parsed === "string") {
        parsed = JSON.parse(parsed);
      }

      const list = parsed.get_child_device_list?.child_device_list 
                || parsed.child_device_list 
                || [];

      if (Array.isArray(list)) {
        return list;
      }
    }
  } catch (e) {
    Logger.log("Fehler bei Hub-Untergeräten: " + e.message);
  }
  return [];
}

/**
 * Login und Ermittlung der regionalen Cloud-URL (mit appVer 3.14.0)
 */
function getKasaAuth(email, password) {
  const payload = {
    method: "login",
    params: {
      appType: "Kasa_Android",
      appVer: "3.14.0",
      cloudUserName: email,
      cloudPassword: password,
      terminalUUID: "b8f05e46-1234-4567-89ab-cdef01234567"
    }
  };

  try {
    let res = UrlFetchApp.fetch("https://eu-wap.tplinkcloud.com", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    let data = JSON.parse(res.getContentText());

    if (data.error_code !== 0) {
      res = UrlFetchApp.fetch("https://wap.tplinkcloud.com", {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      data = JSON.parse(res.getContentText());
    }

    if (data.error_code === 0) {
      return {
        token: data.result.token,
        serverUrl: data.result.appServerUrl || "https://eu-wap.tplinkcloud.com"
      };
    } else {
      Logger.log(`Login fehlgeschlagen für ${email}: [Code ${data.error_code}] ${data.msg}`);
      return null;
    }
  } catch (e) {
    Logger.log("Netzwerkfehler beim Login: " + e.message);
    return null;
  }
}

/**
 * Base64-Decodierung für Klarnamen und Nicknames
 */
function decodeName(alias) {
  if (!alias) return "Unbenannt";
  try {
    if (/^[A-Za-z0-9+/=]+$/.test(alias) && alias.length % 4 === 0 && alias.length >= 8) {
      const decodedBytes = Utilities.base64Decode(alias);
      const decodedString = Utilities.newBlob(decodedBytes).getDataAsString("UTF-8");
      if (/^[\x20-\x7E\xA0-\xFF\u0100-\u017F\s\w\d.,/()_-]+$/.test(decodedString)) {
        return decodedString;
      }
    }
  } catch (e) {}
  return alias;
}