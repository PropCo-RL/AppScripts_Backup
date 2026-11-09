/**
 * Hauptfunktion: Liest Lexware-UUIDs aus Spalte B (Tab: Kontakte),
 * ruft die Kontakte per Lexware API ab und schreibt gefundene IBANs
 * gesammelt in Spalte L.
 */
function synchronisiereIbanAusLexware() {
  const tabellenBlatt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kontakte");
  if (!tabellenBlatt) {
    Logger.log("Tabellenblatt 'Kontakte' nicht gefunden.");
    return;
  }

  const apiSchluessel = PropertiesService.getScriptProperties().getProperty("lxKey");
  if (!apiSchluessel) {
    Logger.log("Fehler: 'lxKey' ist nicht in den Script Properties hinterlegt.");
    return;
  }

  const letzteZeile = tabellenBlatt.getLastRow();
  if (letzteZeile < 2) return;

  // Lese Spalte B (Contact ID / UUID) ab Zeile 2
  const uuidBereich = tabellenBlatt.getRange(2, 2, letzteZeile - 1, 1).getValues();
  const ergebnisSpalteL = [];

  for (let i = 0; i < uuidBereich.length; i++) {
    const kontaktUuid = String(uuidBereich[i][0]).trim();

    if (!kontaktUuid || !istGueltigeUuid(kontaktUuid)) {
      ergebnisSpalteL.push([""]);
      continue;
    }

    const gefundeneIban = ladeIbanZuKontakt(kontaktUuid, apiSchluessel);
    ergebnisSpalteL.push([gefundeneIban || ""]);

    // Lexware Rate-Limit: max 2 Requests/Sekunde -> 550ms Pause
    Utilities.sleep(550);
  }

  // Spalte L ist Spalte 12 -> Gebatchtes Schreiben
  tabellenBlatt.getRange(2, 12, ergebnisSpalteL.length, 1).setValues(ergebnisSpalteL);
  Logger.log("Abgleich beendet. Spalte L erfolgreich aktualisiert.");
}

/**
 * Ruft einen Kontakt von der Lexware API ab und extrahiert die IBAN.
 */
function ladeIbanZuKontakt(kontaktUuid, apiSchluessel) {
  const url = "https://api.lexware.io/v1/contacts/" + encodeURIComponent(kontaktUuid);
  const options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + apiSchluessel,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };

  try {
    const antwort = UrlFetchApp.fetch(url, options);
    const status = antwort.getResponseCode();

    if (status === 429) {
      Logger.log("Rate-Limit erreicht bei UUID " + kontaktUuid + ". Warte 2 Sekunden...");
      Utilities.sleep(2000);
      return ladeIbanZuKontakt(kontaktUuid, apiSchluessel);
    }

    if (status !== 200) {
      Logger.log("Fehler bei Abruf von " + kontaktUuid + ": HTTP " + status);
      return null;
    }

    const kontaktJson = JSON.parse(antwort.getContentText());
    return extrahiereIbanAusKontakt(kontaktJson);
  } catch (fehler) {
    Logger.log("Ausnahme bei Abruf von UUID " + kontaktUuid + ": " + fehler.toString());
    return null;
  }
}

/**
 * Sucht nach Bankdaten oder scannt das gesamte JSON / Notizfeld nach einer IBAN.
 */
function extrahiereIbanAusKontakt(kontaktObjekt) {
  if (!kontaktObjekt) return null;

  // 1. Falls Lexware künftig ein bankAccounts-Objekt liefert
  if (kontaktObjekt.bankAccounts && Array.isArray(kontaktObjekt.bankAccounts) && kontaktObjekt.bankAccounts.length > 0) {
    const directIban = kontaktObjekt.bankAccounts[0].iban;
    if (directIban) return directIban.replace(/\s+/g, "").toUpperCase();
  }

  // 2. Suche per Regex im gesamten Kontakt (z. B. Feld note / Notizen)
  const rohText = JSON.stringify(kontaktObjekt);
  const ibanRegex = /[A-Z]{2}[0-9]{2}[A-Z0-9]{11,30}/gi;
  const bereinigterText = rohText.replace(/\s+/g, "");
  const treffer = bereinigterText.match(ibanRegex);

  if (treffer && treffer.length > 0) {
    return treffer[0].toUpperCase();
  }

  return null;
}

/**
 * Validiert das Format einer Standard-UUID.
 */
function istGueltigeUuid(uuid) {
  const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  return uuidRegex.test(uuid);
}