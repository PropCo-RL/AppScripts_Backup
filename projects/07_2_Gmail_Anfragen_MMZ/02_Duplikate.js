/**
 * Bereinigt die MMZ-Tabelle und filtert Duplikate heraus.
 * Erkennt Anfragen desselben Gasts über URL, Name, Zeitraum und Inhalt.
 * Wertet unvollständige Zeilen automatisch auf, sobald bessere Kontaktdaten nachgeliefert werden.
 */
function verarbeiteUndBereinigeMMZ() {
  const SHEET_NAME = "MMZ";
  const LOOKBACK_ROWS = 150; // Prüft die letzten 150 Zeilen

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  sicherstelleHeader(sheet);

  const startRow = Math.max(2, lastRow - LOOKBACK_ROWS + 1);
  const numRows = lastRow - startRow + 1;
  const rangeAB = sheet.getRange(startRow, 1, numRows, 2);
  const daten = rangeAB.getValues();

  // 1. Bereits vorhandene Uniques in Spalte B indexieren (und historische Duplikate bereinigen)
  const indexResult = indexiereVorhandeneUniques(daten);
  let updatesMade = indexResult.bereinigteAltdaten;

  // 2. Offene Zeilen verarbeiten (Spalte A gefüllt, Spalte B leer)
  if (verarbeiteOffeneZeilen(daten, indexResult.index)) {
    updatesMade = true;
  }

  // 3. Nur mit einem einzigen Schreibvorgang zurückschreiben
  if (updatesMade) {
    rangeAB.setValues(daten);
  }
}

/**
 * Stellt sicher, dass Spalte B die korrekte Überschrift und Formatierung besitzt.
 */
function sicherstelleHeader(sheet) {
  const headerB = sheet.getRange(1, 2).getValue();
  if (!headerB || headerB.toString().trim() === "") {
    sheet.getRange(1, 2).setValue("Bereinigt / Unique");
    sheet.getRange(1, 2).setFontWeight("bold");
    sheet.getRange("B:B").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
}

/**
 * Indexiert bereits bestehende Einträge aus Spalte B.
 * Findet und entwertet eventuell vorhandene historische Duplikate direkt in den Bestandsdaten.
 */
function indexiereVorhandeneUniques(daten) {
  const registrierteUrls = new Map();
  const registrierteFingerprints = new Map();
  let bereinigteAltdaten = false;

  for (let i = 0; i < daten.length; i++) {
    const cellB = daten[i][1] ? daten[i][1].toString().trim() : "";
    if (cellB === "" || cellB === "-") continue;

    try {
      const parsed = JSON.parse(cellB);
      const score = getKontaktdatenScore(parsed);
      const fp = generateFingerprint(parsed);
      const url = extrahiereUrl(parsed);

      const matchUrl = url ? registrierteUrls.get(url) : null;
      const matchFp = fp ? registrierteFingerprints.get(fp) : null;

      // Historische Kollision prüfen
      if (matchUrl || matchFp) {
        const maxPrevScore = Math.max(
          matchUrl ? matchUrl.score : -1,
          matchFp ? matchFp.score : -1
        );

        if (score > maxPrevScore) {
          // Bisherige schlechtere Einträge entwerten
          if (matchUrl) daten[matchUrl.index][1] = "-";
          if (matchFp && (!matchUrl || matchFp.index !== matchUrl.index)) {
            daten[matchFp.index][1] = "-";
          }

          // Veraltete Map-Einträge aufräumen
          if (matchUrl && matchUrl.fp) registrierteFingerprints.delete(matchUrl.fp);
          if (matchUrl && matchUrl.url) registrierteUrls.delete(matchUrl.url);
          if (matchFp && matchFp.fp) registrierteFingerprints.delete(matchFp.fp);
          if (matchFp && matchFp.url) registrierteUrls.delete(matchFp.url);

          const finalUrl = url || (matchUrl ? matchUrl.url : "") || (matchFp ? matchFp.url : "");
          const neuerRecord = { index: i, score: score, fp: fp, url: finalUrl };

          if (fp) registrierteFingerprints.set(fp, neuerRecord);
          if (finalUrl) registrierteUrls.set(finalUrl, neuerRecord);
        } else {
          // Zeile i ist ein schlechteres/gleiches historisches Duplikat -> entwerten
          daten[i][1] = "-";

          // Falls ein Match per Fingerprint vorlag, dieser aber noch keine URL hatte -> anreichern
          if (url && matchFp && !matchFp.url) {
            matchFp.url = url;
            registrierteUrls.set(url, matchFp);
          }
        }
        bereinigteAltdaten = true;
      } else {
        // Neuer Unique-Eintrag in den Bestandsdaten
        const neuerRecord = { index: i, score: score, fp: fp, url: url };
        if (fp) registrierteFingerprints.set(fp, neuerRecord);
        if (url) registrierteUrls.set(url, neuerRecord);
      }
    } catch (e) {
      // Fallback für nicht-JSON Zellen
      registrierteFingerprints.set(cellB, { index: i, score: 0, fp: cellB, url: "" });
    }
  }

  return { 
    index: { registrierteUrls, registrierteFingerprints }, 
    bereinigteAltdaten 
  };
}

/**
 * Verarbeitet neue Datensätze (Spalte A gefüllt, Spalte B leer).
 */
function verarbeiteOffeneZeilen(daten, index) {
  let updatesMade = false;

  for (let i = 0; i < daten.length; i++) {
    const cellA = daten[i][0] ? daten[i][0].toString().trim() : "";
    const cellB = daten[i][1] ? daten[i][1].toString().trim() : "";

    if (cellA === "" || cellB !== "") continue;

    try {
      const jsonObj = JSON.parse(cellA);
      const currentScore = getKontaktdatenScore(jsonObj);
      const fp = generateFingerprint(jsonObj);
      const url = extrahiereUrl(jsonObj);

      const matchUrl = url ? index.registrierteUrls.get(url) : null;
      const matchFp = fp ? index.registrierteFingerprints.get(fp) : null;

      if (matchUrl || matchFp) {
        // Höchsten Score der bisherigen Treffer ermitteln
        const maxPrevScore = Math.max(
          matchUrl ? matchUrl.score : -1,
          matchFp ? matchFp.score : -1
        );

        if (currentScore > maxPrevScore) {
          // UPGRADE: Neuer Datensatz ist vollständiger

          // 1. Alle bisherigen Trefferzeilen in der Tabelle entwerten
          if (matchUrl) daten[matchUrl.index][1] = "-";
          if (matchFp && (!matchUrl || matchFp.index !== matchUrl.index)) {
            daten[matchFp.index][1] = "-";
          }

          // 2. Neue vollständige Zeile setzen
          daten[i][1] = cellA;

          // 3. Veraltete Keys aus beiden Maps entfernen
          if (matchUrl && matchUrl.fp) index.registrierteFingerprints.delete(matchUrl.fp);
          if (matchUrl && matchUrl.url) index.registrierteUrls.delete(matchUrl.url);
          if (matchFp && matchFp.fp) index.registrierteFingerprints.delete(matchFp.fp);
          if (matchFp && matchFp.url) index.registrierteUrls.delete(matchFp.url);

          // 4. URL sichern: Falls die neue Zeile keine hatte, die bekannte alte übernehmen
          const finalUrl = url || (matchUrl ? matchUrl.url : "") || (matchFp ? matchFp.url : "");
          const neuerRecord = { index: i, score: currentScore, fp: fp, url: finalUrl };

          if (fp) index.registrierteFingerprints.set(fp, neuerRecord);
          if (finalUrl) index.registrierteUrls.set(finalUrl, neuerRecord);
        } else {
          // DUPLIKAT: Datensatz ist gleichwertig oder unvollständiger
          daten[i][1] = "-";

          // Falls ein Match per Fingerprint vorlag, dieser aber noch keine URL hatte -> anreichern
          if (url && matchFp && !matchFp.url) {
            matchFp.url = url;
            index.registrierteUrls.set(url, matchFp);
          }
        }
      } else {
        // FALL B: Komplett neuer Datensatz
        daten[i][1] = cellA;
        const neuerRecord = { index: i, score: currentScore, fp: fp, url: url };

        if (fp) index.registrierteFingerprints.set(fp, neuerRecord);
        if (url) index.registrierteUrls.set(url, neuerRecord);
      }

      updatesMade = true;
    } catch (e) {
      // Ungültiges JSON in Spalte A
      daten[i][1] = "-";
      updatesMade = true;
    }
  }

  return updatesMade;
}

/**
 * Extrahiert die Anfrage-URL aus dem JSON-Objekt.
 */
function extrahiereUrl(item) {
  const url = item.anfrageUrl || item.anfrage_url || item.url || "";
  return String(url).trim();
}

/**
 * Bewertet den Vollständigkeitsgrad der Kontaktdaten:
 * 2 = Telefon UND E-Mail unmaskiert vorhanden
 * 1 = Mindestens E-Mail ODER Telefon vorhanden
 * 0 = Keine brauchbaren Kontaktdaten (nur Maskierung oder leer)
 */
function getKontaktdatenScore(item) {
  const phone = String(item.telefonnummer || "").trim();
  const email = String(item.email || "").trim();

  const hatPhone = phone !== "" &&
    !phone.includes("••") &&
    !phone.includes("&bull;") &&
    phone.replace(/[^\d]/g, "").length >= 6;

  const hatEmail = email !== "" &&
    !email.includes("••") &&
    !email.includes("&bull;") &&
    email.includes("@");

  if (hatPhone && hatEmail) return 2;
  if (hatPhone || hatEmail) return 1;
  return 0;
}

/**
 * Erzeugt einen stabilen Fingerprint für MMZ-Anfragen.
 * Gibt einen leeren String zurück, falls keine Merkmale vorhanden sind.
 */
function generateFingerprint(item) {
  const rawName = String(item.gastName || "").toLowerCase();
  const cleanName = rawName
    .replace(/^(frau|herr)\s+/i, "")
    .replace(/[^a-z0-9äöüß]/g, "");

  const period = String(item.zeitraum || "").replace(/[^0-9]/g, "");
  const persons = String(item.personen || "").trim();

  const cleanMsg = String(item.nachricht || "")
    .toLowerCase()
    .replace(/[^a-z0-9äöüß]/g, "")
    .substring(0, 60);

  if (!cleanName && !period && !cleanMsg) {
    return "";
  }

  return `${cleanName}|${period}|${persons}|${cleanMsg}`;
}