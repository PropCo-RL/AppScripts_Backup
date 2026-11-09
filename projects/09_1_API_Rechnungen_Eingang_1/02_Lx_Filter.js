/**
 * Hauptfunktion: Bereinigt Belege anhand der Kontakte-Regeln.
 * Kann manuell ausgeführt oder als Trigger (z. B. zeitgesteuert oder vor Wise) genutzt werden.
 */
function aktualisiereBelegeUndKontakte() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const belegeSheet = ss.getSheetByName("Belege");
  const kontakteSheet = ss.getSheetByName("Kontakte");

  if (!belegeSheet || !kontakteSheet) {
    SpreadsheetApp.getUi().alert("Tabellenblatt 'Belege' oder 'Kontakte' nicht gefunden.");
    return;
  }

  const kontaktRegeln = ladeKontaktRegelnAusSheet(kontakteSheet);
  wendeRegelnAufBelegeAn(belegeSheet, kontaktRegeln);
}

/**
 * Liest Tab 'Kontakte' ein:
 * Spalte A: Name (Index 0)
 * Spalte F: SEPA (Index 5)
 * Spalte G: Bestätigte IBAN (Index 6)
 * Spalte K: Lösch-Markierung 'x' (Index 10)
 */
function ladeKontaktRegelnAusSheet(kontakteSheet) {
  const daten = kontakteSheet.getDataRange().getValues();
  const regeln = new Map();

  for (let i = 1; i < daten.length; i++) {
    const zeile = daten[i];
    const name = String(zeile[0] || "").trim().toLowerCase();
    if (!name) continue;

    const sepa = String(zeile[5] || "").trim().toUpperCase();
    const ibanBestaetigt = String(zeile[6] || "").replace(/\s+/g, '');
    const loeschX = String(zeile[10] || "").trim().toLowerCase();

    regeln.set(name, {
      loeschen: loeschX === "x",
      keineZahlung: sepa === "SEPA" && ibanBestaetigt === ""
    });
  }

  return regeln;
}

/**
 * Verarbeitet das Tab 'Belege':
 * - Zeilen mit 'x' bei Kontakt werden komplett entfernt
 * - Bei SEPA ohne bestätigte IBAN wird in Spalte Q "Keine Zahlung Notwendig" gesetzt
 */
function wendeRegelnAufBelegeAn(belegeSheet, kontaktRegeln) {
  const letzteZeile = belegeSheet.getLastRow();
  if (letzteZeile <= 1) return;

  const anzahlSpalten = Math.max(belegeSheet.getLastColumn(), 17); // Mindestens bis Spalte Q
  const daten = belegeSheet.getRange(1, 1, letzteZeile, anzahlSpalten).getValues();

  const kopfzeile = daten[0];
  const bereinigteDaten = [kopfzeile];

  for (let i = 1; i < daten.length; i++) {
    const zeile = daten[i];
    const firma = String(zeile[1] || "").trim().toLowerCase(); // Spalte B: Firma
    const regel = kontaktRegeln.get(firma);

    // 1. Kontakt ist mit 'x' markiert -> Zeile überspringen (wird gelöscht)
    if (regel && regel.loeschen) {
      continue;
    }

    // 2. SEPA aktiv & keine bestätigte IBAN -> Spalte Q setzen
    if (regel && regel.keineZahlung) {
      zeile[16] = "Keine Zahlung Notwendig"; // Spalte Q (Index 16)
    }

    bereinigteDaten.push(zeile);
  }

  // Altes Sheet leeren und bereinigten Stand in einem Rutsch zurückschreiben
  belegeSheet.clearContents();
  belegeSheet.getRange(1, 1, bereinigteDaten.length, bereinigteDaten[0].length).setValues(bereinigteDaten);
}