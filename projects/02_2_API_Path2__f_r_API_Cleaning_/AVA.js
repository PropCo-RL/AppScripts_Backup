function getLodgifyAvailabilityAll() {

  const props = PropertiesService.getScriptProperties();

  const loKey = props.getProperty('loKey');
  const fixieUrl = props.getProperty('fixieUrl');

  if (!loKey || !fixieUrl) {

    Logger.log("FEHLER: 'loKey' oder 'fixieUrl' fehlen!");
    return;

  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  // --- 1. PROPERTY-DATEN HOLEN ---

  const dSheet = ss.getSheetByName("d");

  const propertyInfo = {};

  if (dSheet) {

    const dData = dSheet.getDataRange().getValues();

    for (let i = 1; i < dData.length; i++) {

      const id = String(dData[i][0]).trim();

      // Neue Spaltenstruktur
      const name = dData[i][2];      // Apartment
      const cleaner = dData[i][4];   // Zuordnung

      if (id) {

        propertyInfo[id] = {
          name: name || "Unbekannt",
          cleaner: cleaner || "-"
        };

      }

    }

  }

  // --- 2. API CALL VORBEREITUNG ---

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const startDate = new Date(today);
  startDate.setDate(startDate.getDate() - 2);

  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() + 30);

  // Long-Term Ausschluss
  const eightDaysFuture = new Date(today);
  eightDaysFuture.setDate(eightDaysFuture.getDate() + 8);

  // Filterfenster
  const filterStart = new Date(today);

  const filterEnd = new Date(today);
  filterEnd.setDate(filterEnd.getDate() + 7);
  filterEnd.setHours(23, 59, 59, 999);

  const startStr = Utilities.formatDate(
    startDate,
    "UTC",
    "yyyy-MM-dd'T'00:00:00'Z'"
  );

  const endStr = Utilities.formatDate(
    endDate,
    "UTC",
    "yyyy-MM-dd'T'23:59:59'Z'"
  );

  const targetUrl =
    `https://api.lodgify.com/v2/availability?start=${encodeURIComponent(startStr)}&end=${encodeURIComponent(endStr)}&includeDetails=true`;

  const responseData = sendThroughFixie(
    targetUrl,
    loKey,
    fixieUrl
  );

  if (!responseData) return;

  const items = Array.isArray(responseData)
    ? responseData
    : [responseData];

  // --- 3. DATEN VORSORTIEREN ---

  let rawBookings = [];
  let rawBlocks = [];

  items.forEach(item => {

    const propId = item.property_id
      ? item.property_id.toString()
      : "";

    // Ausschluss
    if (propId === "502123") return;

    if (
      item.periods &&
      Array.isArray(item.periods)
    ) {

      // Long-Term Check
      const isLongTermBlocked =
        item.periods.some(p => {

          if (
            p.available === 0 &&
            p.bookings &&
            p.bookings.length > 0
          ) {

            const pStart = new Date(p.start);

            let realEnd = new Date(p.end);
            realEnd.setDate(
              realEnd.getDate() + 1
            );

            return (
              pStart <= today &&
              realEnd >= eightDaysFuture
            );

          }

          return false;

        });

      if (isLongTermBlocked) return;

      item.periods.forEach(p => {

        if (p.available === 0) {

          let dStart = new Date(p.start);
          dStart.setHours(0, 0, 0, 0);

          let dEnd = new Date(p.end);
          dEnd.setHours(0, 0, 0, 0);

          // Lodgify Checkout/Reinigungskorrektur
          dEnd.setDate(dEnd.getDate() + 1);

          const entry = {

            propId: propId,

            start: dStart,
            end: dEnd,

            bookingId:
              p.bookings &&
              p.bookings.length > 0
                ? p.bookings[0].id
                : "-",

            closedId:
              p.closed_period &&
              p.closed_period.id
                ? p.closed_period.id
                : "-"

          };

          if (entry.bookingId !== "-") {

            rawBookings.push(entry);

          } else if (entry.closedId !== "-") {

            rawBlocks.push(entry);

          }

        }

      });

    }

  });

  // --- 4. ZUSAMMENFÜHREN ---

  let tempRows = [];

  rawBookings.forEach(b => {

    // Nur relevante Checkouts
    if (
      b.end >= filterStart &&
      b.end <= filterEnd
    ) {

      // Passenden Block finden
      const matchingBlock = rawBlocks.find(bl => {

        return (
          bl.propId === b.propId &&
          bl.start.getFullYear() === b.end.getFullYear() &&
          bl.start.getMonth() === b.end.getMonth() &&
          bl.start.getDate() === b.end.getDate()
        );

      });

      let cleaningUntil = "-";

      if (matchingBlock) {

        cleaningUntil =
          Utilities.formatDate(
            matchingBlock.end,
            tz,
            "MM/dd/yyyy"
          );

      }

      const info =
        propertyInfo[b.propId] || {
          name: "Unbekannt (" + b.propId + ")",
          cleaner: "-"
        };

      tempRows.push([

        info.cleaner, // A
        info.name, // B

        Utilities.formatDate(
          b.start,
          tz,
          "MM/dd/yyyy"
        ), // C

        Utilities.formatDate(
          b.end,
          tz,
          "MM/dd/yyyy"
        ), // D

        b.bookingId, // E

        cleaningUntil // F

      ]);

    }

  });

  // --- 5. SORTIERUNG ---

  tempRows.sort((a, b) => {

    const cleanerComp =
      String(a[0]).localeCompare(
        String(b[0])
      );

    if (cleanerComp !== 0) {
      return cleanerComp;
    }

    const parseDate = (dStr) => {

      if (!dStr || dStr === "-") {
        return 0;
      }

      const parts = dStr.split('/');

      return new Date(
        parts[2],
        parts[0] - 1,
        parts[1]
      ).getTime();

    };

    return parseDate(a[3]) - parseDate(b[3]);

  });

  // --- 6. LEERZEILEN BEI CLEANER-WECHSEL ---

  let finalRows = [];

  if (tempRows.length > 0) {

    let lastCleaner = tempRows[0][0];

    tempRows.forEach(row => {

      if (row[0] !== lastCleaner) {

        finalRows.push([
          "",
          "",
          "",
          "",
          "",
          ""
        ]);

        lastCleaner = row[0];

      }

      finalRows.push(row);

    });

  }

  // --- 7. SCHREIBEN IN AVA ---

  let avaSheet =
    ss.getSheetByName("AVA") ||
    ss.insertSheet("AVA");

  const lastRowTotal =
    avaSheet.getLastRow();

  // Alte Daten löschen
  if (lastRowTotal >= 3) {

    avaSheet
      .getRange(
        3,
        1,
        Math.max(
          lastRowTotal - 2,
          finalRows.length
        ),
        6
      )
      .clearContent();

  }

  // Neue Daten schreiben
  if (finalRows.length > 0) {

    avaSheet
      .getRange(
        3,
        1,
        finalRows.length,
        6
      )
      .setValues(finalRows);

    // Datumsformat erzwingen
    avaSheet
      .getRange(
        3,
        3,
        finalRows.length,
        2
      )
      .setNumberFormat("MM/dd/yyyy");

    avaSheet
      .getRange(
        3,
        6,
        finalRows.length,
        1
      )
      .setNumberFormat("MM/dd/yyyy");

    avaSheet.autoResizeColumns(1, 6);

  }

}

function sendThroughFixie(
  targetUrl,
  loKey,
  fixieUrl
) {

  const authCredentials =
    fixieUrl
      .split('@')[0]
      .replace('http://', '');

  const options = {

    method: "get",

    headers: {

      "X-ApiKey": loKey,

      accept: "application/json",

      "Proxy-Authorization":
        "Basic " +
        Utilities.base64Encode(
          authCredentials
        )

    },

    muteHttpExceptions: true

  };

  try {

    const response =
      UrlFetchApp.fetch(
        targetUrl,
        options
      );

    return JSON.parse(
      response.getContentText()
    );

  } catch (e) {

    Logger.log(
      "Fehler: " + e.message
    );

    return null;

  }

}