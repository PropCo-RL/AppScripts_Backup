/**
 * ZENTRALE FILTER
 */
const FILTERS = {
  maxPastDays: 2,
  maxFutureDays: 7,
  status: "Booked",
  excludeDeleted: true
};

function getLodgifyBookings() {

  const props = PropertiesService.getScriptProperties();

  const loKey = props.getProperty('loKey');
  const fixieUrl = props.getProperty('fixieUrl');

  const baseUrl =
    "https://api.lodgify.com/v2/reservations/bookings?size=100&includeCount=false&stayFilter=All&updatedSince=2025-06-01T00%3A00%3A00.000%2B01%3A00&includeTransactions=false&includeExternal=false&includeQuoteDetails=false";

  let currentPage = 1;
  let allItems = [];
  let keepGoing = true;

  while (keepGoing) {

    const fullUrlWithPage = baseUrl + "&page=" + currentPage;

    const responseData = sendThroughFixie(
      fullUrlWithPage,
      loKey,
      fixieUrl
    );

    if (
      responseData &&
      responseData.items &&
      responseData.items.length > 0
    ) {

      allItems = allItems.concat(responseData.items);

      if (responseData.items.length < 100) {
        keepGoing = false;
      } else {
        currentPage++;
      }

    } else {

      keepGoing = false;

    }

    if (currentPage > 30) {
      keepGoing = false;
    }

  }

  writeToSheet(allItems);

}

function sendThroughFixie(targetUrl, loKey, fixieUrl) {

  const authCredentials =
    fixieUrl.split('@')[0].replace('http://', '');

  const options = {
    method: "get",
    headers: {
      "X-ApiKey": loKey,
      "Proxy-Authorization":
        "Basic " +
        Utilities.base64Encode(authCredentials),
      accept: "application/json"
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(targetUrl, options);

  return JSON.parse(response.getContentText());

}

function writeToSheet(items) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("RES");

  if (!sheet) {
    sheet = ss.insertSheet("RES");
  }

  const today = new Date();

  today.setHours(0, 0, 0, 0);

  let rows = items

    .filter(item => {

      if (!item.departure || !item.arrival) {
        return false;
      }

      const depDate = new Date(
        item.departure.substring(0, 10) +
        "T00:00:00"
      );

      const arrDate = new Date(
        item.arrival.substring(0, 10) +
        "T00:00:00"
      );

      const todayCopy = new Date(today);

      const diffDaysDep = Math.round(
        (todayCopy - depDate) /
        (1000 * 60 * 60 * 24)
      );

      const diffDaysArr = Math.round(
        (todayCopy - arrDate) /
        (1000 * 60 * 60 * 24)
      );

      const matchesDeparture =
        diffDaysDep <= FILTERS.maxPastDays &&
        diffDaysDep >= -FILTERS.maxFutureDays;

      const matchesArrival =
        diffDaysArr <= 0 &&
        diffDaysArr >= -7;

      if (!(matchesDeparture || matchesArrival)) {
        return false;
      }

      if (item.status !== FILTERS.status) {
        return false;
      }

      if (
        FILTERS.excludeDeleted &&
        item.is_deleted
      ) {
        return false;
      }

      return true;

    })

    .map(item => {

      const arrival = item.arrival
        ? new Date(
            item.arrival.substring(0, 10) +
            "T00:00:00"
          )
        : "";

      const departure = item.departure
        ? new Date(
            item.departure.substring(0, 10) +
            "T00:00:00"
          )
        : "";

      const updated = item.updated_at
        ? new Date(item.updated_at)
        : "";

      const peopleCount =
        item.guest_breakdown
          ? item.guest_breakdown.adults
          : (
              item.rooms &&
              item.rooms[0]
            )
            ? item.rooms[0].people
            : "";

      let row = new Array(30).fill("");

      row[0]  = String(item.id).replace(/-/g, "");
      row[1]  = arrival;
      row[2]  = departure;
      row[3]  = String(item.property_id)
                  .replace(/-/g, "");
      row[4]  = peopleCount;
      row[5]  = item.guest
                  ? item.guest.name
                  : "";
      row[6]  = item.guest
                  ? item.guest.email
                  : "";
      row[7]  = item.guest
                  ? item.guest.phone
                  : "";
      row[8]  = item.status || "";
      row[9]  = item.source || "";
      row[10] = updated;
      row[11] = item.is_deleted || false;

      return row;

    });

  rows.sort((a, b) => {

    const dateA = a[2];
    const dateB = b[2];

    if (!(dateA instanceof Date)) {
      return 1;
    }

    if (!(dateB instanceof Date)) {
      return -1;
    }

    return dateA - dateB;

  });

  /**
   * EXTENSION LOGIK
   */

  const finalRows = [];

  rows.forEach(currentRow => {

    const guestName =
      String(currentRow[5] || "").trim();

    if (
      guestName
        .toUpperCase()
        .startsWith("EXT")
    ) {

      const propertyId = currentRow[3];

      const extArrival =
        currentRow[1].getTime();

      const extDeparture =
        currentRow[2];

      const originalBooking =
        finalRows.find(r =>

          r[3] === propertyId &&
          r[2] instanceof Date &&
          r[2].getTime() === extArrival

        );

      if (originalBooking) {

        originalBooking[2] =
          extDeparture;

        Logger.log(
          "EXT merged: " +
          propertyId +
          " extended to " +
          extDeparture
        );

      } else {

        finalRows.push(currentRow);

      }

    } else {

      finalRows.push(currentRow);

    }

  });

  finalRows.sort((a, b) => {

    const dateA = a[2];
    const dateB = b[2];

    if (!(dateA instanceof Date)) {
      return 1;
    }

    if (!(dateB instanceof Date)) {
      return -1;
    }

    return dateA - dateB;

  });

  const lastRow = sheet.getLastRow();

  const maxColumns = 30;

  if (lastRow >= 4) {

    sheet
      .getRange(
        4,
        1,
        lastRow - 3,
        maxColumns
      )
      .clearContent();

  }

  if (finalRows.length > 0) {

    const targetRange = sheet.getRange(
      4,
      1,
      finalRows.length,
      maxColumns
    );

    targetRange.setValues(finalRows);

    sheet
      .getRange(
        4,
        2,
        finalRows.length,
        2
      )
      .setNumberFormat("yyyy-mm-dd");

  }

  Logger.log(
    finalRows.length +
    " Buchungen geschrieben."
  );

}