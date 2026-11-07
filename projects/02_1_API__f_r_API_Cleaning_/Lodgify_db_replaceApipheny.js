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

  const baseUrl = "https://api.lodgify.com/v2/reservations/bookings?size=100&includeCount=false&stayFilter=All&updatedSince=2025-06-01T00%3A00%3A00.000%2B01%3A00&includeTransactions=false&includeExternal=false&includeQuoteDetails=false";

  let currentPage = 1;
  let allItems = [];
  let keepGoing = true;

  while (keepGoing) {
    const fullUrlWithPage = baseUrl + "&page=" + currentPage;
    const responseData = sendThroughFixie(fullUrlWithPage, loKey, fixieUrl);

    if (responseData && responseData.items && responseData.items.length > 0) {
      allItems = allItems.concat(responseData.items);
      if (responseData.items.length < 100) {
        keepGoing = false;
      } else {
        currentPage++;
      }
    } else {
      keepGoing = false;
    }
    if (currentPage > 30) keepGoing = false;
  }

  writeToSheet(allItems);
}

function sendThroughFixie(targetUrl, loKey, fixieUrl) {
  const authCredentials = fixieUrl.split('@')[0].replace('http://', '');
  const options = {
    "method": "get",
    "headers": {
      "X-ApiKey": loKey,
      "Proxy-Authorization": "Basic " + Utilities.base64Encode(authCredentials),
      "accept": "application/json"
    },
    "muteHttpExceptions": true
  };
  const response = UrlFetchApp.fetch(targetUrl, options);
  return JSON.parse(response.getContentText());
}

function writeToSheet(items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("db");

  if (!sheet) sheet = ss.insertSheet("db");

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const rows = items
  .filter(item => {
    if (!item.departure || !item.arrival) return false;

    const depDate = new Date(item.departure);
    depDate.setHours(0,0,0,0);

    const arrDate = new Date(item.arrival);
    arrDate.setHours(0,0,0,0);

    const todayCopy = new Date(today); // safety

    const diffDaysDep = Math.round((todayCopy - depDate) / (1000 * 60 * 60 * 24));
    const diffDaysArr = Math.round((todayCopy - arrDate) / (1000 * 60 * 60 * 24));

    const matchesDeparture =
      diffDaysDep <= FILTERS.maxPastDays &&
      diffDaysDep >= -FILTERS.maxFutureDays;

    const matchesArrival =
      diffDaysArr <= 0 &&
      diffDaysArr >= -7;

    if (!(matchesDeparture || matchesArrival)) return false;
    if (item.status !== FILTERS.status) return false;
    if (FILTERS.excludeDeleted && item.is_deleted) return false;

    return true;
  })
    
    .map(item => {
      // Datum ohne Uhrzeit (Mitternacht)
      const arrival = item.arrival ? new Date(new Date(item.arrival).setHours(0,0,0,0)) : "";
      const departure = item.departure ? new Date(new Date(item.departure).setHours(0,0,0,0)) : "";
      const updated = item.updated_at ? new Date(item.updated_at) : "";
      
      const peopleCount = (item.guest_breakdown) ? item.guest_breakdown.adults : (item.rooms && item.rooms[0] ? item.rooms[0].people : "");

      let row = new Array(30).fill(""); 

      row[1]  = String(item.id).replace(/-/g, ""); // B
      row[3]  = arrival;                            // D
      row[4]  = departure;                          // E
      row[5]  = String(item.property_id).replace(/-/g, ""); // F
      row[11] = peopleCount;                        // L
      row[13] = (item.guest) ? item.guest.name : ""; // N
      row[14] = (item.guest) ? item.guest.email : ""; // O
      row[15] = (item.guest) ? item.guest.phone : ""; // P
      row[18] = item.status || "";                  // S
      row[22] = item.source || "";                  // W
      row[26] = updated;                            // AA
      row[29] = item.is_deleted || false;           // AD

      return row;
    });

  // Sortierung nach Abreise (Index 4)
  rows.sort((a, b) => {
    const dateA = a[4];
    const dateB = b[4];
    if (!(dateA instanceof Date)) return 1;
    if (!(dateB instanceof Date)) return -1;
    return dateA - dateB;
  });

  const lastRow = sheet.getLastRow();
  const maxColumns = 30;
  
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, maxColumns).clearContent();
  }

  if (rows.length > 0) {
    const targetRange = sheet.getRange(4, 1, rows.length, maxColumns);
    targetRange.setValues(rows);
    
    // Formatiert Spalte D (Index 4) und E (Index 5) als reines Datum
    // "yyyy-mm-dd" ist das sicherste Format für USA/EN Einstellungen
    sheet.getRange(4, 4, rows.length, 2).setNumberFormat("yyyy-mm-dd");
  }

  Logger.log(rows.length + " Buchungen geschrieben.");
}