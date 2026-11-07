function importLexofficeContacts() {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('lxKey');
  if (!API_KEY) throw new Error('Missing lexoffice API key');

  const HEADERS = [
    'uuID',
    'Kundennummer',
    'Lieferantennummer',
    'Company name',
    'Salutation',
    'First name',
    'Last name',
    'Street',
    'Address addition',
    'Zip',
    'City',
    'Country',
    'Business Email 1',
    'Business Email 2',
    'Business Email 3'
  ];

  const sheet = SpreadsheetApp.getActive()
    .getSheetByName('lxC') || SpreadsheetApp.getActive().insertSheet('lxC');

  sheet.clearContents();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  let page = 0;
  let row = 2;

  while (true) {
    const res = UrlFetchApp.fetch(
      `https://api.lexoffice.io/v1/contacts?size=250&page=${page++}`,
      { headers: { Authorization: `Bearer ${API_KEY}` } }
    );

    if (res.getResponseCode() !== 200)
      throw new Error(res.getContentText());

    const contacts = JSON.parse(res.getContentText()).content || [];
    if (!contacts.length) break;

    const rows = contacts.map(c => {
      const billing = c.addresses?.billing?.[0] || {};
      const be = c.emailAddresses?.business || [];

      return [
        c.id || '',
        c.roles?.customer?.number || '',
        c.roles?.vendor?.number || '',
        c.company?.name || '',
        c.person?.salutation || '',
        c.person?.firstName || '',
        c.person?.lastName || '',
        billing.street || '',
        billing.supplement || '',
        billing.zip || '',
        billing.city || '',
        billing.countryCode || '',
        be[0] || '',
        be[1] || '',
        be[2] || ''
      ];
    });

    sheet.getRange(row, 1, rows.length, HEADERS.length).setValues(rows);
    row += rows.length;
  }
}
