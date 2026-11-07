/**
 * Universelle Funktion, um Anfragen über Fixie an Lodgify zu senden
 */
function sendLodgifyRequest(endpoint) {
  const props = PropertiesService.getScriptProperties();
  const loKey = props.getProperty('loKey');
  const fixieUrl = props.getProperty('fixieUrl');

  const targetUrl = "https://api.lodgify.com" + endpoint;
  const authCredentials = fixieUrl.split('@')[0].replace('http://', '');

  const options = {
    "method": "get",
    "headers": {
      "X-ApiKey": loKey,
      "Proxy-Authorization": "Basic " + Utilities.base64Encode(authCredentials),
      "Accept-Language": "de"
    },
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(targetUrl, options);
  return JSON.parse(response.getContentText());
}