function listHooks() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("loKey");

  const res = UrlFetchApp.fetch(
    "https://api.lodgify.com/webhooks/v1/list",
    {
      method: "get",
      headers: {
        "X-ApiKey": apiKey,
        "accept": "application/json"
      }
    }
  );

  const data = JSON.parse(res.getContentText());
  Logger.log(JSON.stringify(data, null, 2));
}