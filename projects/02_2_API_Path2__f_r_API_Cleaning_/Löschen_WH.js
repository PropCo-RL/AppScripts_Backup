function removeAllHooks() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("loKey");
  
  // 1. Liste holen
  const listRes = UrlFetchApp.fetch("https://api.lodgify.com/webhooks/v1/list", {
    method: "get",
    headers: { "X-ApiKey": apiKey, "accept": "application/json" }
  });
  const hooks = JSON.parse(listRes.getContentText());

  // 2. Jeden Hook löschen
  hooks.forEach(hook => {
    const delRes = UrlFetchApp.fetch("https://api.lodgify.com/webhooks/v1/unsubscribe", {
      method: "delete",
      headers: { 
        "X-ApiKey": apiKey, 
        "accept": "application/json",
        "content-type": "application/json"
      },
      payload: JSON.stringify({ id: hook.id }) // Hier wird die ID gesendet
    });
    Logger.log("Gelöscht: " + hook.event + " (ID: " + hook.id + ") → " + delRes.getResponseCode());
  });
}