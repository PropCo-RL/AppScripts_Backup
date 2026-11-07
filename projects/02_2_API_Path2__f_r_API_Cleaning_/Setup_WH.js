function subscribeFresh() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("loKey");

  const base = "https://script.google.com/macros/s/AKfycbzosYQsPYFcoIAe1oLRZh02VLCrnxkkAoABbYVPhZo0rhpnA4tRGyY_bjQCIh1l-oES/exec";

  const webhooks = [
    { event: "booking_new_any_status", url: base + "?type4=booking_new" },
  ];

  webhooks.forEach(hook => {
    const res = UrlFetchApp.fetch(
      "https://api.lodgify.com/webhooks/v1/subscribe",
      {
        method: "post",
        headers: {
          "X-ApiKey": apiKey,
          "content-type": "application/json",
          "accept": "application/json"
        },
        payload: JSON.stringify({
          target_url: hook.url,
          event: hook.event
        }),
        muteHttpExceptions: true
      }
    );

    Logger.log(hook.event + " → " + res.getContentText());
  });
}