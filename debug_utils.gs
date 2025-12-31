function debugDashScopeConnection() {
  var api_key = PropertiesService.getScriptProperties().getProperty('QWEN_API_KEY');
  
  if (!api_key) {
    Logger.log("ERROR: 'QWEN_API_KEY' is not set in Script Properties.");
    return;
  }

  Logger.log("API Key loaded: " + api_key.substring(0, 5) + "..." + api_key.substring(api_key.length - 4));
  Logger.log("Key length: " + api_key.length);

  // Test International Endpoint
  testEndpoint(api_key, "https://api.novita.ai/v3/openai/chat/completions", "Novita-ai");
}

function testEndpoint(api_key, url, label) {
  Logger.log("--- Testing " + label + " Endpoint ---");
  Logger.log("URL: " + url);

  var headers = {
    "Authorization": "Bearer " + api_key,
    "Content-Type": "application/json"
  };
  
  var data = {
    "model": "qwen/qwen2.5-vl-72b-instruct", // Qwen2.5-7B
    "messages": [{ "role": "user", "content": "Hello" }]
  };

  var options = {
    "method": "POST",
    "headers": headers,
    "payload": JSON.stringify(data),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("Response Code: " + response.getResponseCode());
    Logger.log("Response Body: " + response.getContentText());
  } catch (e) {
    Logger.log("Exception: " + e.toString());
  }
}