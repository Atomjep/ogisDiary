function debugDashScopeConnection() {
  var api_key = PropertiesService.getScriptProperties().getProperty('DEEPSEEK_API_KEY');
  
  if (!api_key) {
    Logger.log("ERROR: 'DEEPSEEK_API_KEY' is not set in Script Properties.");
    return;
  }

  Logger.log("API Key loaded: " + api_key.substring(0, 5) + "..." + api_key.substring(api_key.length - 4));
  Logger.log("Key length: " + api_key.length);

 
  testEndpoint(api_key, "https://api.deepseek.com/v1/chat/completions", "deepseek-chat");
  // testEndpoint(api_key, "https://api.novita.ai/v3/openai/chat/completions", "Novita-ai");
}

function testEndpoint(api_key, url, label) {
  Logger.log("--- Testing " + label + " Endpoint ---");
  Logger.log("URL: " + url);

  var headers = {
    "Authorization": "Bearer " + api_key,
    "Content-Type": "application/json"
  };
  
  var data = {
    "model": "deepseek-chat", // deepseek
    // "model": "qwen/qwen3-max", // Qwen3
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