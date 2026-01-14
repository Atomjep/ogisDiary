function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('GAS実行')
      .addItem('今日のセルに移動', 'fcToday')
      .addItem('サマリメール送信 (OpenAI)', 'summarizeWeekly')
      .addItem('サマリメール送信 (Gemini)', 'summarizeWeeklyGemini')
      .addToUi();
}


function testSendEmail() {
  sendEmail(
    "これはHTMLメールのテストです。\n\n改行も\n反映されますか？\n<b>太字</b>も使えるはずです。", 
    "kouta.ogihara@gmail.com", 
    "HTMLメールテスト"
  );
}

function summarizeMonthly() {
  const today = new Date();
  var api_key = PropertiesService.getScriptProperties().getProperty('DEEPSEEK_API_KEY'); 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Monthly"); // Monthlyシートを指定
  sheet.insertRowBefore(2); //ヘッダーのすぐ下に行を追加する
  
  // 前月の年月を取得して表示
  const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const yearMonth = `${lastMonth.getFullYear()}年${lastMonth.getMonth() + 1}月`;
  sheet.getRange("A2").setValue(yearMonth);
  
  // 最初の指示
  var promptCell 
  = "以下に続く日記の内容から1ヶ月を整理して振り返ってください。\n" +
    "最後に総括として私に対してやる気の出るコメントも添えてください\n"; 
  
  const result = makeDiaryPrompt(today);
  //日記が半分以上記入されていなかったら｢入力日記数不足｣と表示する。
  if(result.inputCellCount/result.totalCells<0.5) {
    sheet.getRange("C2").setValue("入力日記数不足");
    return;
  }

  var prompt = promptCell + result.prompt //前月の1ヶ月分の日記情報を取得する。
  var model = "deepseek-chat"; // 使用するOpenAIモデルのID
  var headers = {
    "Authorization": "Bearer " + api_key,
    "Content-Type": "application/json"
  };
  var data = {
    "model": model,
    "messages": [
      { "role": "user", "content": prompt }
    ],
    "temperature": 0.5,
    "max_tokens": 2048
  };

  var options = {
    "method": "POST",
    "headers": headers,
    "payload": JSON.stringify(data),
    "muteHttpExceptions": true // エラー原因特定のため
  };

  var response = UrlFetchApp.fetch("https://api.deepseek.com/v1/chat/completions", options);
  var json = JSON.parse(response.getContentText());
  if (json.choices && json.choices.length > 0) {
    var generatedText = json.choices[0].message.content;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Monthly"); // Monthlyシートに修正
    sheet.getRange("B2").setValue(generatedText);
  } else {
    sheet.getRange("B2").setValue("エラー: " + JSON.stringify(json));
  }
}

function summarizeWeekly() {
  runWeeklySummary('openai');
}

function summarizeWeeklyGemini() {
  runWeeklySummary('gemini');
}

function runWeeklySummary(modelType) {
  const lastSunday = getLastSunday();
  const lastMonday = new Date(lastSunday);
  
  lastMonday.setDate(lastSunday.getDate() - 6);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly"); // Weeklyシートを指定
  sheet.insertRowBefore(2); //ヘッダーのすぐ下に行を追加する
  sheet.getRange("A2").setValue(`${lastMonday.toLocaleDateString('ja-JP')}`);
  sheet.getRange("B2").setValue(`${lastSunday.toLocaleDateString('ja-JP')}`);

  const sheetMemo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('memo');
  var promptCell = sheetMemo.getRange("B2").getValue();
  
  const result = makeDiaryPrompt(lastMonday, 7);
  //日記が半分以上記入されていなかったら｢入力日記数不足｣と表示する。
  if(result.inputCellCount/result.totalCells<0.5) {
    sheet.getRange("C2").setValue("入力日記数不足");
    return;
  }

  var prompt = promptCell + result.prompt //7日分の日記データを取得する。
  var generatedText = "";
  var instructions = "出力はHTML形式で行ってください。見出しやリストを使って見やすく整形してください。";
  var modelName = "";

  if (modelType === 'gemini') {
    modelName = "gemini-2.5-flash";
    generatedText = callGeminiAPI(prompt, instructions);
  } else {
    // Default to OpenAI/Qwen
    modelName = "qwen/qwen3-max";
    generatedText = callOpenAIAPI(prompt, instructions);
  }

  if (generatedText && !generatedText.startsWith("エラー:")) {
    var finalText = "使用モデル: " + modelName + "\n\n" + generatedText;
    sheet.getRange("C2").setValue(finalText);
    const userEmail = Session.getActiveUser().getEmail();
    var subject = `${lastMonday.toLocaleDateString('ja-JP')}`+"~"+`${lastSunday.toLocaleDateString('ja-JP')}`+"までの1週間振り返り";
    sendEmail(finalText, userEmail, subject, "kouta.ogihara@gmail.com", true);
  } else {
    sheet.getRange("C2").setValue(generatedText);
  }
}

function callOpenAIAPI(prompt, instructions) {
  var api_key = PropertiesService.getScriptProperties().getProperty('QWEN_API_KEY');
  var model = "qwen/qwen3-max"; // 使用するAIモデルのID
  var headers = {
    "Authorization": "Bearer " + api_key,
    "Content-Type": "application/json"
  };
  var data = {
    "model": model,
    "messages": [
      { "role": "user", "content": prompt + "\n\n" + instructions }
    ],
    "temperature": 0.5,
    "max_tokens": 2048
  };

  var options = {
    "method": "POST",
    "headers": headers,
    "payload": JSON.stringify(data),
    "muteHttpExceptions": true // エラー原因特定のため
  };

  try {
    var response = UrlFetchApp.fetch("https://api.novita.ai/v3/openai/chat/completions", options);
    var json = JSON.parse(response.getContentText());
    if (json.choices && json.choices.length > 0) {
      return json.choices[0].message.content;
    } else {
      return "エラー: " + JSON.stringify(json);
    }
  } catch (e) {
    return "エラー: " + e.toString();
  }
}

function callGeminiAPI(prompt, instructions) {
  var api_key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + api_key;
  
  var data = {
    "contents": [{
      "parts": [{"text": prompt + "\n\n" + instructions}]
    }],
    "generationConfig": {
        "temperature": 0.5
    }
  };

  var options = {
    "method": "POST",
    "headers": {
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(data),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
      return json.candidates[0].content.parts[0].text;
    } else {
      return "エラー: " + JSON.stringify(json);
    }
  } catch (e) {
    return "エラー: " + e.toString();
  }
}



//最も近い日曜日を取得する。
function getLastSunday() {
  // const today = new Date();
  const today = new Date("2026-01-07");
  
  // 今日の曜日を取得（0:日曜日, 1:月曜日, ..., 6:土曜日）
  const dayOfWeek = today.getDay();
  
  // 日曜日からの日数を計算
  // 今日が日曜日の場合は7日前の日曜日を返す
  const daysToSubtract = dayOfWeek === 0 ? 7 : dayOfWeek;
  
  // 過去の日曜日の日付を計算
  const recentSunday = new Date(today);
  recentSunday.setDate(today.getDate() - daysToSubtract);
  
  return recentSunday;
}

function makeDiaryPrompt(date, days = null) {
  
  // 日付が指定されていなければ、今日の日付を使用
  if (!date) {
    date = new Date(); // 今日の日付
  }

  const sheetName = '日記'; // シート名を固定
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  let startDate, daysToRetrieve;

  if (days !== null) {
    // 日数が指定された場合：指定日からN日分のデータを取得
    startDate = new Date(date);
    daysToRetrieve = days;
  } else {
    // 日数が指定されなかった場合：1ヶ月前の月の全日分を取得
    const targetYear = date.getMonth() === 0 ? date.getFullYear() - 1 : date.getFullYear();
    const targetMonth = date.getMonth() === 0 ? 11 : date.getMonth() - 1; //1ヶ月前
    startDate = new Date(targetYear, targetMonth, 1);
    
    // その月の日数を取得
    daysToRetrieve = new Date(targetYear, targetMonth + 1, 0).getDate();
  }

  // スプレッドシート上の行位置を決定（1月1日が2行目）
  const startOfYear = new Date(startDate.getFullYear(), 0, 1);
  const dayOffset = Math.floor((startDate - startOfYear) / (1000 * 60 * 60 * 24));
  const targetRow = 2 + dayOffset;

  // F列（列6）から指定日数分のデータを取得
  const contents = sheet
    .getRange(targetRow, 6, daysToRetrieve, 1)
    .getValues(); // 二次元配列 [ [内容], [内容], ... ]

  // 入力のあるセル数をカウント（空文字列、null、undefinedでない場合）
  let inputCellCount = 0;
  contents.forEach(row => {
    if (row[0] !== "" && row[0] !== null && row[0] !== undefined) {
      inputCellCount++;
    }
  });

  // [◯月◯日, 内容] の二次元配列に整形
  const result = contents.map((row, i) => {
    const currentDate = new Date(startDate);
    currentDate.setDate(startDate.getDate() + i);
    const month = currentDate.getMonth() + 1;
    const day = currentDate.getDate();
    const dateLabel = `${month}月${day}日`;
    return [dateLabel, row[0]];
  });

  // prompt: ◯月◯日:\n内容\n◯月◯日:\n内容... の文字列に変換
  const prompt = result
    .map(([dateLabel, content]) => `${dateLabel}:\n${content}`)
    .join('\n');

  // プロンプト文字列と入力セル数を含むオブジェクトを返す
  const resultObject = {
    prompt: prompt,
    inputCellCount: inputCellCount,
    totalCells: daysToRetrieve
  };

  Logger.log(`プロンプト: ${prompt}`);
  Logger.log(`入力済みセル数: ${inputCellCount}/${daysToRetrieve}`);
  
  return resultObject;
}

function sendEmail(body="test", to="atomjep@gmail.com", subject = "test", from = "kouta.ogihara@gmail.com") {
  try {
    // 入力値の検証
    if (!to || !body) {
      throw new Error("送信先メールアドレスと本文は必須です");
    }
    
    // メールアドレスの簡単な形式チェック
    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailPattern.test(to)) {
      throw new Error("有効なメールアドレスを入力してください");
    }

    // HTMLテンプレートを読み込み
    const template = HtmlService.createTemplateFromFile('email_template');
    
    // テンプレート変数を設定
    // 改行を<br>に変換してHTMLとして安全に埋め込む
    template.body = body.replace(/\n/g, '<br>');
    template.subject = subject;
    
    // HTMLを生成
    const htmlBody = template.evaluate().getContent();
    
    // メール送信のオプション
    const options = {
      htmlBody: htmlBody,
    };
    
    // 送信者が指定されている場合は追加
    if (from && emailPattern.test(from)) {
      options.replyTo = from;
    }
    
    // メール送信
    MailApp.sendEmail(to, subject, "", options); // body引数は空にしてoptions.htmlBodyを使用
    
    Logger.log(`メール送信完了: ${to}`);
    return { success: true, message: "メール送信完了" };
    
  } catch (error) {
    Logger.log("メール送信エラー:", error.message);
    return { success: false, message: error.message };
  }
}



