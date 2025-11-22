function rand() {
  
  var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet(); // 現在のSpreadSheetを取得
  var sheet=activeSpreadSheet.getSheetByName('日記'); // シート(SpreadSheetの下のタブ名を指定)

  var ogiRange = sheet.getRange("A1");
  var rand = 0;

  while(rand == 0){// 指定するセルナンバーが０となった時は繰り返し

    rand = Math.random();
    rand = Math.floor(rand*365)+1;
    Logger.log(rand);

    if(rand != 0)
    ogiRange = sheet.getRange(rand,3);

  }
  
  var sheet=activeSpreadSheet.getSheetByName('ランダム表示'); // シート(SpreadSheetの下のタブ名を指定)
  sheet.getRange(3,2).setValue(rand-1);
}

function fcToday() {
  /*Dateオブジェクトで取得した日付をdに代入*/
  var d = new Date(); 

  // 2025年は5列目
  var y = 5 

 // dから月と日だけを抽出する。
  var mon = d.getMonth() + 1;
  var d2 = d.getDate();

  // mm/dd の文字列を生成する。
  var ogitoday = mon +"/" + d2;

  // Logger.log(ogitoday);
  
  var sheet=SpreadsheetApp.getActiveSheet();

  // B列の文字列を配列に格納する。
  let value=sheet.getRange("B2:B366").getValues();

  // 今日の日付と一致する要素番号を探す。
  var i = 0;
  while(value[i] != ogitoday){
    i++;
  }

  // 今日のセルへ移動
  sheet.getRange(i + 2,y).activateAsCurrentCell()
}

function summarizeMonthly() {
  const today = new Date();
  var api_key = PropertiesService.getScriptProperties().getProperty('APIKEY'); 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Monthly"); // Monthlyシートを指定
  sheet.insertRowBefore(2); //ヘッダーのすぐ下に行を追加する
  
  // 前月の年月を取得して表示
  const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const yearMonth = `${lastMonth.getFullYear()}年${lastMonth.getMonth() + 1}月`;
  sheet.getRange("A2").setValue(yearMonth);
  
  // 最初の指示
  var promptCell 
  = "以下に続く日記の内容から1ヶ月を整理して振り返ってください。\n" +
    + "最後に総括として私に対してやる気の出るコメントも添えてください\n"; 
  
  const result = makeDiaryPrompt(today);
  //日記が半分以上記入されていなかったら｢入力日記数不足｣と表示する。
  if(result.inputCellCount/result.totalCells<0.5) {
    sheet.getRange("C2").setValue("入力日記数不足");
    return;
  }

  var prompt = promptCell + result.prompt //前月の1ヶ月分の日記情報を取得する。
  var model = "gpt-4.1-nano"; // 使用するOpenAIモデルのID
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

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
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
  const lastSunday = getLastSunday();
  const lastMonday = new Date(lastSunday);
  
  lastMonday.setDate(lastSunday.getDate() - 6);
  var api_key = PropertiesService.getScriptProperties().getProperty('APIKEY'); 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly"); // Weeklyシートを指定
  sheet.insertRowBefore(2); //ヘッダーのすぐ下に行を追加する
  sheet.getRange("A2").setValue(`${lastMonday.toLocaleDateString('ja-JP')}`);
  sheet.getRange("B2").setValue(`${lastSunday.toLocaleDateString('ja-JP')}`);

  const sheetMemo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('memo');
  var promptCell = sheetMemo.getRange("B2").getValue();
  // 最初の指示
  // var promptCell 
  // = "##指示\n" +
  //   "あなたはGROWモデルを完全に習得したキャリアコーチです。またITコンサルタントおよび健康管理の専門家でもあります。\n" +
  //   "私の今週1週間の日記の内容からGROWモデルをもとにしてアドバイスをしてください。\n" +
  //   "回答は1200文字程度で構造的に記載して、マークダウン記法を使用しないでください。\n\n##日記\n" ; 
  
  const result = makeDiaryPrompt(lastMonday, 7);
  //日記が半分以上記入されていなかったら｢入力日記数不足｣と表示する。
  if(result.inputCellCount/result.totalCells<0.5) {
    sheet.getRange("C2").setValue("入力日記数不足");
    return;
  }

  var prompt = promptCell + result.prompt //7日分の日記データを取得する。
  var model = "gpt-4.1-mini"; // 使用するOpenAIモデルのID
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
  const userEmail = Session.getActiveUser().getEmail();
  var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
  var json = JSON.parse(response.getContentText());
  if (json.choices && json.choices.length > 0) {
    var generatedText = json.choices[0].message.content;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly");
    sheet.getRange("C2").setValue(generatedText);
    MailApp.sendEmail(userEmail, `${lastMonday.toLocaleDateString('ja-JP')}`+"~"+`${lastSunday.toLocaleDateString('ja-JP')}`+"までの1週間振り返り", generatedText);
  } else {
    sheet.getRange("C2").setValue("エラー: " + JSON.stringify(json));
  }

}

//最も近い日曜日を取得する。
function getLastSunday() {
  const today = new Date();
  
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

  // E列（列5）から指定日数分のデータを取得
  const contents = sheet
    .getRange(targetRow, 5, daysToRetrieve, 1)
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
    
    // メール送信のオプション
    const options = {
      htmlBody: body.replace(/\n/g, '<br>'), // 改行をHTMLの改行に変換
    };
    
    // 送信者が指定されている場合は追加
    if (from && emailPattern.test(from)) {
      options.replyTo = from;
    }
    
    // メール送信
    MailApp.sendEmail(to, subject, body, options);
    
    console.log(`メール送信完了: ${to}`);
    return { success: true, message: "メール送信完了" };
    
  } catch (error) {
    console.error("メール送信エラー:", error.message);
    return { success: false, message: error.message };
  }
}


