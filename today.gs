function fcToday() {
  /*Dateオブジェクトで取得した日付をdに代入*/
  var d = new Date(); 

  // 2026年は5列目
  var y = 6 

 // dから月と日だけを抽出する。
  var mon = d.getMonth() + 1;
  var d2 = d.getDate();

  // mm/dd の文字列を生成する。
  var ogitoday = mon +"/" + d2;


  
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
