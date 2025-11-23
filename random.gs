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