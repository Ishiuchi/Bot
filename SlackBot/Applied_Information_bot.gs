/* 
　　スプレッドシート（事前にスクレイピングで応用情報技術者試験の重要語句を収集）を読み込み、指定したチャネルに送信する
*/
function postSlack()
{
  //自分のどのチャンネルに書き込みをするか選択
  var url = "https://hooks.slack.com/services/hoge/hoge/hoge";

  // スプレッドシート取得
  var id = "hoge";
  var spreadsheet = SpreadsheetApp.openById(id);
  // どのシートを使うか指定
  var sheet = spreadsheet.getSheetByName('ap');        
  
  //2行目～最終行の間で、ランダムな行番号を算出する
  var lastRow = sheet.getLastRow();
  var row = Math.ceil(Math.random() * (lastRow-1)) + 1;

  var terminology = sheet.getRange(row,1).getValue();
  var details = sheet.getRange(row,2).getValue();

  var options = {
    "method"  : "POST",
    "headers" : {"Content-type":"application/json"},
    "payload" : '{"text":"'+'『 ' + terminology +' 』\n\n　' + details + '"}'
  };
  UrlFetchApp.fetch(url, options);
}