/* 
　　機能1: 自己紹介シートを読み込み、誕生日の方がいたら祝う
　　機能2: 別途用意したスプレッドシートを読み込み、格言・名言を発言する
*/
function chatworkBot() { 
  // 使用するスプレッドシートを選択
  var id = "hoge";
  var spreadsheet = SpreadsheetApp.openById(id);

  // 0,1,2のいずれかの値を格納する
  var _random=Math.floor(Math.random()*2);

  // 疑似乱数が1なら、名言を送信する
  if(_random==1){
    /* ここからは、名言を送信する処理 */
    // どのシートを使うか指定
    var sheet = spreadsheet.getSheetByName('chatwork_bot');
  
    // シートの使用範囲のうち最終行を取得
    var maxRow = sheet.getDataRange().getLastRow(); 
    
    // 2～maxRowの間のランダムの整数を格納
    var numRow = Math.floor(2+Math.random()*maxRow);
 
    // チャットワークに送る文字列を生成
    var strBody = "[info]" + "'' " +
    sheet.getRange(numRow,1).getValue() + "''"  + "[hr]" +
    sheet.getRange(numRow,2).getValue() +
    sheet.getRange(numRow,3).getValue() + "[/info]";
 
    // チャットワークにメッセージを送信
    var cwClient = ChatWorkClient.factory({token: 'hoeg'}); //チャットワークAPI
    cwClient.sendMessage({
      room_id:hoge, //ルームID
      body: strBody
    });
  }
  
  /* ここからは、誕生日を祝う処理をする部分 */
  // どのシートを使うか指定
  var _sheet = spreadsheet.getSheetByName('ver_2022');
  
  // シートの使用範囲のうち最終行を取得
  var _maxRow = _sheet.getDataRange().getLastRow(); 
  
  // 今日の日付を取得
  var date = new Date();
  var today = Utilities.formatDate( date, 'Asia/Tokyo', 'MM/dd');

  // ループを回して、今日の日付と誕生日が一致するかを判定（3行目から開始する）
  for(let i = 3 ; i < _maxRow - 1; i++){
    var month = _sheet.getRange(i,55).getValue();
    var day = _sheet.getRange(i,56).getValue();

    // 適当な乱数を生成
    let rand=Math.floor(Math.random()*month*day);
    let message=['本年もご健康で、幸多き1年となりますように。','ますます充実した人生となりますように。','この1年が、たくさんの幸福と可能性に満ちた日々でありますように。','この1年、健康でたくさんの幸せが訪れますように。','この1年が、あなたにとって素晴らしい年になりますように。','これからの人生が、幸せで溢れますように。','素敵な1年になることを心から願ってます。','笑顔溢れる素敵な1年になりますように。'];

    // 日付の表示形式を○○/○○に固定
    if(month < 10){
      month = "0" + month;
    }
    if(day < 10){
     day = "0" + day;
    }
    // 月と日を連結
    var birthday = month + "/" + day;

    // 一致したら（今日が誕生日）、チャットワークにメッセージを送信
    if(birthday == today){
      // チャットワークに送る文字列を生成
      var strBody = "[info][title]Wishing you all the best today and throughout the coming year![/title]" +
      "本日（" + birthday + "）は、"+ _sheet.getRange(i,2).getValue() + 
      "さんのお誕生日です。" + "おめでとうございます!!(^)👏\n"　+message[(rand)%message.length]+ "[/info]";
 
      // チャットワークにメッセージを送信
      var cwClient = ChatWorkClient.factory({token: "hoge"}); //チャットワークAPI
      cwClient.sendMessage({
        room_id:hoge, //ルームID
        body: strBody
      });
    }    
  };  
}