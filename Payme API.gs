//paymeデータをcsv形式で送信する関数　paymeAPIをもらい次第修正

function sendSheetAsCSV() {
  // スプレッドシートの設定
  var sheetId = '1cioe0mH6ZwgzJumMNRDzJ13yQVwpPCJjSU9g42T6M88';
  var sheetName = 'payme貼付用'; // 対象シートの名前を指定 ここを変更することで参照するシートを変更可能です
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  
  // A列で最後のデータがある行を取得
  var lastRow = sheet.getRange("A:A").getValues().filter(String).length;
  
  // シートデータをA2:Bまでの範囲で取得
  var dataRange = sheet.getRange(1, 1, lastRow, 3); // A2:Bの範囲を取得
  
  // シートデータをCSV形式に変換
  var csvData = convertSheetToCSV(dataRange);
  console.log(csvData)

  // HTTP POSTリクエストの設定
  var url = 'https://lz2wwtcwpebov7y7jro247cyay0mqvvh.lambda-url.ap-northeast-1.on.aws/'; 


  var boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW";
  
  // 固定情報
  var auth_email = PropertiesService.getScriptProperties().getProperty("AUTH_EMAIL");
  var auth_password = PropertiesService.getScriptProperties().getProperty("AUTH_PASSWORD");  // 正式なパスワードを後から共有
  var login_email = PropertiesService.getScriptProperties().getProperty("LOGIN_EMAIL");  // 使用しているアカウントのログインEメール
  var login_password = PropertiesService.getScriptProperties().getProperty("LOGIN_PASSWORD");   // 使用しているアカウントのログインパスワード
  
  // リクエストボディの作成
  var body = '--' + boundary + '\r\n'
             + 'Content-Disposition: form-data; name="file"; filename="attendance.csv"\r\n'
             + 'Content-Type: text/csv\r\n\r\n'
             + csvData + '\r\n'
             + '--' + boundary + '\r\n'
             + 'Content-Disposition: form-data; name="auth_email"\r\n\r\n'
             + auth_email + '\r\n'
             + '--' + boundary + '\r\n'
             + 'Content-Disposition: form-data; name="auth_password"\r\n\r\n'
             + auth_password + '\r\n'
             + '--' + boundary + '\r\n'
             + 'Content-Disposition: form-data; name="login_email"\r\n\r\n'
             + login_email + '\r\n'
             + '--' + boundary + '\r\n'
             + 'Content-Disposition: form-data; name="login_password"\r\n\r\n'
             + login_password + '\r\n'
             + '--' + boundary + '--';

  var options = {
    'method': 'post',
    'contentType': 'multipart/form-data; boundary=' + boundary,
    'payload': body
  };
  
  // POSTリクエストを送信
  var response = UrlFetchApp.fetch(url, options);
  
  // レスポンスをログに出力
  Logger.log(response.getContentText());
}

// CSVに変換する関数
function convertSheetToCSV(range) {
  var data = range.getValues();
  var csv = "";

  // 2次元配列をCSV形式に変換
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      // 日付データの場合、フォーマットを変更
      if (data[i][j] instanceof Date) {
        data[i][j] = Utilities.formatDate(data[i][j], Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
    }
    csv += data[i].join(",") + "\n";
  }
  
  return csv;
}


