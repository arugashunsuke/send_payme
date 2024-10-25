function sendSheetAsCSV() {
  const ui = SpreadsheetApp.getUi();
  const email = Session.getActiveUser().getEmail(); // 実行者のメールアドレスを取得
  const ulid = generateULID(); // ULIDを生成
  const executionTime = new Date(); // 実行時間
  let status = ''; // 実行ステータスを後で設定
  let apiResponseLog = ''; // APIレスポンスのログ

  // 確認1
  const response1 = ui.alert('確認', 'CSVをアップロードしてもよろしいですか？', ui.ButtonSet.YES_NO);
  if (response1 != ui.Button.YES) {
    ui.alert('キャンセルされました。');
    status = 'Cancelled'; // キャンセルとしてステータスを設定
    logExecutionToBigQuery(ulid, email, executionTime, status, apiResponseLog);
    return;
  }

  // 確認2
  const response2 = ui.alert('確認', '本当にCSVをアップロードしますか？', ui.ButtonSet.YES_NO);
  if (response2 != ui.Button.YES) {
    ui.alert('キャンセルされました。');
    status = 'Cancelled'; // キャンセルとしてステータスを設定
    logExecutionToBigQuery(ulid, email, executionTime, status, apiResponseLog);
    return;
  }

  // ユーザープロパティからログイン情報を取得
  const userProperties = PropertiesService.getUserProperties();
  let login_email = userProperties.getProperty("LOGIN_EMAIL");
  let login_password = userProperties.getProperty("LOGIN_PASSWORD");

  // ユーザープロパティが設定されていない場合、プロンプトで入力を促す
  if (!login_email || !login_password) {
    login_email = ui.prompt('ログイン情報が設定されていません', 'ログインに使用するメールアドレスを入力してください:', ui.ButtonSet.OK).getResponseText();
    login_password = ui.prompt('ログイン情報が設定されていません', 'ログインに使用するパスワードを入力してください:', ui.ButtonSet.OK).getResponseText();
    
    // ユーザープロパティに保存
    userProperties.setProperty("LOGIN_EMAIL", login_email);
    userProperties.setProperty("LOGIN_PASSWORD", login_password);
  }

  // スプレッドシートの設定
  const sheetId = '1cioe0mH6ZwgzJumMNRDzJ13yQVwpPCJjSU9g42T6M88';
  const sheetName = 'payme貼付用'; 
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  
  // A列で最後のデータがある行を取得
  const lastRow = sheet.getRange("A:A").getValues().filter(String).length;
  
  // シートデータをA2:Bまでの範囲で取得
  const dataRange = sheet.getRange(1, 1, lastRow, 3); 
  
  // シートデータをCSV形式に変換
  const csvData = convertSheetToCSV(dataRange);
  console.log(csvData);

  // HTTP POSTリクエストの設定
  const url = 'https://lz2wwtcwpebov7y7jro247cyay0mqvvh.lambda-url.ap-northeast-1.on.aws/'; 
  const boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW";
  
  // 固定情報
  const auth_email = PropertiesService.getScriptProperties().getProperty("AUTH_EMAIL");
  const auth_password = PropertiesService.getScriptProperties().getProperty("AUTH_PASSWORD");

  // リクエストボディの作成
  const body = '--' + boundary + '\r\n'
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

  const options = {
    'method': 'post',
    'contentType': 'multipart/form-data; boundary=' + boundary,
    'payload': body
  };
  
  // POSTリクエストを送信
  try {
    const response = UrlFetchApp.fetch(url, options);
    apiResponseLog = response.getContentText(); // APIレスポンスを取得
    Logger.log(apiResponseLog);
    status = 'Complete'; // 完了ステータスを設定
  } catch (error) {
    apiResponseLog = `Error: ${error.toString()}`;
    Logger.log(apiResponseLog);
    status = 'Failed'; // 失敗した場合
  }

  // BigQueryにログを保存
  logExecutionToBigQuery(ulid, email, executionTime, status, apiResponseLog);
}