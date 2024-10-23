function writeToSpreadsheet() {
  var spreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1cioe0mH6ZwgzJumMNRDzJ13yQVwpPCJjSU9g42T6M88/edit?gid=1970839194#gid=1970839194');
  var tourSheet = spreadsheet.getSheetByName('ツアー計算');
  var currentDateValue = tourSheet.getRange('H2').getValue();
  var previousDateValue = tourSheet.getRange('G2').getValue();
  var e2Value = tourSheet.getRange('E2').getValue();

  // 日付の形式をyyyy-mm-ddに変更
  var currentDate = Utilities.formatDate(new Date(currentDateValue), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var previousDate = Utilities.formatDate(new Date(previousDateValue), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // E2セルに値が入力されている場合のみポップアップを表示
  if (e2Value) {
    // 日付の形式をyyyy/mm/ddに変更してポップアップに表示
    var currentDateDisplay = Utilities.formatDate(new Date(currentDateValue), Session.getScriptTimeZone(), 'yyyy/MM/dd');

    // ユーザーに確認ポップアップを表示
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('確認', '' + currentDateDisplay + 'の支払日のものを出力します。よろしいですか？', ui.ButtonSet.YES_NO);

    // ユーザーが「YES」を選択した場合のみ処理を続行
    if (response == ui.Button.YES) {
      executeQueryAndWriteToSpreadsheet(spreadsheet, currentDate, previousDate);
    } else {
      ui.alert('キャンセルされました。');
    }
  } else {
    executeQueryAndWriteToSpreadsheet(spreadsheet, currentDate, previousDate);
  }
}

function executeQueryAndWriteToSpreadsheet(spreadsheet, currentDate, previousDate) {
  // BigQueryのクエリを定義
  var sqlQuery = `SELECT CT.work_date , CT.worker_name , CT.room_name_common_area_name, CT.work_name, CT.cleaning_id, CT.status, CT.work_start_time, CT.work_end_time, RE.cleaner_id,
                  FROM \`m2m-core.su_wo.wo_cleaning_tour\` AS CT
                  LEFT JOIN \`m2m-core.m2m_cleaning_prod.cleaning_cleaner_relations\` AS RE
                  ON CT.cleaning_id = RE.cleaning_id
                  WHERE work_date <= '${currentDate}' AND work_date >= '${previousDate}'`;

  Logger.log(sqlQuery);

  // BigQueryのクエリを実行
  var request = {
    query: sqlQuery,
    useLegacySql: false
  };
  var queryResults = BigQuery.Jobs.query(request, 'm2m-core');

  // クエリの実行を待機し、結果を取得
  var jobId = queryResults.jobReference.jobId;
  while (!queryResults.jobComplete) {
    Utilities.sleep(1000);
    queryResults = BigQuery.Jobs.getQueryResults('m2m-core', jobId);
  }

  // クエリ結果のデータを2次元配列に変換
  var data = [];
  data.push(queryResults.schema.fields.map(function(field) {
    return field.name;
  }));
  queryResults.rows.forEach(function(row) {
    var rowData = row.f.map(function(field) {
      return field.v;
    });
    data.push(rowData);
  });

  // Googleスプレッドシートにデータを書き込む前にシートの値をクリアする
  var sheet = spreadsheet.getSheetByName('BQ元');
  sheet.getRange('A2:I').clear(); // A2:Iの範囲をクリア

  // 取得したデータを元の範囲に直接書き込む
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // データの書き込みが完了した後に追加処理を呼び出す
  executeCostAndSpecialBonus();
}

function executeCostAndSpecialBonus() {
  // スプレッドシートとシートの取得
  let ss = SpreadsheetApp.openById("1cioe0mH6ZwgzJumMNRDzJ13yQVwpPCJjSU9g42T6M88");
  let sheet_1 = ss.getSheetByName("BQ:wo_cleaning"); // yest/day のコピーシート

  // getApiToken関数からトークンを取得
  const token = getApiToken();
  
  // トークンが正しく取得できていない場合、処理を終了
  if (!token) {
    SpreadsheetApp.getUi().alert("トークンの取得に失敗しました");
    return;
  }

  // E3:EのcleaningIdとL3:LおよびI3:Iの列を取得
  let cleaningIds = sheet_1.getRange("E2:E").getValues().flat(); // E列のcleaningIdを取得
  let lColumnValues = sheet_1.getRange("L2:L").getValues().flat(); // L列の値を取得
  let iColumnValues = sheet_1.getRange("I2:I").getValues().flat(); // I列のcleanerIdを取得
  sheet_1.getRange('J2:K').clear(); // J2:Kの範囲をクリア

  // L列が空白でない行のみを対象とする
  cleaningIds.forEach((cleaningId, index) => {
    if (cleaningId && lColumnValues[index]) {  // L列が空白でない場合に処理を実行
      // 1. cost API リクエスト
      let costApiUrl = `https://api-cleaning.m2msystems.cloud/v4/cleanings/approval_requests/latest`;
      const costPayload = {
        cleaningIds: [cleaningId]
      };

      const costOptions = {
        'method': 'POST',
        'contentType': 'application/json',
        'headers': {
          'Authorization': 'Bearer ' + token
        },
        'payload': JSON.stringify(costPayload),
        'muteHttpExceptions': true
      };

      // Cost APIリクエストの送信
      let costResponse;
      try {
        costResponse = UrlFetchApp.fetch(costApiUrl, costOptions);
      } catch (error) {
        Logger.log(`Cost APIリクエストに失敗しました: ${error}`);
        return;
      }

      const costResponseText = costResponse.getContentText();
      const costData = JSON.parse(costResponseText);
      
      // APIからのレスポンスを解析して、statusが'Approved'のものを探す
      const cleaningInfo = costData[cleaningId];
      if (cleaningInfo) {
        for (const cleanerId in cleaningInfo) {
          const data = cleaningInfo[cleanerId];
          if (data && data.status === 'Approved' && data.cleanerId === iColumnValues[index]) {
            // スプレッドシートにunitPriceとspecialBonusを書き込む
            sheet_1.getRange(index + 2, 10).setValue(data.unitPrice);  // J列にUnit Price
            const specialBonus = data.specialBonus !== null ? data.specialBonus : 0;
            sheet_1.getRange(index + 2, 11).setValue(specialBonus);  // K列にSpecial Bonusを出力
          }
        }
      }
    }
  });
}

function getApiToken() {
  const url = 'https://api.m2msystems.cloud/login';
  
  // スクリプトプロパティからメールアドレスとパスワードを取得
  const mail = PropertiesService.getScriptProperties().getProperty("mail_address");
  const pass = PropertiesService.getScriptProperties().getProperty("pass");

  // APIリクエストのペイロード
  const payload = {
    email: mail,
    password: pass
  };

  // HTTPリクエストのオプション
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // APIリクエストを送信し、レスポンスを取得
  const response = UrlFetchApp.fetch(url, options);
  
  // ステータスコードが200の場合はトークンを返す
  if (response.getResponseCode() == 200) {
    const json = JSON.parse(response.getContentText());
    const token = json.accessToken;
    return token;
  } else {
    Logger.log("トークン取得エラー: ステータスコード " + response.getResponseCode());
    return null;
  }
}