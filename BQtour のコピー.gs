// function writeToSpreadsheet() {
//   var spreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1cioe0mH6ZwgzJumMNRDzJ13yQVwpPCJjSU9g42T6M88/edit#gid=1572048633');
//   var tourSheet = spreadsheet.getSheetByName('ツアー計算');
//   var currentDateValue = tourSheet.getRange('H2').getValue();
//   var previousDateValue = tourSheet.getRange('G2').getValue();
//   var e2Value = tourSheet.getRange('E2').getValue();

  
//     // 日付の形式をyyyy-mm-ddに変更
//     var currentDate = Utilities.formatDate(new Date(currentDateValue), Session.getScriptTimeZone(), 'yyyy-MM-dd');
//     var previousDate = Utilities.formatDate(new Date(previousDateValue), Session.getScriptTimeZone(), 'yyyy-MM-dd');

// // E2セルに値が入力されている場合のみポップアップを表示
//   if (e2Value) {
//     // 日付の形式をyyyy/mm/ddに変更してポップアップに表示
//     var currentDateDisplay = Utilities.formatDate(new Date(currentDateValue), Session.getScriptTimeZone(), 'yyyy/MM/dd');

//     // ユーザーに確認ポップアップを表示
//     var ui = SpreadsheetApp.getUi();
//     var response = ui.alert('確認', '' + currentDateDisplay + 'の支払日のものを出力します。よろしいですか？', ui.ButtonSet.YES_NO);

//     // ユーザーが「YES」を選択した場合のみ処理を続行
//     if (response == ui.Button.YES) {
//       executeQueryAndWriteToSpreadsheet(spreadsheet, currentDate, previousDate);
//     } else {
//       ui.alert('キャンセルされました。');
//     }
//   } else {
//     executeQueryAndWriteToSpreadsheet(spreadsheet, currentDate, previousDate);
//   }
// }

// function executeQueryAndWriteToSpreadsheet(spreadsheet, currentDate, previousDate) {
//   // BigQueryのクエリを定義
//   var sqlQuery = `SELECT CT.work_date , CT.worker_name , CT.room_name_common_area_name, CT.work_name, CT.cleaning_id, CT.status, CT.work_start_time, CT.work_end_time, RE.cleaner_id
//                   FROM \`m2m-core.su_wo.wo_cleaning_tour\` AS CT
//                   LEFT JOIN \`m2m-core.m2m_cleaning_prod.cleaning_cleaner_relations\` AS RE
//                   ON CT.cleaning_id = RE.cleaning_id
//                   WHERE work_date <= '${currentDate}' AND work_date >= '${previousDate}'`;

//   Logger.log(sqlQuery);

//   // BigQueryのクエリを実行
//   var request = {
//     query: sqlQuery,
//     useLegacySql: false
//   };
//   var queryResults = BigQuery.Jobs.query(request, 'm2m-core');

//   // クエリの実行を待機し、結果を取得
//   var jobId = queryResults.jobReference.jobId;
//   while (!queryResults.jobComplete) {
//     Utilities.sleep(1000);
//     queryResults = BigQuery.Jobs.getQueryResults('m2m-core', jobId);
//   }

//   // クエリ結果のデータを2次元配列に変換
//   var data = [];
//   data.push(queryResults.schema.fields.map(function(field) {
//     return field.name;
//   }));
//   queryResults.rows.forEach(function(row) {
//     var rowData = row.f.map(function(field) {
//       return field.v;
//     });
//     data.push(rowData);
//   });

//   // Googleスプレッドシートにデータを書き込む前にシートの値をクリアする
//   var sheet = spreadsheet.getSheetByName('BQ:wo_cleaning');
//   sheet.clear(); // シートの内容をクリア

//   // 取得したデータを元の範囲に直接書き込む
//   sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
// }
