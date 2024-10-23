function copyAndPaste() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // コピー元のシートと範囲を指定
  var sourceSheet = spreadsheet.getSheetByName('統合');
  var sourceRange = sourceSheet.getRange('A:G');

  // コピー先のシートを指定し、Aセル以下の値だけをクリア
  var destinationSheet = spreadsheet.getSheetByName('編集');
  destinationSheet.getRange("A:G").clearContent(); // clearContent()メソッドでA:Gの範囲をクリア

  // データをコピーして貼り付け
  var data = sourceRange.getValues();
  
  // ツアー計算シートのH2とG2セルから日付を取得
  var tourSheet = spreadsheet.getSheetByName('ツアー計算');
  var today = new Date(tourSheet.getRange('H2').getValue());
  var yesterday = new Date(tourSheet.getRange('G2').getValue());

  // A列が前日の日付の場合、当日の日付に変換し、D列のツアー名の前に"前日 "を追加
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] instanceof Date && data[i][0].getTime() === yesterday.getTime()) {
      data[i][0] = today; // 当日の日付に変換
      data[i][3] = "前日 " + data[i][3]; // ツアー名の前に"前日 "を追加
    }
  }

  // 常に1行目にデータを貼り付け
  destinationSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}
