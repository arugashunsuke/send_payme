function copyDataToDetailList() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var editSheet = spreadsheet.getSheetByName("編集");
  var detailSheet = spreadsheet.getSheetByName("明細");
  
  // 編集シートのデータを取得
  var editData = editSheet.getRange("A2:G").getValues();
  
  // 明細リストのデータの最終行を取得
  var lastRow = detailSheet.getLastRow();
  
  // 明細リストにデータを貼り付け
  detailSheet.getRange(lastRow + 1, 1, editData.length, editData[0].length).setValues(editData);
}

