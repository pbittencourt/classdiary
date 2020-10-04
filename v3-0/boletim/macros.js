function reduceTurmas() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D:D').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('17:17').activate();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  spreadsheet.getRange('B2:C16').activate()
  .sort({column: 2, ascending: true});
};