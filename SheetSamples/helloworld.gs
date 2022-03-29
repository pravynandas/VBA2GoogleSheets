
function HelloWorld() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setValue('Hello');
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('World');
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setFormula('=CONCAT(A1,B1)');
  spreadsheet.getRange('D1').activate();
  spreadsheet.getCurrentCell().setFormula('=CONCATENATE(A1,B1,C1)');
  spreadsheet.getRange('D2').activate();
};