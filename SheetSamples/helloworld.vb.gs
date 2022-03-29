function HelloWorld() {

	var spreadsheet = SpreadsheetApp.getActive();
//
// HelloWorld Macro
//

//
spreadsheet.getCurrentCell().setValue('Hello')
spreadsheet.getRange('B1').activate();
spreadsheet.getCurrentCell().setValue('World')
    spreadsheet.getRange('C1').activate();
spreadsheet.getCurrentCell().setFormula('=CONCATENATE(R[-4]C[0],R[5]C[0],R[0]C[-3],R[0]C[4])')
    spreadsheet.getRange('D1').activate();
spreadsheet.getCurrentCell().setFormula('=CONCATENATE(R[0]C[-3],R[0]C[-2],R[0]C[-1])')
    spreadsheet.getRange('A2').activate();
}

