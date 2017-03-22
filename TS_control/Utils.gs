function Utils() {

}
Utils.getValuesBySheetName = function(sheetname) {
  var dataSheet = SpreadsheetApp.getActive().getSheetByName(sheetname)
  return dataSheet.getDataRange().getValues()
}