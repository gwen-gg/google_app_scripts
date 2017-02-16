function Format() {

}


Format.convertDate = function(inputFormat) {
  function pad(s) { return (s < 10) ? '0' + s : s; }
  var d = new Date(inputFormat);
  return [pad(d.getDate()), pad(d.getMonth()+1), d.getFullYear()].join('/');
}

Format.getMonthName = function(value){
  var month = {};
  month[0] = "jan";
  month[1] = "feb";
  month[2] = "mar";
  month[3] = "apr";
  month[4] = "may";
  month[5] = "jun";
  month[6] = "jul";
  month[7] = "aug";
  month[8] = "sep";
  month[9] = "oct";
  month[10] = "nov";
  month[11] = "dec";
  return month[value];
}

Format.formatDateForReports = function(date) {
  var mm = date.getMonth()+1
  var yyyy = date.getFullYear()

  if ( mm < 10 ) {
    mm = '0'+ mm
  }
  return mm+'/'+yyyy
}

Format.fontify = function(range, colorArray, hAlign, sizeFont) {
  range.setBackground(colorArray[0]);
  range.setFontColor(colorArray[1]);
  range.setHorizontalAlignment(hAlign);
  range.setFontSize(sizeFont);
}

Format.setErrorRow = function(range, colorArray) {
  Format.fontify(range, colorArray, 'left', 10);
}

/*Format.setTitleSheet = function(sheet, cell, cellRange, title) {
 var getCell = sheet.getRange(cell);
 sheet.getRange(cellRange).merge();
 Format.fontify(sheet.getRange(cell), ['#00a8c0','#FF0000'], 'left', 10);
 getCell.setValue(title);
 }*/