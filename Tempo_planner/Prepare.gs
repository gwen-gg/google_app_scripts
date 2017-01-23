function prepareData() {
  this.getListOfLinesInScope()
}

function getListOfLinesInScope() {

  var allValues = Utils.getValuesBySheetName('Tempo Extract'),
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    startDateScope = ss.getRange('Rates by Team!H2').getValue(),
    endDateScope = ss.getRange('Rates by Team!I2').getValue()

  //Get Months to Forecast
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
    aProjects = Utils.getProjectsArray(),
    aCacheData = []

  //Prepare Data and Clean Tempo Extract
  var ss = SpreadsheetApp.getActive()
  var resExtract = ss.getSheetByName('Data')
  if(resExtract){
    ss.deleteSheet(resExtract)
  }
  resExtract = ss.insertSheet('Data',1)
  resExtract.setTabColor('#606060')

  for (var i = 1; i < allValues.length; i++) {
    var startDate = Utils.removeWeekEndFromPeriod(allValues[i][11], 1),
      endDate = Utils.removeWeekEndFromPeriod(allValues[i][12], -1)

    if ( Utils.isDateInScope(startDate, startDateScope, endDateScope) || Utils.isDateInScope(endDate, startDateScope, endDateScope)) {
      if (Utils.isProjectExists(allValues[i][4], aProjects)) {
        var rate = Utils.getRate(allValues[i][2], allValues[i][4], allValues[i][8])
        aCacheData.push([allValues[i][2], allValues[i][0], allValues[i][4], Utils.getProjectByCode(allValues[i][4]), allValues[i][8], allValues[i][9], allValues[i][10], startDate, endDate, Utils.getProfile(allValues[i][2]), Utils.getOrg(allValues[i][2]), rate])
      }
    }
  }

  for (var i = 0; i < aCacheData.length; i++) {
    resExtract.appendRow([aCacheData[i][0], aCacheData[i][1], aCacheData[i][2], aCacheData[i][3], aCacheData[i][4], aCacheData[i][5], aCacheData[i][6], aCacheData[i][7], aCacheData[i][8], aCacheData[i][9], aCacheData[i][10], aCacheData[i][11]])
  }

  //resExtract.hideSheet()
  //this.SplitByMonthlyPeriod()

}


function cacheData() {

  var allValues = Utils.getValuesBySheetName('Tempo Extract'),
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    startDateScope = ss.getRange('Rates by Team!H2').getValue(),
    endDateScope = ss.getRange('Rates by Team!I2').getValue(),
    aProjects = Utils.getProjectsArray(),
    aCacheData = []


  for (var i = 1; i < allValues.length; i++) {
    var startDate = Utils.removeWeekEndFromPeriod(allValues[i][11], 1),
      endDate = Utils.removeWeekEndFromPeriod(allValues[i][12], -1)

    if ( Utils.isDateInScope(startDate, startDateScope, endDateScope) || Utils.isDateInScope(endDate, startDateScope, endDateScope)) {
      if (Utils.isProjectExists(allValues[i][4], aProjects)) {
        var rate = Utils.getRate(allValues[i][2], allValues[i][4], allValues[i][8])
        aCacheData.push([allValues[i][2], allValues[i][0], allValues[i][4], Utils.getProjectByCode(allValues[i][4]), allValues[i][8], allValues[i][9], allValues[i][10], startDate, endDate, Utils.getProfile(allValues[i][2]), Utils.getOrg(allValues[i][2]), rate])
      }
    }
  }
}

function SplitByMonthlyPeriod() {
//Prepare Data2

  var ss = SpreadsheetApp.getActiveSpreadsheet(),
    startDateScope = ss.getRange('Rates by Team!H2').getValue(),
    endDateScope = ss.getRange('Rates by Team!I2').getValue(),
    numDaysByMonth = ss.getRange('Rates by Team!G2').getValue(),
    errors = false

  //REmove old Sheets
  var resExtract = ss.getSheetByName('Data 2')
  if(resExtract){
    resExtract.clear()
  } else {
    ss.insertSheet('Data 2', 1)
  }



  var errorData = ss.getSheetByName('Errors')
  if(errorData){
    ss.deleteSheet(errorData)
  }

  var resExtract = ss.getSheetByName('Data 2')
  resExtract.appendRow(['Consultant', 'Profile', 'Team', 'Project', 'Organisation', 'month', 'accounting', 'days', 'Rate', 'Price / day', 'Cost', 'FTE']).setTabColor('#10f020')

  var newValues = Utils.getValuesBySheetName('Data')

  for (var i = 0; i < newValues.length; i++) {

    var startDate = newValues[i][7],
      endDate = newValues[i][8],
      numDays = Utils.getWorkDays(startDate, endDate),
      compareDays = 0,
      nextBegin = startDate


    while(nextBegin < endDate) {

      var endMonthPeriod = Utils.getLastDayOfMonth(nextBegin)

      if(Utils.isDateInScope(nextBegin, startDateScope, endDateScope)) {
        var ratio = newValues[i][6],
          workDays = 0,
          username = newValues[i][0],
          team = newValues[i][2],
          accounting = newValues[i][4],
          totalPrice = 0,
          price = 0,
          rate = newValues[i][11]

        if (endDate < endMonthPeriod ) {
          workDays = Utils.getWorkDays(nextBegin, endDate)
        } else {
          workDays = Utils.getWorkDays(nextBegin, endMonthPeriod)
        }

        if(rate === '' || rate === 'Not defined') {
          price = ''
          totalPrice = ''
        } else {
          price = Utils.getPriceByRate(rate, accounting)
          totalPrice =  price * workDays * ratio
        }

        // Write to file
        resExtract.appendRow([newValues[i][1], newValues[i][9], team, newValues[i][3], newValues[i][10], Format.formatDateForReports(nextBegin), accounting, workDays * ratio, rate, price, totalPrice, (workDays* ratio)/numDaysByMonth])

        //Format data
        resExtract.getRange(resExtract.getLastRow(), 10).setNumberFormat('#,##0.00;(#,##0.00)')
        resExtract.getRange(resExtract.getLastRow(), 11).setNumberFormat('#,##0.00;(#,##0.00)')

        //If rate is not the correct one for the consultant, format row
        if(Utils.isDefaultRateForConsultant(username, rate) === false) {
          var errorRange = resExtract.getRange(resExtract.getLastRow(), 1, 1, 12)
          Format.setErrorRow(errorRange, ['#f6a067','#000000'])
        }

        if (rate === 'Not defined' || newValues[i][9] === 'Not defined' || newValues[i][10] === 'Not defined') {
          if(errors === false) {
            errorData = ss.insertSheet('Errors',1)
            errorData.setTabColor('#ff5300')
          }
          errors = true
          errorData.appendRow([newValues[i][1], newValues[i][9], team, newValues[i][3], Format.formatDateForReports(nextBegin), nextBegin, accounting, workDays * ratio, rate, price, totalPrice, (workDays* ratio)/numDaysByMonth])
        }
      }
      nextBegin = Utils.getFirstDayOfNextMonth(nextBegin)

    }

  }

  //resExtract.hideSheet()
}