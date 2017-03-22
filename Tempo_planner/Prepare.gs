function preparePlanningData() {

  var allValues = Utils.getValuesBySheetName('Planning Extract'),
    ss = SpreadsheetApp.getActiveSpreadsheet(),
    startDateScope = ss.getRange('Rates by Team!H2').getValue(),
    endDateScope = ss.getRange('Rates by Team!I2').getValue(),
    aProjects = Utils.getProjectsArray(),
    aAccounts = Utils.getAccountsArray()

  //Prepare Data and Clean Planning Extract
  var resExtract = ss.getSheetByName('Planning Data')
  if(resExtract){
    ss.deleteSheet(resExtract)
  }
  resExtract = ss.insertSheet('Planning Data',1)
  resExtract.setTabColor('#606060')

  for (var i = 1; i < allValues.length; i++) {
    var startDate = Utils.removeWeekEndFromPeriod(allValues[i][11], 1),
      endDate = Utils.removeWeekEndFromPeriod(allValues[i][12], -1)

    if ( Utils.isDateInScope(startDate, startDateScope, endDateScope) || Utils.isDateInScope(endDate, startDateScope, endDateScope)) {
      if (Utils.isProjectExists(allValues[i][4], aProjects)) {
        var rate = Utils.getRate(allValues[i][2], allValues[i][4], allValues[i][8])
        resExtract.appendRow([allValues[i][2], allValues[i][0], allValues[i][4], Utils.getProjectByCode(allValues[i][4]), allValues[i][8], allValues[i][9], allValues[i][10], startDate, endDate, Utils.getProfile(allValues[i][2]), Utils.getOrg(allValues[i][2]), rate])
      }
    }
  }
  resExtract.hideSheet()
}

//Prepare Data2
function prepareForecastData() {

  var ss = SpreadsheetApp.getActiveSpreadsheet(),
    startDateScope = ss.getRange('Rates by Team!H2').getValue(),
    endDateScope = ss.getRange('Rates by Team!I2').getValue(),
    numDaysByMonth = ss.getRange('Rates by Team!G2').getValue(),
    diffToRealWorkHours = ss.getRange('Rate Card!E7').getValue(),
    errors = false

  //Remove old Sheets
  var resExtract = ss.getSheetByName('Forecast Data')
  if(resExtract){
    resExtract.clear()
  } else {
    ss.insertSheet('Forecast Data', 1)
  }

  var errorData = ss.getSheetByName('Errors')
  if(errorData){
    ss.deleteSheet(errorData)
  }

  var resExtract = ss.getSheetByName('Forecast Data')
  resExtract.appendRow(['Consultant', 'Profile', 'Team', 'Project', 'Organisation', 'month', 'accounting', 'days', 'Rate', 'Price / day', 'Cost', 'FTE', 'Rate atc.', 'days act.', 'Price / day act.', 'Cost act.', 'FTE act.']).setTabColor('#10f020')

  var newValues = Utils.getValuesBySheetName('Planning Data')

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
          totalPrice =  price * workDays * ratio * diffToRealWorkHours
        }

        // Write to file
        resExtract.appendRow([newValues[i][1], newValues[i][9], team, newValues[i][3], newValues[i][10], Format.formatDateForReports(nextBegin), accounting, workDays * ratio, rate, price, totalPrice, (workDays* ratio)/numDaysByMonth])

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
          errorData.appendRow([newValues[i][1], newValues[i][9], team, newValues[i][3], newValues[i][10], Format.formatDateForReports(nextBegin), accounting, workDays * ratio, rate, price, totalPrice, (workDays* ratio)/numDaysByMonth])
        }
      }
      nextBegin = Utils.getFirstDayOfNextMonth(nextBegin)

    }

  }
  // Format prices
  resExtract.getRange(2, 10, resExtract.getLastRow(), 2).setNumberFormat('#,##0.00;(#,##0.00)')

  //resExtract.hideSheet()
}

function prepareActualsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
    cache = CacheService.getScriptCache()

  //Sort Actuals Extract by username
  ss.getSheetByName('Actuals Extract').sort(5)

  //Cache data
  var allValues = Utils.getValuesBySheetName('Actuals Extract'),
    aAccounts = Utils.getAccountsArray(),
    oProfiles = {},
    tempArray = [],
    usersArray = [],
    currentUser = allValues[0][4]

  usersArray.push(currentUser)

  for (var i = 0; i < allValues.length; i++) {

    if(Utils.isAccountExists(allValues[i][7], aAccounts)) {

      if (currentUser != allValues[i][4]) {

        // sort tempArray by account
        tempArray.sort(function(a, b) {
          var textA = a[2]
          var textB = b[2]
          return (textA < textB) ? -1 : (textA > textB) ? 1 : 0
        })

        // clear existing cache
        if (cache.get(currentUser)) {
          cache.remove(currentUser)
        }

        // cache for 1 hour
        cache.put(currentUser, JSON.stringify(tempArray), 3600)

        // re-init value for next username
        currentUser = allValues[i][4]
        usersArray.push(currentUser)
        tempArray = []

      }
      var date = allValues[i][3].toString()
      tempArray.push([allValues[i][2], date, allValues[i][7], allValues[i][5]])

    }

  }

  //cache last user
  // sort tempArray by account
  tempArray.sort(function(a, b) {
    var textA = a[2]
    var textB = b[2]
    return (textA < textB) ? -1 : (textA > textB) ? 1 : 0
  })

  // clear existing cache
  if (cache.get(currentUser)) {
    cache.remove(currentUser)
  }

  // cache for 1 hour
  cache.put(currentUser, JSON.stringify(tempArray), 3600)

  if (cache.get('users')) {
    cache.remove('users')
  }
  // cache for 1 hour
  cache.put('users',JSON.stringify(usersArray), 3600)

  //SpreadsheetApp.getUi().alert('Success : Data extracted')

  writeToActualsData()

}

function writeToActualsData() {

  var ss = SpreadsheetApp.getActiveSpreadsheet(),
    cache = CacheService.getScriptCache(),
    aUsers = JSON.parse(cache.get('users')),
    aLines = []

  /*var resExtract = ss.getSheetByName('Actuals Data')
   if(resExtract){
   ss.deleteSheet(resExtract)
   }
   resExtract = ss.insertSheet('Actuals Data',3)
   resExtract.setTabColor('#606060')
   */

  // prepare temporary lines array
  if(aUsers) {

    for (var i = 0; i < aUsers.length; i++) {

      var username = aUsers[i],
        profile = cache.get(username)

      if (profile) {

        var aProfile = JSON.parse(profile),
          hours = 0,
          startDate = new Date(aProfile[0][1]),
          endDate = new Date(aProfile[0][1]),
          currentDate = new Date(aProfile[0][1]),
          currentAccount = aProfile[0][2],
          currentUser =  aProfile[0][3],
          j = 0

        while(j < aProfile.length) {


          //Logger.log(Object.prototype.toString.call(currentDate))

          if (currentAccount != aProfile[j][2]) {

            aLines.push([username, currentUser, Format.convertDate(startDate), Format.convertDate(endDate), hours, currentAccount, Utils.getBillableByKey(currentAccount)])

            // reinit data
            currentAccount = aProfile[j][2]
            hours = 0
            startDate = new Date(aProfile[j][1])
            endDate = new Date(aProfile[j][1])
            currentDate = new Date(aProfile[j][1])
            currentUser = aProfile[j][3]

          }

          if(currentDate != new Date(aProfile[j][1])) {
            if(new Date(aProfile[j][1]) < startDate) {
              startDate = new Date(aProfile[j][1])
            }
            if(new Date(aProfile[j][1]) > endDate) {
              endDate = new Date(aProfile[j][1])
            }
          }
          currentDate = new Date(aProfile[j][1])
          hours += aProfile[j][0]

          j += 1

        }

        aLines.push([username, currentUser, Format.convertDate(startDate), Format.convertDate(endDate), hours, currentAccount, Utils.getBillableByKey(currentAccount)])

      }
    }

    // Write to file

    // copy Forecast

    var resExtract = ss.getSheetByName('Forecast + Actuals Data')
    if(resExtract){
      ss.deleteSheet(resExtract)
    }

    var sheet = ss.getSheetByName('Forecast Data')
    ss.setActiveSheet(sheet.copyTo(ss))
    ss.moveActiveSheet(3)
    ss.renameActiveSheet('Forecast + Actuals Data')

    var insertSheet = ss.getSheetByName('Forecast + Actuals Data')

    var hoursDay = ss.getRange('Rate Card!F2').getValue(),
      numDaysByMonth = ss.getRange('Rates by Team!G2').getValue(),
      arrDate = aLines[0][2].split('/'),
      month = arrDate[1] + '/' + arrDate[2],
      k = 0,
      totalPrice = '',
      price = ''

    //Logger.log(aLines[0][2].split('/'))
    Logger.log(aLines)
    //Logger.log(Object.prototype.toString.call(aLines[0][2]))

    while(k < aLines.length) {

      var rate = Utils.getRateByAccountActuals(aLines[k][0], aLines[k][5]),
        accounting = aLines[k][6],
        hoursAct = aLines[k][4],
        daysAct = hoursAct/hoursDay



      if(rate === '' || rate === 'Not defined') {
        price = ''
        totalPrice = ''
      } else {
        price = Utils.getPriceByRate(rate, accounting)
        totalPrice = price*daysAct
      }


      insertSheet.appendRow([aLines[k][1], Utils.getProfile(aLines[k][0]), '', '', Utils.getOrg(aLines[k][0]), month, aLines[k][6], '', '', '', '', '', rate, daysAct, price , totalPrice, daysAct/numDaysByMonth ])

      k += 1
    }
    // Format prices
    insertSheet.getRange(2, 15, insertSheet.getLastRow(), 2).setNumberFormat('#,##0.00;(#,##0.00)')

  } else {
    SpreadsheetApp.getUi().alert('Cache expired : Launch Actuals Data Extract again')
  }

  //Logger.log(aLines)
}
