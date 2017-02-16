function Utils() {
}

Utils.isEmpty = function(cell) {
  if (typeof cell === 'string') {
    return true;
  }
  return false;
}

Utils.getValuesBySheetName = function(sheetname) {
  var dataSheet = SpreadsheetApp.getActive().getSheetByName(sheetname)
  return dataSheet.getDataRange().getValues()
}


Utils.getWorkDays = function(firstDate, lastDate) {

  if (firstDate > lastDate) return -1
  var start = new Date(firstDate.getTime())
  var end = new Date(lastDate.getTime())
  var count = 0
  while (start <= end) {
    if (start.getDay() != 0 && start.getDay() != 6)
      count++
    start.setDate(start.getDate() + 1)
  }
  return count
}

Utils.isWorkDay = function(date) {
  if (date.getDay() != 0 && date.getDay() != 6) {
    return true
  }
  return false
}

// remove weekends from extremities
Utils.removeWeekEndFromPeriod = function(date, direction) {
  var returnDate = new Date(date.getTime())
  //crop start (left to right)
  if (direction > 0) {
    if (returnDate.getDay() === 0) {
      returnDate.setDate(returnDate.getDate() + 1)
    }
    if (returnDate.getDay() === 6) {
      returnDate.setDate(returnDate.getDate() + 2)
    }
  } else { //crop end (right to left)
    if (returnDate.getDay() === 0) {
      returnDate.setDate(returnDate.getDate() -2)
    }
    if (returnDate.getDay() === 6) {
      returnDate.setDate(returnDate.getDate() -1)
    }
  }
  return returnDate

}

Utils.isDateInScope = function(date, startDateScope, endDateScope) {

  if (date >= startDateScope && date <= endDateScope) {
    return true
  }

  return false
}

/*Utils.isMonthInScope = function(month, aMonths) {
 for (var i = 0; i < aMonths.length; i++) {
 if (aMonths[i] === month+1) {
 return true
 }
 }
 return false
 }*/

Utils.isProjectExists = function(projectKey, aProjects) {
  for (var i = 0; i < aProjects.length; i++) {
    if (aProjects[i] === projectKey) {
      return true
    }
  }
  return false
}

Utils.isAccountExists = function(accountKey, aAccounts) {
  for (var i = 0; i < aAccounts.length; i++) {
    if (aAccounts[i] === accountKey) {
      return true
    }
  }
  return false
}

/*Utils.getMonthsInScopeArray = function() {
 var startDate = new Date(SpreadsheetApp.getActiveSpreadsheet().getRange('Rates by Team!K2').getValue()),
 endDate = new Date(SpreadsheetApp.getActiveSpreadsheet().getRange('Rates by Team!L2').getValue()),
 aMonths = []

 while(startDate < endDate) {
 aMonths.push(this.getMonthName(startDate.getMonth()) +' ' + startDate.getFullYear())
 startDate = Utils.getFirstDayOfNextMonth(startDate)
 }
 return aMonths
 }*/

Utils.getProjectsArray = function() {
  var fProjects = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects').getDataRange().getValues(),
    aProjects = []
  for (var i = 0; i < fProjects.length; i++) {
    aProjects.push(fProjects[i][0])
  }
  return aProjects
}

Utils.getAccountsArray = function() {
  var fAccounts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accounts').getDataRange().getValues(),
    aAccounts = []
  for (var i = 0; i < fAccounts.length; i++) {
    aAccounts.push(fAccounts[i][0])
  }
  return aAccounts
}

Utils.getLastDayOfMonth = function(date) {
  var d = new Date(date.getYear(), date.getMonth()+1, 0)
  return d
}

Utils.getFirstDayOfNextMonth = function(date) {
  var nextMonthDate
  if (date.getMonth() == 11) {
    nextMonthDate = new Date(date.getFullYear() + 1, 0, 1)
  } else {
    nextMonthDate = new Date(date.getFullYear(), date.getMonth() + 1, 1)
  }
  return nextMonthDate
}


Utils.getRate = function(username, team, accounting) {
  var rates = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rates by Team').getDataRange().getValues()
  for (var i = 1; i < rates.length; i++) {
    if (username === rates[i][1] && team === rates[i][2] && accounting === rates[i][3]) {
      return rates[i][4]
    }
  }
  return 'Not defined'
}

Utils.getRateByAccountActuals = function(username, accountKey) {
  var rates = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rates by Account').getDataRange().getValues()
  for (var i = 1; i < rates.length; i++) {
    if (username === rates[i][1] && accountKey === rates[i][2]) {
      return rates[i][3]
    }
  }
  return 'Not defined'
}

Utils.getPriceByRate = function(rate, accounting) {
  var sPrices = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rate Card').getDataRange().getValues(),
    discountAccounting = SpreadsheetApp.getActiveSpreadsheet().getRange('Rate Card!D2').getValue()
  for (var i = 1; i < sPrices.length; i++) {
    if (rate === sPrices[i][0]) {
      if (accounting === discountAccounting) {
        return sPrices[i][2]
      } else {
        return sPrices[i][1]
      }
    }
  }
}

Utils.isDefaultRateForConsultant = function(username, rate) {
  var sResources = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resources').getDataRange().getValues()
  for (var i = 1; i < sResources.length; i++) {
    if (username === sResources[i][1]) {
      if (rate === sResources[i][4]) {
        return true
      }
    }
  }
  return false
}

// Get project name by JIRA project code from sheet Projects
Utils.getProjectByCode = function(code) {
  var sProjects = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects').getDataRange().getValues()
  for (var i = 1; i < sProjects.length; i++) {
    if (code === sProjects[i][0]) {
      return sProjects[i][1]
    }
  }
  return 'Not defined'
}

// Get account key name by Tempo billable field
Utils.getBillableByKey = function(key) {
  var sAccounts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Accounts').getDataRange().getValues()
  for (var i = 1; i < sAccounts.length; i++) {
    if (key === sAccounts[i][0]) {
      return sAccounts[i][2]
    }
  }
  return 'Not defined'
}

// Get profile by username from sheet Resources
Utils.getProfile = function(username) {
  var sConsultants = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resources').getDataRange().getValues()
  for (var i = 1; i < sConsultants.length; i++) {
    if (username === sConsultants[i][1]) {
      return sConsultants[i][2]
    }
  }
  return 'Not defined'
}

// Get organisation by username from sheet Resources
Utils.getOrg = function(username) {
  var sConsultants = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resources').getDataRange().getValues()
  for (var i = 1; i < sConsultants.length; i++) {
    if (username === sConsultants[i][1]) {
      return sConsultants[i][3]
    }
  }
  return 'Not defined'
}