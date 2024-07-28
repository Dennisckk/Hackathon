function getData() {
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Access "Progression" sheet and retrieve data
  var progressionSheet = ss.getSheetByName("Progression");
  var progressionDataRange = progressionSheet.getDataRange();
  var progressionValues = progressionDataRange.getValues();

  // Access "Sheet1" (task details) sheet and retrieve data
  var taskSheet = ss.getSheetByName("Sheet1");
  var taskDataRange = taskSheet.getDataRange();
  var taskValues = taskDataRange.getValues();

  // Access "Home" (staff details) sheet and retrieve data
  var assigneeSheet = ss.getSheetByName("Home");
  var assigneeDataRange = assigneeSheet.getDataRange();
  var assigneeValues = assigneeDataRange.getValues();

  return {
    progressionValues: progressionValues,
    taskValues: taskValues,
    assigneeValues: assigneeValues
  };
}

function checkNA(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName("Home");
  if (sheet) {
    var lastRow = sheet.getLastRow();
    var dataRange = sheet.getRange(1, 1, lastRow, 8);
    var dataValues = dataRange.getValues();

    for (var row = 0; row < dataValues.length; row++) {
      for (var col = 0; col < dataValues[row].length; col++) {
        if (dataValues[row][col] === '') {
          dataValues[row][col] = 'NA';
        }
      }
    }
    dataRange.setValues(dataValues);
    Logger.log('Number of existing data entries in column A: ' + lastRow);
  } else {
    Logger.log('Sheet "Home" not found.');
  }
}
