function onOpen(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Docs');
  menu.addItem('Create Progression List', 'createNewGoogleDocs');
  menu.addToUi();
}

function createNewGoogleDocs() {

  const parentFolder = DriveApp.getFolderById(DriveApp.getRootFolder().getId());
  const destinationFolder = parentFolder.createFolder("Progression Folder");

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet1 = ss.getSheetByName("Sheet1");
  var lrsheet1 = sheet1.getLastRow();
  var uniqueIDRange = sheet1.getRange(2, 1, lrsheet1 - 1).getValues();
  var assigneeIDRange = sheet1.getRange(2, 6, lrsheet1 - 1).getValues();
  var tasknameIDRange = sheet1.getRange(2, 2, lrsheet1 - 1).getValues();
  
  var progressionsheet = ss.getSheetByName("Progression");
  var lrprogsheet = progressionsheet.getLastRow();
  var progressionIDRange = progressionsheet.getRange(2, 1, lrprogsheet - 1).getValues();
  var taskIDRange = progressionsheet.getRange(2, 2, lrprogsheet - 1).getValues();

  var homesheet = ss.getSheetByName("Home");
  var lrhome = homesheet.getLastRow();
  var assigneeNameRange = homesheet.getRange(2, 4, lrhome - 1).getValues();
  var keyIDRange = homesheet.getRange(2, 1, lrhome - 1).getValues();

  const documentName = 'New Progression Docs';
  var files = destinationFolder.getFilesByName(documentName);

  if (files.hasNext()) {
    const existingFile = files.next();
    existingFile.setTrashed(true); 
  }

  const doc = DocumentApp.create('New Progression Docs'); 
  const docID = doc.getId();
  const file = DriveApp.getFileById(docID);
  destinationFolder.addFile(file);
  const body = doc.getBody();
  var headers = ['Progression ID', 'Task ID', 'Assignee Name'];
  var worklist = body.appendParagraph('Working Progression List');

  worklist.setFontFamily('Comfortaa');

  for(var i=0;i<lrsheet1-1;i++){

    var uniqueid = uniqueIDRange[i][0];
    var assigneeid = assigneeIDRange[i][0];
    var taskname = tasknameIDRange[i][0];
    body.appendParagraph(taskname + ' | Task ID: '+ uniqueid);
    const table = body.appendTable();

    var headerRow = table.appendTableRow();
      headers.forEach(function(header) {
      headerRow.appendTableCell(header);
    });
    var matchFound = false; 

    for(var j=0;j<lrprogsheet-1;j++){
      var progID = progressionIDRange[j][0];
      var taskID = taskIDRange[j][0];
      if (taskID === uniqueid) { 
        for (var k = 0; k < lrhome - 1; k++) {
          var assigneeName = assigneeNameRange[k][0];
          if (assigneeid === keyIDRange[k][0]) {
            var row = table.appendTableRow();
            row.appendTableCell(progID);
            row.appendTableCell(taskID);
            row.appendTableCell(assigneeName);
            matchFound = true; 
          }
        }
      }
    }

    if (!matchFound) {
    var emptyRow = table.appendTableRow();
      emptyRow.appendTableCell('');
      emptyRow.appendTableCell('');
      emptyRow.appendTableCell('');
    }
  }
  // doc.saveAndClose();
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



  // const googleDocTemplate = DriveApp.getFileById('1-VbwumC8OfqOlVAZxEByDiRWWUiteDEAnI5LFQsNcUs');
  // const copy = googleDocTemplate.makeCopy(documentName, destinationFolder);
  // var progressionArray = [];
  // var uniqueidarray = [];
  // var assigneeID = "";
  // var assigneeName = "";
  // var taskName = "";
  // for (var i = 0; i < lrprogsheet - 1; i++) {
  //   var progID = progressionIDRange[i][0];
  //   var taskID = taskIDRange[i][0];
  //   for (var j = 0; j < lrsheet1 - 1; j++) {
  //     if (uniqueIDRange[j][0] == taskID) {
  //       assigneeID = assigneeIDRange[j][0];
  //       taskName = tasknameIDRange[j][0];
  //       break;
  //     }
  //   }
  //   for (var k = 0; k < lrhome - 1; k++) {
  //     if (keyIDRange[k][0] == assigneeID) {
  //       assigneeName = assigneeNameRange[k][0];
  //       break;
  //     }
  //   }
  //   if (assigneeID && assigneeName) {
  //     progressionArray.push(progID + ":" + taskID + ":" + assigneeName + ":" + taskName);
  //   }
  // }
  // function convertIntoArray() {
  // // Call function
  //   var data =  getData();

  //   // Assignee section
  //   var assigneeArrays = []
  //   var assigneeNames = [];
  //   var assigneeIDs = [];

  //   data.assigneeValues.forEach(function(row) {
  //     assigneeArrays.push(row); // Convert data into array
  //   })

  //   assigneeArrays.forEach(function(row) {
  //     assigneeID = row[0];
  //     assigneeName = row[3];

  //     assigneeNames.push(assigneeName);
  //     assigneeIDs.push(assigneeID)
  //   })

  //   // Task section
  //   var taskArrays = [];
  //   var taskIDs = [];
  //   var taskNames = []
  //   var taskAssignees = [];

  //   data.taskValues.forEach(function(row) {
  //     taskArrays.push(row);
  //   })

  //   taskArrays.forEach(function(row) {
  //     taskID = row[0];
  //     taskName = row[1]
  //     taskAssignee = row[5];

  //     taskIDs.push(taskID);
  //     taskNames.push(taskName);
  //     taskAssignees.push(taskAssignee);

  //     // console.log(taskID + " : " + taskName + " : " + taskAssignee)
  //   })

  //   // Progression Section
  //   var progressionArrays = [];
  //   var progressionIDs = [];
  //   var progression_taskIDs = [];
  //   var progressionDetails = [];
  //   var completionDate = [];

  //   data.progressionValues.forEach(function(row) {
  //     progressionArrays.push(row);
  //   })

  //   progressionArrays.forEach(function(row) {
  //     var progressionID = row[0];
  //     var progression_taskID = row[1];
  //     var details = row[2];
  //     var date = row[3];

  //     progressionIDs.push(progressionID);
  //     progression_taskIDs.push(progression_taskID);
  //     progressionDetails.push(details);
  //     completionDate.push(date);

  //     // console.log(progressionID + " : " + taskID + " : " + details + " : " + date)
  //     //this.comparisionData();
  //   })
  // }





















