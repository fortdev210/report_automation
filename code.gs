// Add Custom Item In Menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
    .addItem('Daily Hours Logged Report', 'logReport')
    .addItem('Daily Hours New Month Creation', 'newMonthCreate')
    .addItem('Es Vs Actual Daily Report', 'dailyReport')
    .addItem('Es Vs Actual Monthly Report', 'monthlyReport')
    .addToUi();
}
// Project Detail Tracking Sheet URL
var mainUrl = '1Msj3CmSd7BI9vxdqQ44vgKIBzdN7sBGa5CZfmtKkWpA';

// Get Current Month
var month = Utilities.formatDate(new Date(), 'PST', 'MMMM');

function formatSheet(sheet) {
  var range = sheet.getDataRange();
  var rangeVals = range.getValues();
  // Delete Unnecessary Columns
  if (rangeVals[0].length > 17) {
    sheet.deleteColumns(1, 2); // Delete A, B Columns
    sheet.deleteColumns(2, 17); // Delete D ~ T Columns
    sheet.deleteColumn(3); // Delete V Column
    sheet.deleteColumns(4, 2); // Delete X, Y Columns

    // Beautify Sheet Format
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 600);
    sheet.setColumnWidth(3, 100);

    // Rearrange Columns Order
    var columnSpec = sheet.getRange("A1:A");
    sheet.moveColumns(columnSpec, 4);
    var columnSpec = sheet.getRange("B1:B");
    sheet.moveColumns(columnSpec, 4);
  }
}

function deleteRows(sheet) {
  var range = sheet.getDataRange();
  var rangeVals = range.getValues();
  // Remove Unnecessary Rows
  var delAngelaVal = "CH Internal-Angela Harper"; // Delete Value In "Project" Column 
  var delPMVal1 = "PM Activities"; // Delete Value In Name Column
  var delPMVal2 = "Project Management"; // Delete Value In Name Column

  for (var i = rangeVals.length - 1; i >= 0; i--) {
    // Delete PM Activities, Project Management(Name), or CH Internal-Anglea Harper(Project)
    if (rangeVals[i][0].toString().toLowerCase() === delPMVal1.toLowerCase() ||
      rangeVals[i][0].toString().toLowerCase() === delPMVal2.toLowerCase() ||
      rangeVals[i][1].toString().toLowerCase() === delAngelaVal.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }
}

function fixSameNameDifferentProject(sheet) {
  var rangeVals = sheet.getRange('A2:C').sort(1).getValues();
  var rowLength = sheet.getLastRow() - 1;

  // Get Duplicate Names In Array
  var dupArray = [];
  var flag = false;
  for (var i = 0; i < rowLength - 1; i++) {
    if (rangeVals[i][0] == rangeVals[i + 1][0] && rangeVals[i][1] != rangeVals[i + 1][1]) {
      if (flag == false) {
        dupArray.push(i + 2);
        dupArray.push(i + 3);
        flag = true;
      } else {
        dupArray.push(i + 3);
      }
    } else {
      flag = false;
    }
  }
  Utilities.sleep(3000);
  // Fix Duplicate Names By Adding Hypen & Project Name (i.e Landing Page -> Landing Page-Image3D)
  for (var i = 0; i < dupArray.length; i++) {
    var index = dupArray[i];
    sheet.getRange(index, 1).setBackground('yellow');
    sheet.getRange(index, 1).setValue(rangeVals[index - 2][0] + '-' + rangeVals[index - 2][1]);
  }
}

function fixSameNameSameProject(sheet) {
  var rangeVals = sheet.getRange('A2:C').sort(1).getValues();
  var rowLength = sheet.getLastRow() - 1;

  // Get Duplicate Names In Array
  var dupArray = [];
  var flag = false;
  for (var i = 0; i < rowLength - 1; i++) {
    if (rangeVals[i][0] == rangeVals[i + 1][0] && rangeVals[i][1] == rangeVals[i + 1][1]) {
      if (flag == false) {
        dupArray.push(i + 2);
        dupArray.push(i + 3);
        flag = true;
      } else {
        dupArray.push(i + 3);
      }
    } else {
      flag = false;
    }
  }
  Utilities.sleep(3000);
  // Fix Duplicate Names By Adding Hypen & Project Name (i.e Landing Page -> Landing Page-Image3D)
  for (var i = 0; i < dupArray.length; i++) {
    var index = dupArray[i];
    sheet.getRange(index, 1).setBackground('yellow');
    sheet.getRange(index, 1).setValue(rangeVals[index - 2][0] + '-' + (i + 1));
  }
}

function modifyDuplicateValues(sheet) {
  var rangeVals = sheet.getRange('A2:C').sort(1).getValues();
  var rowLength = sheet.getLastRow() - 1;

  // Get Duplicate Names In Array
  var dupArray = [];
  var flag = false;
  for (var i = 0; i < rowLength - 1; i++) {
    if (rangeVals[i][0] == rangeVals[i + 1][0]) {
      if (flag == false) {
        dupArray.push(i + 2);
        dupArray.push(i + 3);
        flag = true;
      } else {
        dupArray.push(i + 3);
      }
    } else {
      flag = false;
    }
  }
  Utilities.sleep(3000);
  // Fix Duplicate Names By Adding Hypen & Project Name (i.e Landing Page -> Landing Page-Image3D)
  for (var i = 0; i < dupArray.length; i++) {
    var index = dupArray[i];
    sheet.getRange(index, 1).setBackground('yellow');
    sheet.getRange(index, 1).setValue(rangeVals[index - 2][0] + '-' + (i + 1));
  }
}

// function compare(sheet) {

// }

function dailyReport() {
  // Get Sheet By Name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("assignment_filter");
  formatSheet(sheet);
  deleteRows(sheet);
  // fixSameNameDifferentProject(sheet);
  // fixSameNameSameProject(sheet);
  modifyDuplicateValues(sheet);
}

function logReport () {
  return;
}

function monthlyReport () {
  return;
}