// Add Custom Item In Menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
    .addItem('Daily Report', 'dailyReport')
    .addItem('Monthly Report', 'monthlyReport')
    .addToUi();
}
// Project Detail Tracking Sheet URL
var mainUrl = '1Msj3CmSd7BI9vxdqQ44vgKIBzdN7sBGa5CZfmtKkWpA';

// Get Current Month
var month = Utilities.formatDate(new Date(), 'PST', 'MMMM');
// Get Sheet By Name
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Worksheet");

function formatSheet() {
  var range = sheet.getDataRange();
  var rangeVals = range.getValues();
  // Delete Unnecessary Columns
  if (rangeVals[0].length > 17) {
    sheet.deleteColumns(1, 2); // Delete A, B Columns
    sheet.deleteColumns(2, 17); // Delete D ~ T Columns
    sheet.deleteColumn(3); // Delete V Column
    sheet.deleteColumns(4, 2); // Delete X, Y Columns

    // Beautify Sheet Format
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 1000);
    sheet.setColumnWidth(3, 100);

    // Rearrange Columns Order
    var columnSpec = sheet.getRange("A1:A");
    sheet.moveColumns(columnSpec, 4);
    var columnSpec = sheet.getRange("B1:B");
    sheet.moveColumns(columnSpec, 4);
  }
}

function deleteRows() {
  // Remove Unnecessary Rows
  var delAngelaVal = "CH Internal-Angela Harper"; // Delete Value In "Project" Column 
  var delPMVal1 = "PM Activities"; // Delete Value In Name Column
  var delPMVal2 = "Project Management"; // Delete Value In Name Column

  var range = sheet.getDataRange();
  var rangeVals = range.getValues();
  for (var i = rangeVals.length - 1; i >= 0; i--) {
    // Delete PM Activities, Project Management(Name), or CH Internal-Anglea Harper(Project)
    if (rangeVals[i][0].toString().toLowerCase() === delPMVal1.toLowerCase() ||
      rangeVals[i][0].toString().toLowerCase() === delPMVal2.toLowerCase() ||
      rangeVals[i][1].toString().toLowerCase() === delAngelaVal.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }
}

function dailyReport() {
  formatSheet();
  deleteRows();
}