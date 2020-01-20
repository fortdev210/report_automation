var months = ["January", "February", "March", "Aprill", "May", "June", "July", "August", "September", "Octobor", "November", "December"];
// Add Custom Item In Menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Auto')
    .addItem('Monthly Report and Actual Hours', 'TotalAutomation')
    .addItem('Es Vs Actual Monthly Report', 'monthlyReport')
    .addItem('Es Vs Actual Hours Report', 'ActualHoursReport')
    .addToUi();
}

function TotalAutomation()
{
  monthlyReport();
  ActualHoursReport();
}
function IndexofSubstring(s,text)
{
  var lengths = s.length;
  var lengtht = text.length;
  for (var i = 0;i < lengths - lengtht + 1;i++)
  {
    if (s.substring(i,lengtht + i) == text)
      return i;
  }
  return -1;
}

function monthlyReport() {
 var source = SpreadsheetApp.getActiveSpreadsheet();
 var sheets = source.getSheets();
 var sheetNameArray = [];
 var k = -1;
 for (var i=0 ; i<sheets.length ; i++) 
 {
    sheetNameArray.push(sheets[i].getName());
    if (IndexofSubstring( sheets[i].getName(), 'Tasks' ) != -1)
       for (var j = 0 ; j < months.length-1; j++) 
       {
         if (IndexofSubstring(sheets[i].getName(), months[j]) != -1) 
         {
           if (j > k) k = j;
         } 
       }
 } 
  var SelectedMonth = '';
  if ( k == -1 ) SelectedMonth = months[11].concat(' Tasks'); 
  else SelectedMonth = months[k].concat(' Tasks');
  var CurrentMonth = Utilities.formatDate(new Date(), 'PST', 'MMMM');
  var CurrentMonthtasks = CurrentMonth.concat(' Tasks');
  var sheet = source.getSheetByName(SelectedMonth);
  if ( sheetNameArray.indexOf(CurrentMonthtasks) == -1) 
  {
    sheet.copyTo(source).setName(CurrentMonthtasks);
    sheet = source.getSheetByName(CurrentMonthtasks);
    sheet.getRange('A2:C').clearContent(); 
    var formula = sheet.getRange("D2");
    var formula_node = formula.getFormula().split("+");
    var addedformularnode = formula_node[formula_node.length-1];
    for (i = 0 ; i < months.length; i++) 
    {
      if (IndexofSubstring(addedformularnode, months[i]) != -1) 
      {
        var node = addedformularnode.replace(new RegExp(months[i],'gi'), CurrentMonth);
        break;
      } 
    }
    var updatedformula = formula.getFormula().concat('+').concat(node);
    //var activesheet = source.getActiveSheet();
    formula.setFormula(updatedformula);
    var lastrow = sheet.getLastRow();
    var fillDownRange = sheet.getRange(2, 4, lastrow-1);
    formula.copyTo(fillDownRange);
  }
}

function ActualHoursReport(){
 var source = SpreadsheetApp.getActiveSpreadsheet();
 var sheets = source.getSheets();
 var sheetNameArray = [];
 var k = -1;
 for (var i=0 ; i<sheets.length ; i++) 
 {
    sheetNameArray.push(sheets[i].getName());
    if (IndexofSubstring( sheets[i].getName(), 'Est vs Actual Hours' ) != -1)
       for (var j = 0 ; j < months.length-1; j++) 
       {
         if (IndexofSubstring(sheets[i].getName(), months[j]) != -1) 
         {
           if (j > k) k = j;
         } 
       }
 } 
 var SelectedMonth = '';
 if ( k == -1 ) SelectedMonth = "Est vs Actual Hours ".concat(months[11]);
 else SelectedMonth = "Est vs Actual Hours ".concat(months[k]);
 var CurrentMonth = Utilities.formatDate(new Date(), 'PST', 'MMMM');
 var actualHoursCurrentMonth = "Est vs Actual Hours ".concat(CurrentMonth);
 var sheet = source.getSheetByName(SelectedMonth);
 if ( sheetNameArray.indexOf(actualHoursCurrentMonth) == -1)
  {
    sheet.copyTo(source).setName(actualHoursCurrentMonth);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(actualHoursCurrentMonth);
    sheet.getRange('A2:C').clearContent(); 
    var date = new Date();
    var year = date.getYear();
    sheet.getRange('D1').setValue(CurrentMonth.toString()+'/1/'+year.toString());
    sheet.getRange('AJ1').setValue(CurrentMonth.toString()+' Total');
    sheet.insertColumns(37, 1);
    var valuestocopy = sheet.getRange("AL1:AL");
    valuestocopy.copyTo(sheet.getRange("AK1:AK"));
    var pastmonth = months[CurrentMonth-1];
    if (!pastmonth) pastmonth = "December";
    sheet.getRange("AK1").setValue(pastmonth + " Total");
    
    var formula = sheet.getRange("AK2");
    var formula_node = formula.getFormula();
    for (i = 0 ; i < months.length; i++) 
    {
      if (IndexofSubstring(formula_node, months[i]) != -1) 
      {
        var node = formula_node.replace(new RegExp(months[i],'gi'), pastmonth);
        break;
      } 
    }
    formula.setFormula(node);   
//    var activesheet = source.getActiveSheet();
    var lastrow = sheet.getLastRow();
    var fillDownRange = sheet.getRange(2, 37, lastrow-1);
    formula.copyTo(fillDownRange);
    
    var formula1 = sheet.getRange("D2");
    var formula_node1 = formula1.getFormula();
    var node1 = formula_node1.replace(new RegExp(months[k-1],'gi'), CurrentMonth);
    formula1.setFormula(node1);   
    var getDayRange = sheet.getRange("D2:AH2");
    var fillDownRange1 = sheet.getRange(2, 4, lastrow-1,getDayRange.getValues()[0].length);
    formula1.copyTo(fillDownRange1);
    
    var formula2 = sheet.getRange(2,getByName("Actual",1, actualHoursCurrentMonth)+1);
    var formula_node2 = formula2.getFormula();
    var node2 = formula_node2.replace(new RegExp(months[k-1],'gi'), CurrentMonth);
    formula2.setFormula(node2);
    var fillDownRange2 = sheet.getRange(2, getByName("Actual",1, actualHoursCurrentMonth)+1, lastrow-1);
    formula2.copyTo(fillDownRange2);
    
    var formula3 = sheet.getRange(1,getByName("Verify",1, actualHoursCurrentMonth)+2);
    var formula_node3 = formula3.getFormula();
    var node3 = formula_node3.replace(new RegExp(months[k-1],'gi'), CurrentMonth);
    formula3.setFormula(node3);
  }
}

function getByName(colName, row, actualHoursMonth) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(actualHoursMonth);
  var data = sheet.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  if (col != -1) {
    return col;
  }
}
function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}