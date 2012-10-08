/*
 * Copyright "Haridas N" <haridas.nss@gmail.com>

     This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
     any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>
   
 */



var SPREAD_SHEET_ID = '0AsN9BpYRL9hmdDcxVzlsYkdBMmZ4MmtkTWwyTDRMWmc';

var MONTH_NAME = {
  0:"JANUARY",
  1:"FEBVARY",
  2:"MARCH",
  3:"APRIL",
  4:"MAY",
  5:"JUNE",
  6:"JULY",
  7:"AUGUST",
  8:"SEPTEMBER",
  9:"OCTOBER",
  10:"NOVEMBER",
  11:"DECEMBER"
  
}
    
// The onOpen function is executed automatically every time a Spreadsheet is opened
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // When the user selects "addMenuExample" menu, and clicks "Menu Entry 1", the function function1 is executed.
  menuEntries.push({name: "Get Aggr. Sum", functionName: "runExample"});
  ss.addMenu("Results", menuEntries);
}

function myFunction() {
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu_entries = [];
  
  menu_entries.push({'name': "Script testing", "functionName": 'fun1'});
  ss.addMenu("Menu Example",menu_entries);
}



function runExample() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  // Get the range of cells that store employee data.
  var employeeDataRange = ss.getRangeByName("expenses");

  // For every row of employee data, generate an employee object.
  var expenseRows = getRowsData(sheet, employeeDataRange);

  
  
  //var stringToDisplay = "The third row is: " + thirdEmployee.timestamp + " " + thirdEmployee.date;
  //stringToDisplay += " (Name : " + thirdEmployee.name + "\n ";
  //stringToDisplay += "\nCost:  "+ thirdEmployee.cost;
  //stringToDisplay += "\n Paid: " + thirdEmployee.paid;
  
  
  
 
  
  //var first_row = employeeObjects[0];
  
  
  
  // Hold the Monthley report for all expenses.
  // {'month': {'name1':[], 'name2':[],..}, 'month2':{}, ..}
  // 
  
  var expense_report = {};
  
  var text = ""
  //for(i in employeeObjects){
    
    //text += employeeObjects[i].name + " " + employeeObjects[i].cost + " " + employeeObjects[i].paid + " " + employeeObjects[i].timestamp.getMonth();
    
    //text += employeeObjects[i].timestamp.getMonth()
  //}

  
  for(var row in expenseRows){
    var curr_row = expenseRows[row];
    var curr_month = MONTH_NAME[curr_row.timestamp.getMonth()];
    
    
    if(!expense_report[curr_month]){
      expense_report[curr_month] = {};
    }
    
    if(!expense_report[curr_month][curr_row.name]){
      expense_report[curr_month][curr_row.name] = {};
      
      // Create the rows first time for a user.
      var result_row = {'name':curr_row.name,'totalPaid':curr_row.paid,'totalExpense':curr_row.cost};
      result_row['net'] = curr_row.paid - curr_row.cost;
      
      expense_report[curr_month][curr_row.name] = result_row;
      
    }else{
      var user = expense_report[curr_month][curr_row.name];
      var total_paid =  user.totalPaid;
      var total_expense = user.totalExpense;
      
      user.totalPaid = total_paid + curr_row.paid;
      user.totalExpense = total_expense + curr_row.cost;
      user.net = user.totalPaid - user.totalExpense;
      
      expense_report[curr_month][curr_row.name] = user;
    }
    
  }
  
  // Converting dict format to single row compatible to write back to excel sheet.
  var text = "";
  var results = [];
  for(var month in expense_report){
   
    var flag = true;
    for(var user in expense_report[month]){
      
      var report = expense_report[month][user];
      
      if(flag == true){
        report['month'] = month;
        flag = false;
      }else{
       report['month'] = "";
      }
      
     results.push(report);
      
    }
    
  }
  
  
  //for(res in results){
  //  text += results[res].month + " " + results[res].name + " " + results[res].totalPaid + " " + results[res].totalExpense + " " + results[res].net;
  //}
  
  //Browser.msgBox(text);
  
  WriteAggregateData(ss,results);
  
  
}


// WriteAggregateData is a function which calculate monthely expense for food and give aggreate
// information about how much paid and how much need to pay in monthely basis.

function WriteAggregateData(ss,results){
  var columnHeaders = ['Month','Name','Total Paid','Total Expense','Net'];
  
  var sheet = ss.getSheetByName("Result") || ss.insertSheet("Result");
  
  //sheet.clear();
  
  //Pick background color from 1,1 column.
  var headerBackgroundColor = sheet.getRange(1,1).getBackgroundColor();
  
  // Pick the area from the ;heet where we want to write.
  var headerRange = sheet.getRange(1, 1, 1, 5);
  
  // Set Row headers.
  headerRange.setValues([columnHeaders]);
  
  
  headerRange.setBackgroundColor(headerBackgroundColor);
  
  // Write data to sheet.
  setRowsData(sheet,results,headerRange);
  
}



// setRowsData fills in one row of data per object defined in the objects Array.
// For every Column, it checks if data objects define a value for it.
// Arguments:
//   - sheet: the Sheet Object where the data will be written
//   - objects: an Array of Objects, each of which contains data for a row
//   - optHeadersRange: a Range of cells where the column headers are defined. This
//     defaults to the entire first row in sheet.
//   - optFirstDataRowIndex: index of the first row where data should be written. This
//     defaults to the row immediately below the headers.
function setRowsData(sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  var headers = normalizeHeaders(headersRange.getValues()[0]);

  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < headers.length; ++j) {
      var header = headers[j];
      values.push(header.length > 0 && objects[i][header] ? objects[i][header] : "");
    }
    data.push(values);
  }

  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(), 
                                        objects.length, headers.length);
  destinationRange.setValues(data);
}


// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}


// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}




// Exercise:


function runExercise() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];

  // Get the range of cells that store employee data.
  var employeeDataRange = sheet.getRange("B1:F5");

  // For every row of employee data, generate an employee object.
  var employeeObjects = getColumnsData(sheet, employeeDataRange);

  var thirdEmployee = employeeObjects[2];
  var stringToDisplay = "The third column is: " + thirdEmployee.firstName + " " + thirdEmployee.lastName;
  stringToDisplay += " (id #" + thirdEmployee.employeeId + ") working in the ";
  stringToDisplay += thirdEmployee.department + " department and with phone number ";
  stringToDisplay += thirdEmployee.phoneNumber;
  ss.msgBox(stringToDisplay);
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}

// getColumnsData iterates column by column in the input range and returns an array of objects.
// Each object contains all the data for a given column, indexed by its normalized row name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - rowHeadersColumnIndex: specifies the column number where the row names are stored.
//       This argument is optional and it defaults to the column immediately left of the range; 
// Returns an Array of objects.
function getColumnsData(sheet, range, rowHeadersColumnIndex) {
  rowHeadersColumnIndex = rowHeadersColumnIndex || range.getColumnIndex() - 1;
  var headersTmp = sheet.getRange(range.getRow(), rowHeadersColumnIndex, range.getNumRows(), 1).getValues();
  var headers = normalizeHeaders(arrayTranspose(headersTmp)[0]);
  return getObjects(arrayTranspose(range.getValues()), headers);
}



function fun1(){
 
  // Working Rules.
  
  var time = Date();
}
