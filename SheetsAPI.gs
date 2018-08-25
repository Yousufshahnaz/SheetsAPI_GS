//********************************************************************
//Google Sheets Functions
//********************************************************************

//Return the Sheet object with 'name'
//*******************************************
function getSheet(spreadsheet_id,name){
  return SpreadsheetApp.openById(spreadsheet_id).getSheetByName(name);
}

//Return the value of a given cell
//*******************************************
function getValue(sheet,row,column){
  return sheet.getRange(row,column).getValue();
}


//Return the row number of the cell matching
//the given 'value' in 'column'
//*******************************************
function rowLookup(sheet, value, column) {
  var columnValues = sheet.getRange(1,column,sheet.getLastRow()).getValues();
  for (var i=0; i<columnValues.length; i++) {
    if (columnValues[i] == value) {
      var rowNum = i+1;
      return rowNum;
    }
  }
  return columnValues;
}

//Return the column number of the cell matching
//the given 'value' in 'row'
//*******************************************
function colLookup(sheet, value, row) {
  var rowValues = sheet.getRange(row,1,row,sheet.getLastColumn()).getValues();
  for (var i=0; i<columnValues.length; i++) {
    if (rowValues[i] == value) {
      var colNum = i+1;
      return colNum;
    }
  }
  return rowValues;
}

//Return an array with each element the cell value
//for each column (L-R) for a given row.  Must
//indicate whether or not to compress empty cells.
//*******************************************

function rowValues(sheet, row, cws) {
  var rowArr = [];
  var cellValue;
  for (i=1;i<(sheet.getLastColumn()+1); i++) {
    cellValue = sheet.getRange(row,i).getValue();
    if (cws == "y" || cws == "Y" || cws == "yes" || cws == "Yes") {
      if(cellValue != "") {rowArr.push(cellValue);}
    } else {
      if(cellValue == "") {rowArr.push("");} else {rowArr.push(cellValue);}
    }
  }
  return rowArr;
}

//Return an array with each element the cell value
//for each column (L-R) for a given column.  Must
//indicate whether or not to compress empty cells.
//*******************************************

function colValues(sheet, column, cws) {
  var rowArr = [];
  var cellValue;
  for (i=1; i<(sheet.getLastRow()+1); i++) {
    cellValue = sheet.getRange(i,column).getValue();
    if (cws == "y" || cws == "Y" || cws == "yes" || cws == "Yes") {
      if(cellValue != null) {rowArr.push(cellValue);}
    } else {
      if(cellValue == null) {rowArr.push("");} else {rowArr.push(cellValue);}
    }
  }
  return rowArr;
}



//Append row with values taken from 'values' array.
//Fills from A->ZZ
//*******************************************
function appendRow(sheet,values) {
  var fillArr = [];
  for (var i=0; i<values.length; i++) {
    fillArr.push(values[i]);
  }
  sheet.appendRow(fillArr);
}

//Return the value of the given script property
//Standard Utility
//*******************************************
function getProperty(propertyName){
  return PropertiesService.getScriptProperties().getProperty(propertyName);
}
