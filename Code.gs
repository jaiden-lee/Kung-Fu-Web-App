let sheetID = "REDACTED";
let spreadsheet = SpreadsheetApp.openById(sheetID);
//https://gist.github.com/teddyleung/3f085bc1cc4b4b21980c4b05589bc392?adlt=strict&toWww=1&redig=4C9A4E97471147BD988D3329F09360C1

let range = "Sheet1";
let studentList = {};

let monthList = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function checkToAddNewMonth () {
  let date = getLatestMonth();
  let month = date.substring(0,date.length-5).trim();
  let currentMonth = (new Date()).getMonth();
  let monthIndex = 99;
  for (let i=0;i<monthList.length;i++) {
    if (monthList[i]==month) {
      monthIndex = i;
    }
  }
  for (let i=monthIndex;i<currentMonth;i++) {
    addNewMonth();
  }
}

function addNewMonth () {
  let date = getLatestMonth();
  let month = date.substring(0, date.length-5).trim();
  let year = date.substring(date.length-4).trim();
  let monthIndex = -1;
  for (let i=0;i<monthList.length;i++) {
    if (monthList[i]==month) {
      monthIndex = i;
    }
  }
  if (monthIndex==-1) {
    return false;
  }
  let newMonth = monthIndex+1;
  let newYear = year;
  if (newMonth>=12) {
    newMonth = 0;
    newYear++;
  }
  let newMonthText = monthList[newMonth];
  addMonthHelper(newMonthText, newYear);
  return true;
}

function getLatestMonth () {
  let lastColumn = spreadsheet.getLastColumn()-3;
  let cellRange = "Sheet1!"+R1C1toA1(9, lastColumn)+":"+R1C1toA1(9, lastColumn+2);
  let values = spreadsheet.getRange(cellRange).getValues()[0];
  return values[0]+" "+values[2];
}

function getMonthsList () {
  let lastColumn = spreadsheet.getLastColumn();
  let values = spreadsheet.getRange("Sheet1!"+R1C1toA1(9, 1)+":"+R1C1toA1(9, lastColumn-1)).getValues()[0];
  let months = [];
  for (let i=1;i<values.length;i+=3) { //"MONTH_YEAR"
    let month = values[i].trim();
    let year = values[i+2];
    let object = {
      text: month+" "+year,
      rowNumber: i+1
    };
    months.push(object);
  }
  return months;
}

function addMonthHelper (month, year) {
  let today = new Date();
  // let month = monthList[today.getMonth()];
  // let year = today.getFullYear();
  let lastColumn = spreadsheet.getLastColumn();
  spreadsheet.insertColumnBefore(lastColumn);
  spreadsheet.insertColumnBefore(lastColumn);
  spreadsheet.insertColumnBefore(lastColumn);
  spreadsheet.setColumnWidth(lastColumn, spreadsheet.getColumnWidth(2));
  spreadsheet.setColumnWidth(lastColumn+1, spreadsheet.getColumnWidth(2));
  spreadsheet.setColumnWidth(lastColumn+2, spreadsheet.getColumnWidth(4));
  let cellRange = "Sheet1!"+R1C1toA1(9, lastColumn-3)+":"+R1C1toA1(spreadsheet.getLastRow(), lastColumn-1);
  spreadsheet.getRange(cellRange).copyFormatToRange(spreadsheet.getSheetId(), lastColumn, lastColumn+2, 9, spreadsheet.getLastRow());
  let cellRange2 = "Sheet1!"+R1C1toA1(9, lastColumn-3)+":"+R1C1toA1(11, lastColumn-1);
  spreadsheet.getRange(cellRange2).copyValuesToRange(spreadsheet.getSheetId(), lastColumn, lastColumn+2, 9, 11);
  spreadsheet.getRange("Sheet1!"+R1C1toA1(9, lastColumn)+":"+R1C1toA1(9, lastColumn)).setValue(month);
  spreadsheet.getRange("Sheet1!"+R1C1toA1(9, lastColumn+2)+":"+R1C1toA1(9, lastColumn+2)).setValue(year);
}

function addStudent (studentName) {
  let students = getNamesList();
  if (students[studentName]!=null) {
    return false;
  }
  let lastRow = spreadsheet.getLastRow();
  let lastColumn = spreadsheet.getLastColumn();
  spreadsheet.appendRow([studentName]);
  spreadsheet.appendRow([" "]);
  let cellRange = "Sheet1!"+R1C1toA1(lastRow-1,1)+":"+R1C1toA1(lastRow,lastColumn);
  spreadsheet.getRange(cellRange).copyFormatToRange(spreadsheet.getSheetId(), 1, lastColumn, lastRow+1, lastRow+2);
  lastRow = spreadsheet.getLastRow();
  let changeRange = "Sheet1!"+R1C1toA1(lastRow, 1)+":"+R1C1toA1(lastRow, 1);
  // spreadsheet.getRange(changeRange).setValue(" ");
  return true;
}

function getNamesList () {
  let lastRow = spreadsheet.getLastRow();
  let values = spreadsheet.getRange("Sheet1!A1:A"+lastRow).getValues();

  let namesList = {};
  for (let rowNumber=11; rowNumber<values.length;rowNumber+=2) {
    namesList[values[rowNumber][0]] = rowNumber;
  }
  studentList = namesList;
  return namesList;
}

function getRemarks (rowNumber) {
  let lastColumn = spreadsheet.getLastColumn();
  let cellCoordinates = R1C1toA1(rowNumber+1, lastColumn);
  let cellRange = "Sheet1!"+cellCoordinates+":"+cellCoordinates;
  return spreadsheet.getRange(cellRange).getValue();
}

function updateRemarks (rowNumber, message) {
  let lastColumn = spreadsheet.getLastColumn();
  let cellCoordinates = R1C1toA1(rowNumber+1, lastColumn);
  let cellRange = "Sheet1!"+cellCoordinates+":"+cellCoordinates;
  spreadsheet.getRange(cellRange).setValue(message);
}

function updateMonthlyDues (rowNumber, paymentMethod, paymentId, month) {
  let cellCoordinates = R1C1toA1(rowNumber+2, parseInt(month)+2);
  let cellRange = "Sheet1!"+cellCoordinates+":"+cellCoordinates;
  let cellValue = "";

  if (paymentMethod=="Paypal") {
    cellValue="P";
  } else {
    cellValue="M";
  }
  cellValue+=" "+paymentId;

  // spreadsheet.getRange(cellValue).setValue(paymentValue);

  let redColor = SpreadsheetApp.newTextStyle()
    .setForegroundColor("red").build();

  let blueColor = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#4a86e8").build();

  let richTextValue = SpreadsheetApp.newRichTextValue()
    .setText(cellValue)
    .setTextStyle(0,1, redColor)
    .setTextStyle(1, cellValue.length, blueColor)
    .build();
  Logger.log(richTextValue);
  Logger.log(cellRange);
  spreadsheet.getRange(cellRange).setRichTextValue(richTextValue);
}

function getMonthlyDues (rowNumber, month) {
  let cellCoordinates = R1C1toA1(rowNumber+2, parseInt(month)+2);
  let cellRange = "Sheet1!"+cellCoordinates+":"+cellCoordinates;
  return spreadsheet.getRange(cellRange).getValue();
}

function getWeekAttendance (rowNumber, weekNumber, month) {
  lastColumn = spreadsheet.getLastColumn(); // this gives us the index of the last column (starting from 1)
  if (month!=-1) {
    lastColumn = month+3;
  }

  setWeekNumbers: {
    if (weekNumber==1) {
      lastColumn-=3;
    } else if (weekNumber==2) {
      lastColumn-=3;
      rowNumber+=1;
    } else if (weekNumber==3) {
      lastColumn-=2;
    } else if (weekNumber==4){
      lastColumn-=2;
      rowNumber+=1;
    } else if (weekNumber==5) {
      lastColumn-=1;
    }
  }

  let cellCoordinates = R1C1toA1(rowNumber+1, lastColumn);
  let cellRange = "Sheet1!"+cellCoordinates+":"+cellCoordinates;

  return [weekNumber, spreadsheet.getRange(cellRange).getValue()];
}

const R1C1toA1 = (row, column) => {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const base = chars.length;
  let columnRef = '';

  if (column < 1) {
    columnRef = chars[0];
  } else {
    let maxRoot = 0;
    while (base**(maxRoot + 1) < column) {
      maxRoot++;
    }

    let remainder = column;
    for (let root = maxRoot; root >= 0; root--) {
      const value = Math.floor(remainder / base**root);
      remainder -= (value * base**root);
      columnRef += chars[value - 1];
    }
  }

  // Use Math.max to ensure minimum row is 1
  return `${columnRef}${Math.max(row, 1)}`;
};

function updateAttendance (rowNumber, isAttendance, weekNumber, month) {
  // Determining what the range of the cell we are trying to update is
  lastColumn = spreadsheet.getLastColumn(); // this gives us the index of the last column (starting from 1)
  if (month!=-1) {
    lastColumn = month+3;
  }
  

  setWeekNumbers: {
    if (weekNumber==1) {
      lastColumn-=3;
    } else if (weekNumber==2) {
      lastColumn-=3;
      rowNumber+=1;
    } else if (weekNumber==3) {
      lastColumn-=2;
    } else if (weekNumber==4){
      lastColumn-=2;
      rowNumber+=1;
    } else if (weekNumber==5) {
      lastColumn-=1;
    }
  }

  let cellCoordinates = R1C1toA1(rowNumber+1, lastColumn);
  let cellRange = "Sheet1!"+cellCoordinates+":"+cellCoordinates;

  // Getting and Updating Cell Value (not updating the actual database yet)
  let cellValue = spreadsheet.getRange(cellRange).getValue();
  
  if (isAttendance) {
    if (cellValue.length>=1) {
      if (cellValue[cellValue.length-1]=="/") {
        cellValue = cellValue.substring(0,cellValue.length-1)+"X";
      } else {
        cellValue+="/";
      }
    } else {
      cellValue+="/";
    }
  } else {
    cellValue+="E";
  }

  // Updating the Database
  spreadsheet.getRange(cellRange).setValue(cellValue);
}


function myFunction() {
  updateMonthlyDues(11, "Paypal", 123456);
}

// this loads the HTML File
function doGet () {
  return HtmlService.createTemplateFromFile("index").evaluate();
}

// Like the HTML includes tag
function include(fileName) {
  return HtmlService.createTemplateFromFile(fileName).getRawContent();
}
