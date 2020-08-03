var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1glXwRHRWMnl6cpr-_nOYqBmAtdlmTpsMSYg8rNWYqg0/edit");

function getSheet(name) {
  return spreadsheet.getSheetByName(name);
}

function returnParamsDefault() {
  var length = getSheet("Reference").getRange("B1").getValue();
  var range = getSheet("Reference").getRange("B2:C" + length).getValues();
  return range;
}

function createFormula() {
  var length = getSheet("Delimited").getRange("A1").getValue();
  var range = getSheet("Delimited").getRange("A1:AK" + length);
  var strs = forEachRangeCell(range);
  getSheet("Logs").getRange("B2").setValue(strs[0]);
  getSheet("Logs").getRange("C2").setValue(strs[1]);
  getSheet("Logs").getRange("D2").setValue(strs[2]);
}

function forEachRangeCell(range) {
  var timeFrameSTR = "={";
  var rdevSTR = "={";
  var paramSTR = "={";
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  for (var i = 2; i <= numRows; i++) {
    const paramsCellVal = range.getCell(i, 1).getValue();
    var params = doubleSplit(paramsCellVal, ";", "=");
    for (var j = 2; j <= numCols; j++) {
      var newCellVal = range.getCell(i, j).getValue();
      if (newCellVal !== "") {
        timeFrameSTR += "Reference!A" + j + "; ";
        rdevSTR += (parseFloat(newCellVal) / 100).toString() + "; ";
        var localParams = returnParamsDefault();
        if (paramsCellVal !== "") {
          var count = 0;
          var toAdd = "{";
          for (var k = 0; k < localParams.length; k++) {
            if (localParams[k][0] === params[count][0]) {
              toAdd += params[count][1] + ", ";
              if (count < params.length - 1) {
                count++;
              }
            } else {
              toAdd += localParams[k][1] + ", ";
            }
          }
          paramSTR += toAdd.substring(0, toAdd.length - 2) + "}; ";
        } else {
          var toAdd = "{";
          for (var k = 0; k < localParams.length; k++) {
            toAdd += localParams[k][1] + ", ";
          }
          paramSTR += toAdd.substring(0, toAdd.length - 2) + "}; ";
        }
      }
    }
  }
  timeFrameSTR = timeFrameSTR.substring(0, timeFrameSTR.length - 2) + "}";
  rdevSTR = rdevSTR.substring(0, rdevSTR.length - 2) + "}";
  paramSTR = paramSTR.substring(0, paramSTR.length - 2) + "}";
  var strs = [timeFrameSTR, rdevSTR, paramSTR];
  return strs;
}

function doubleSplit(str, outer, inner) {
  var singleSplit = str.split(outer);
  var doubleSplit = [];
  for (var i = 0; i < singleSplit.length; i++) {
    doubleSplit[i] = singleSplit[i].split(inner);
  }
  return doubleSplit;
}