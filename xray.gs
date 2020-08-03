var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1m00wLm_LkEVhOqWPq70uFWheFD3OO7z1sdEHvI0nOvk/edit");

function getSheet(name) {
  return spreadsheet.getSheetByName(name);
}

function createFormula() {
  var range = getSheet("Delimited").getRange("A1:V22");
  var str = forEachRangeCell(range);
  getSheet("ScrapLogs").getRange("B2").setValue(str);
}

function forEachRangeCell(range) {
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  var str1 = "={({"
  var sumstr = "";
  var totalLength = 1;
  for (let i = 2; i <= numCols; i++) {
    for (let j = 2; j <= numRows; j++) {
      const cell = range.getCell(j, i);
      if (cell.getValue() != "") {
        const n0 = range.getCell(1, i);
        const T0 = range.getCell(j, 1);
        var packets = cell.getValue().split(";");
        var begin = packets[1].replace(" ", "");
        var length = packets[2].split(",").length;
        totalLength += length;
        var toAdd = "{" + "ARRAYFORMULA(Reference!A" + begin + ":A" + (parseInt(length) + parseInt(begin) - 1) + "), ARRAYFORMULA(" + T0.getValue() + " * Reference!A" + begin + ":A" + (parseInt(length) + parseInt(begin) - 1) + "^0), ARRAYFORMULA(" + n0.getValue() + " * Reference!A" + begin + ":A" + (parseInt(length) + parseInt(begin) - 1) + "^0), ARRAYFORMULA(" + packets[0] + " * Reference!A" + begin + ":A" + (parseInt(length) + parseInt(begin) - 1) + "^0), ARRAYFORMULA(TRANSPOSE(SPLIT(INDEX(SPLIT(Delimited!" + cell.getA1Notation() + ", \";\"), 3), \",\")) / 100)}; ";
        sumstr += toAdd;
      }
    }
  }
  var str = str1 + sumstr;
  str = str.substring(0, str.length - 2);
  str += "})}";
  return str;
}