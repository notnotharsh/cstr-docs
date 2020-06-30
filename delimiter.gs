var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1m00wLm_LkEVhOqWPq70uFWheFD3OO7z1sdEHvI0nOvk/edit");

function getSheet(name) {
  return spreadsheet.getSheetByName(name);
}

function createFormula() {
  var range = getSheet("Delimited").getRange("A1:V22");
  var str = forEachRangeCell(range);
  getSheet("logs").getRange("A2").setValue(str);
}

function forEachRangeCell(range) {
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  str = "=SORT({"
  for (let i = 2; i <= numCols; i++) {
    for (let j = 2; j <= numRows; j++) {
      const cell = range.getCell(j, i);
      if (cell.getValue() != "") {
        const n0 = range.getCell(1, i);
        const T0 = range.getCell(j, 1);
        var packets = cell.getValue().split(";");
        var length = packets[1].split(",").length;
        var toAdd = "{" + n0.getValue() + " * Reference!A1:A" + length + "^0, " + T0.getValue() + " * Reference!A1:A" + length + "^0, Reference!A1:A" + length + ", " + packets[0] + " * Reference!A1:A" + length + "^0, TRANSPOSE(SPLIT(INDEX(SPLIT(Delimited!" + cell.getA1Notation() + ", \";\"), 2), \",\")) / 100};";
        str += toAdd;
      }
    }
  }
  str = str.substring(0, str.length - 1);
  str += "}, 3, TRUE)";
  return str;
}