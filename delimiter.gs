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
        var length = cell.getValue().split(", ").length;
        var toAdd = "{{" + (n0.getValue() + "; ").repeat(length - 1) + n0.getValue() + "}, {" + (T0.getValue() + "; ").repeat(length - 1) + T0.getValue() + "}, {" + frames(length) + "}, TRANSPOSE(SPLIT(Delimited!" + cell.getA1Notation() + ", \", \"))};";
        str += toAdd;
      }
    }
  }
  str = str.substring(0, str.length - 1);
  str += "}, 3, TRUE)";
  return str;
}

function frames(num) {
  var p = "";
  for (var i = 1; i < num; i++) {
    p += i.toString() + "; "
  }
  p += num.toString();
  return p;
}