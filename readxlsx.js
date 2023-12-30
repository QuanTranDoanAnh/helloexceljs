// Read Simple Cell Values

const Excel = require("exceljs");
const wb = new Excel.Workbook();
const ws = wb.addWorksheet("My Sheet");

ws.addRows([
  [10, 2, 3, 4, 5],
  [6, 11, 8, 9, 10],
  [10, 11, 12, 14, 15],
  [16, 17, 18, 13, 20],
]);

const valueOne = ws.getCell("B1").value;
console.log(valueOne);

const valueTwo = ws.getCell(1, 1).value;
console.log(valueTwo);

const valueThree = ws.getRow(3).getCell(3).value;
console.log(valueThree);
