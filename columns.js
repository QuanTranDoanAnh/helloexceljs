const Excel = require("exceljs");

const wb = new Excel.Workbook();
const ws = wb.addWorksheet("My Sheet");

const headers = [
  {
    header: "First name",
    key: "fn",
    width: 15,
  },
  {
    header: "Last name",
    key: "ln",
    width: 15,
  },
  {
    header: "Occupation",
    key: "occ",
    width: 15,
  },
  {
    header: "Salary",
    key: "sl",
    width: 15,
  },
];

ws.columns = headers;

ws.addRow(["Mukul", "Latiyan", "Software Developer", 1230]);
ws.addRow(["Prince", "Yadav", "Driver", 980]);
ws.addRow(["Mayank", "Agarwal", "Maali", 770]);

ws.getColumn("fn").eachCell((cell, rn) => {
  console.log(cell.value);
});

console.log("--------------");

ws.getColumn("B").eachCell((cell, rn) => {
  console.log(cell.value);
});

console.log("--------------");

ws.getColumn(3).eachCell((cell, rn) => {
  console.log(cell.value);
});

console.log("--------------");
console.log(`There are ${ws.actualColumnCount} columns`);
