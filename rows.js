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

console.log(`There are ${ws.actualRowCount} rows`);

let rows = ws.getRows(1, 4).values();

for (let row of rows) {
  row.eachCell((cell, cn) => {
    console.log(cell.value);
  });
  console.log("--");
}
