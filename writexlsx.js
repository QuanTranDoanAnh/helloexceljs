const Excel = require("exceljs");
const fileName = "myexcel.xlsx";

const wb = new Excel.Workbook();
const ws = wb.addWorksheet("My Sheet");

ws.getCell("A1").value = "Mukul Latiyan";
ws.getCell("B1").value = "Software Developer";
ws.getCell("C1").value = new Date().toLocaleString();

const r3 = ws.getRow(3);
r3.values = [1, 2, 3, 4];

wb.xlsx
  .writeFile(fileName)
  .then(() => {
    console.log("file created");
  })
  .catch((err) => {
    console.log(err.message);
  });
