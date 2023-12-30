const Excel = require("exceljs");

const fileName = "simple.xlsx";

const wb = new Excel.Workbook();
const ws = wb.addWorksheet("My Sheet");

ws.getCell("A1").value = "John Doe";
ws.getCell("B1").value = "gardener";
ws.getCell("C1").value = new Date().toLocaleString();

const r3 = ws.getRow(3);
r3.values = [1, 2, 3, 4, 5, 6];

wb.xlsx
  .writeFile(fileName)
  .then(() => {
    console.log("file created");
  })
  .catch((err) => {
    console.log(err.message);
  });
