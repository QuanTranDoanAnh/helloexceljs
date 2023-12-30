const Excel = require("exceljs");

const fileName = "merged.xlsx";

const wb = new Excel.Workbook();
const ws = wb.addWorksheet("My Sheet");

ws.getCell("A1").value = "old falcon";
ws.getCell("A1").style.alignment = { horizontal: "center", vertical: "middle" };

ws.mergeCells("A1:C4");

wb.xlsx
  .writeFile(fileName)
  .then(() => {
    console.log("file created");
  })
  .catch((err) => {
    console.log(err.message);
  });
