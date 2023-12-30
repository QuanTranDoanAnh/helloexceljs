const Excel = require("exceljs");

const fileName = "formula.xlsx";

const wb = new Excel.Workbook();
const ws = wb.addWorksheet("My Sheet");

ws.getCell("A1").value = 1;
ws.getCell("A2").value = 2;
ws.getCell("A3").value = 3;
ws.getCell("A4").value = 4;
ws.getCell("A5").value = 5;
ws.getCell("A6").value = 6;

let a7 = ws.getCell("A7");
a7.value = { formula: "SUM(A1:A6)" };
a7.style.font = { bold: true };

writeFile(wb);

async function writeFile(wb) {
  await wb.xlsx.writeFile(fileName);
}
