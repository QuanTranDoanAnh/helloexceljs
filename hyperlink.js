const Excel = require("exceljs");

const fileName = "hyperlink.xlsx";

const wb = new Excel.Workbook();
const ws = wb.addWorksheet("My Sheet");

ws.getCell("A1").value = {
  hyperlink: "http://webcode.me",
  text: "WebCode",
  tooltip: "http://webcode.me",
};

ws.getCell("A1").font = { underline: true, color: "blue" };

wb.xlsx
  .writeFile(fileName)
  .then(() => {
    console.log("Done.");
  })
  .catch((err) => {
    console.log(err.message);
  });
