const ExcelJS = require("exceljs");

const wb = new ExcelJS.Workbook();

const fileName = "items.xlsx";

wb.xlsx
  .readFile(fileName)
  .then(() => {
    const ws = wb.getWorksheet("Sheet1");

    const c1 = ws.getColumn(1);

    c1.eachCell((c) => {
      console.log(c.value);
    });

    const c2 = ws.getColumn(2);

    c2.eachCell((c) => {
      console.log(c.value);
    });
  })
  .catch((err) => {
    console.log(err.message);
  });
