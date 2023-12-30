const Excel = require("exceljs");
const wb = new Excel.Workbook();
const fileName = "myexcel.xlsx";

wb.xlsx
  .readFile(fileName)
  .then(() => {
    const ws = wb.getWorksheet("My Sheet");
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
