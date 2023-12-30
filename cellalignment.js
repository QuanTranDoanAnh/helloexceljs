const Excel = require("exceljs");

const fileName = "align.xlsx";

const wb = new Excel.Workbook();
const ws = wb.addWorksheet("My Sheet");

const headers = [
  { header: "First name", key: "fn", width: 15 },
  { header: "Last name", key: "ln", width: 15 },
  { header: "Occupation", key: "occ", width: 15 },
  { header: "Salary", key: "sl", width: 15 },
];

ws.columns = headers;

ws.addRow(["John", "Doe", "gardener", 1230]);
ws.addRow(["Roger", "Roe", "driver", 980]);
ws.addRow(["Lucy", "Mallory", "teacher", 780]);
ws.addRow(["Peter", "Smith", "programmer", 2300]);

ws.getColumn("A").alignment = { vertical: "middle", horizontal: "left" };
ws.getColumn("B").alignment = { vertical: "middle", horizontal: "left", indent: 1 };
ws.getColumn("C").alignment = { vertical: "middle", horizontal: "left", wrapText: true };
ws.getColumn("D").alignment = { vertical: "middle", horizontal: "right" };

ws.getRow(1).alignment = { vertical: "middle", horizontal: "center" };

wb.xlsx
  .writeFile(fileName)
  .then(() => {
    console.log("file created");
  })
  .catch((err) => {
    console.log(err.message);
  });
