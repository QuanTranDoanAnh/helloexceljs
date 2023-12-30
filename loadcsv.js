const Excel = require("exceljs");

const fileName = "cars.xlsx";
const data = "cars.csv";

const wb = new Excel.Workbook();

wb.csv
  .readFile(data)
  .then((ws) => {
    console.log(
      `Sheet ${ws.id} - ${ws.name}, Dims=${JSON.stringify(ws.dimensions)}`
    );

    for (let i = 1; i <= ws.actualRowCount; i++) {
      for (let j = 1; j <= ws.actualColumnCount; j++) {
        const val = ws.getRow(i).getCell(j);
        process.stdout.write(`${val} `);
      }
      console.log();
    }
  })
  .then(() => {
    writeData();
  });

function writeData() {
  wb.xlsx
    .writeFile(fileName)
    .then(() => {
      console.log("Done.");
    })
    .catch((err) => {
      console.log(err.message);
    });
}
