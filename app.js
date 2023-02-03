const Excel = require('exceljs');
const _ = require('lodash');
const fs = require("fs");

(async function () {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile("sac.xlsx");

    const sac = {};

    workbook.definedNames.forEach((name, cell) => {
        if (cell.sheetName === "General Information") {
            const value = workbook.getWorksheet(cell.sheetName).getCell(cell.row, cell.col).value;
            _.set(sac, name, value);
            return;
        }

        if (cell.sheetName === "Federal Awards") {
            const value = workbook.getWorksheet(cell.sheetName).getCell(cell.row, cell.col).value;
            _.set(sac, `FederalAwards[${cell.row - 2}].${name}`, value)
        }
    });

    fs.writeFileSync("sac.json", JSON.stringify(sac, null, 2));
})();