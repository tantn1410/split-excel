import * as XLSX from "xlsx";

const range = 10000;
let workbook = XLSX.readFile(
  "./excel_files/input/Copy of WorldPay Card Bin List (as at 14 March 2022).xlsx"
);
const wsname = workbook.SheetNames[0];
const ws = workbook.Sheets[wsname];
const exceldata = XLSX.utils.sheet_to_json(ws, {
  header: ["RANGEFROM", "RANGEUNTIL", "COUNTRY", "BRAND", "ISSUER", "FAMILY"],
});

exceldata.shift();
//console.log(exceldata);

let index = 0;
const dataLength = exceldata.length;
while (index * range - 1 <= dataLength -1) {
    const subData = exceldata.slice(index * range, Math.min(dataLength, (index + 1) * range));
    //console.log(">>>" + index * range + ", " + Math.min(dataLength - 1, (index + 1) * range -1));
    const newWorkSheet = XLSX.utils.json_to_sheet(subData, {
        header: ["RANGEFROM", "RANGEUNTIL", "COUNTRY", "BRAND", "ISSUER", "FAMILY"],
        skipHeader: false,
      });
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorkSheet, "CARDBIN", true);
    XLSX.writeFile(newWorkbook, `./excel_files/output/Result${index + 1}.xlsx`);
    index++;
}
