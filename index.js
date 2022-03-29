"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const XLSX = __importStar(require("xlsx"));
const range = 10000;
let workbook = XLSX.readFile("./excel_files/input/Copy of WorldPay Card Bin List (as at 14 March 2022).xlsx");
const wsname = workbook.SheetNames[0];
const ws = workbook.Sheets[wsname];
const exceldata = XLSX.utils.sheet_to_json(ws, {
    header: ["RANGEFROM", "RANGEUNTIL", "COUNTRY", "BRAND", "ISSUER", "FAMILY"],
});
exceldata.shift();
//console.log(exceldata);
let index = 0;
const dataLength = exceldata.length;
while (index * range - 1 <= dataLength - 1) {
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
