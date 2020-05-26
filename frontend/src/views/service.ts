import { Workbook } from "exceljs";
export const SHEET_NAME = {
    196000: "196000",
}

function hasSheet(workbook: Workbook, sheetName: string) {
  return workbook && workbook.getWorksheet(sheetName);
}

export function validateTB(workbook: Workbook) {
  if (!hasSheet(workbook, SHEET_NAME[196000])) {
    return `tb sheet ${SHEET_NAME[196000]} is not found`;
  }
  return ``;
}

export function parse(workbook: Workbook, sheetName: string){
    return {};
}

export function loadTemplate(name: string){
    return ''
}

export async function fillData(workbook: Workbook, data: object){
    return null;
}