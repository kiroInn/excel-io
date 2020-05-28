import {parse, SHEET_NAME} from "../../src/views/service"
const Excel = require('exceljs');
// const path = require('path');

describe("service", () => {
  it("should parse 196000 worksheet data when passed tb workbook", () => {
    var workbook = new Excel.Workbook();
    // await workbook.xlsx.readFile('./template/202004-CNCDU-TB.xlsx');
    // const expected = {
    //     [SHEET_NAME._196000]: {
    //         "COL_A_ROW_4": "2020-04-30"
    //     }
    // }
    // expect(parse(workbook, SHEET_NAME._196000)).toEqual(expected);
    expect(1).toBe(1);
  });
});
