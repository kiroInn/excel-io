import Excel, { CellValue } from "exceljs";
import _ from "lodash";

export const EXPRESS_VALUE_NOT_FOUND = "❓";
function parseCellPosition(
  workSheet: Excel.Worksheet,
  addressStr: string
): { columnIndex: string | number; rowIndex: string | number } {
  let columnIndex: string | number = -1,
    rowIndex = -1;
  if (workSheet && addressStr) {
    columnIndex = _.get(addressStr.split("$"), 0);
    const express = _.get(/\${(.*?)}/.exec(addressStr), 1);
    const expressPrefix = _.get(express?.split("="), 0);
    const expressValues = _.get(express?.split("="), 1)?.split("|");
    rowIndex = _.findIndex(
      workSheet.getColumn(expressPrefix || "").values,
      value => {
        if (_.isObject(value) && _.has(value, "richText")) {
          return _.includes(
            expressValues,
            _.map(_.get(value, "richText"), item => _.get(item, "text")).join(
              ""
            )
          );
        }
        return _.includes(expressValues, `${value}`);
      }
    );
  }
  return { columnIndex, rowIndex };
}

export function parse(
  files: [{ fileName: string; workbook: Excel.Workbook }],
  variableArray: []
): [][] {
  const result: any[] = [];
  _.each(variableArray, variables => {
    const rowResult: any[] = [];
    _.each(variables, (content: string, index) => {
      const fileName = _.get(content.split(":"), 0);
      const sheetName = _.get(content.split(":"), 1);
      if (index === 1) rowResult.push(fileName);
      const addressExp = _.get(content.split(":"), 2);
      const isNegative = _.get(content.split(":"), 3) === "isNegative";
      const workbook: Excel.Workbook | undefined = _.get(
        _.find(
          files,
          ({ fileName: name }) =>
            !!name.match(new RegExp(`${fileName}\\.`, "g"))
        ),
        "workbook"
      );
      let value: CellValue | string = EXPRESS_VALUE_NOT_FOUND;
      if (workbook) {
        const workSheet: Excel.Worksheet = workbook.getWorksheet(sheetName);
        const address = parseCellPosition(workSheet, addressExp);
        if (address.rowIndex != -1) {
          const cell: Excel.Cell = workSheet.getCell(
            `${address.columnIndex}${address.rowIndex}`
          );
          value = _.get(cell, "value");
        }
        value = _.isObject(value) ? _.get(value, "result", 0) : value;
        try {
          value =
            isNegative && value !== EXPRESS_VALUE_NOT_FOUND
              ? Number.parseFloat(`-${value}`).toFixed(2)
              : value;
        } catch (err) {
          console.error(`${sheetName} ${address} expression rule is err.`);
        }
        if (sheetName === "200000") console.log("value：", value);
      }
      rowResult.push(value);
    });
    result.push(rowResult);
  });
  return result;
}
