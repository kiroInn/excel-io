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
    console.log("expressPrefix:", expressPrefix);
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
      const workbook: Excel.Workbook | undefined = _.get(
        _.find(files, ({ fileName: name }) => _.includes(name, fileName)),
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
      }
      rowResult.push(_.isObject(value) ? _.get(value, "result", 0) : value);
    });
    result.push(rowResult);
  });
  return result;
}

export const DEFAULT_EXPRESS = [
  [
    "TB_BU_FN:Sheet1:N${B=106700}",
    "106700:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "106700:Reconciliation:C${B=BOOK BALANCE}",
    "106700:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:N${B=106700}",
    "106700:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=106710}",
    "106710:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "106710:Reconciliation:C${B=BOOK BALANCE}",
    "106710:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:N${B=106700}",
    "106710:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=106750}",
    "106750:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "106750:Reconciliation:C${B=BOOK BALANCE}",
    "106750:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:N${B=106700}",
    "106750:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=112000}",
    "112000:112000:C${B=Totals}",
    "112000:112000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "112000:112000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=119000}",
    "119000:119000:C${B=Totals}",
    "119000:119000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "119000:119000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=120000}",
    "120000:120000:AX${B=Totals}",
    "120000:120000:AX${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "120000:120000:AX${B=Check}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=140000}",
    "140000:140000:C${B=Totals}",
    "140000:140000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "140000:140000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=190500}",
    "190500:190500:AC${D=Totals}",
    "190500:190500:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "190500:190500:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=191000}",
    "191000:191000:C${B=Totals}",
    "191000:191000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "191000:191000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=196000}",
    "196000:196000:C${B=Totals}",
    "196000:196000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "196000:196000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=200000}",
    "200000:200000:C${B=Balance per subledger}",
    "200000:200000:C${B=Balance per GL}",
    "200000:200000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=205000}",
    "205000:205000:AC${D=Totals}",
    "205000:205000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "205000:205000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=210000}",
    "210000:210000:AC${D=Totals}",
    "210000:210000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "210000:210000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=214000}",
    "214000:214000:AC${D=Totals}",
    "214000:214000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "214000:214000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=215000}",
    "215000:215000:AC${D=Totals}",
    "215000:215000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "215000:215000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=215100}",
    "215100:215100:AC${D=Totals}",
    "215100:215100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "215100:215100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=215300}",
    "215300:215300:AC${D=Totals}",
    "215300:215300:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "215300:215300:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=215600}",
    "215600:215600:AC${D=Totals}",
    "215600:215600:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "215600:215600:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=216000}",
    "216000:216000:AC${D=Totals}",
    "216000:216000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "216000:216000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=217500}",
    "217500:217500:AC${D=Totals}",
    "217500:217500:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "217500:217500:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=221850}",
    "221850:221850:AC${D=Totals}",
    "221850:221850:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "221850:221850:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=230000}",
    "230000:230000:AC${D=Totals}",
    "230000:230000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "230000:230000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=235000}",
    "235000:235000:AC${D=Totals}",
    "235000:235000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "235000:235000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=235100}",
    "235100:235100:AC${D=Totals}",
    "235100:235100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "235100:235100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=274000}",
    "274000:274000:AC${D=Totals}",
    "274000:274000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "274000:274000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=897100}",
    "897100:897100:AC${D=Totals}",
    "897100:897100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "897100:897100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=897600}",
    "897600:897600:AC${D=Totals}",
    "897600:897600:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "897600:897600:AC${D=Variance}"
  ]
];
