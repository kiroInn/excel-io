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

export const DEFAULT_EXPRESS = [
  [
    "TB_BU_FN:Sheet1:N${B=106700}",
    "CNCDU-106700:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "CNCDU-106700:Reconciliation:C${B=BOOK BALANCE}",
    "CNCDU-106700:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:N${B=106700}",
    "CNCDU-106700:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=106710}",
    "CNCDU-106710:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "CNCDU-106710:Reconciliation:C${B=BOOK BALANCE}",
    "CNCDU-106710:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:N${B=106710}",
    "CNCDU-106710:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=106750}",
    "CNCDU-106750:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "CNCDU-106750:Reconciliation:C${B=BOOK BALANCE}",
    "CNCDU-106750:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:N${B=106700}",
    "CNCDU-106750:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=112000}",
    "CNCDU-112000:112000:C${B=Totals}",
    "CNCDU-112000:112000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-112000:112000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=119000}",
    "CNCDU-119000:119000:C${B=Totals}",
    "CNCDU-119000:119000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-119000:119000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=120000}",
    "CNCDU-120000:120000:AX${B=Totals}",
    "CNCDU-120000:120000:AX${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-120000:120000:AX${B=Check}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=140000}",
    "CNCDU-140000:140000:C${B=Totals}",
    "CNCDU-140000:140000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-140000:140000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=190500}",
    "CNCDU-190500:190500:AC${D=Totals}",
    "CNCDU-190500:190500:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-190500:190500:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=191000}",
    "CNCDU-191000:191000:C${B=Totals}",
    "CNCDU-191000:191000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-191000:191000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=196000}",
    "CNCDU-196000:196000:C${B=Totals}",
    "CNCDU-196000:196000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-196000:196000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=200000}",
    "CNCDU-200000:200000:C${B=Balance per subledger}:isNegative",
    "CNCDU-200000:200000:C${B=Balance per GL}",
    "CNCDU-200000:200000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=205000}",
    "CNCDU-205000:205000:AC${D=Totals}",
    "CNCDU-205000:205000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-205000:205000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=210000}",
    "CNCDU-210000:210000:AC${D=Totals}",
    "CNCDU-210000:210000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-210000:210000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=214000}",
    "CNCDU-214000:214000:AC${D=Totals}",
    "CNCDU-214000:214000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-214000:214000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=215000}",
    "CNCDU-215000:215000:AC${D=Totals}",
    "CNCDU-215000:215000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-215000:215000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=215100}",
    "CNCDU-215100:215100:AC${D=Totals}",
    "CNCDU-215100:215100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-215100:215100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=215300}",
    "CNCDU-215300:215300:AC${D=Totals}",
    "CNCDU-215300:215300:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-215300:215300:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=215600}",
    "CNCDU-215600:215600:AC${D=Totals}",
    "CNCDU-215600:215600:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-215600:215600:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=216000}",
    "CNCDU-216000:216000:AC${D=Totals}",
    "CNCDU-216000:216000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-216000:216000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=217500}",
    "CNCDU-217500:217500:AC${D=Totals}",
    "CNCDU-217500:217500:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-217500:217500:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=221850}",
    "CNCDU-221850:221850:AC${D=Totals}",
    "CNCDU-221850:221850:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-221850:221850:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=230000}",
    "CNCDU-230000:230000:AC${D=Totals}",
    "CNCDU-230000:230000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-230000:230000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=235000}",
    "CNCDU-235000:235000:AC${D=Totals}",
    "CNCDU-235000:235000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-235000:235000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=235100}",
    "CNCDU-235100:235100:AC${D=Totals}",
    "CNCDU-235100:235100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-235100:235100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=274000}",
    "CNCDU-274000:274000:AC${D=Totals}",
    "CNCDU-274000:274000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-274000:274000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=897100}",
    "CNCDU-897100:897100:AC${D=Totals}",
    "CNCDU-897100:897100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-897100:897100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:N${B=897600}",
    "CNCDU-897600:897600:AC${D=Totals}",
    "CNCDU-897600:897600:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNCDU-897600:897600:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=106400}",
    "CNBSU-106400:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "CNBSU-106400:Reconciliation:C${B=BOOK BALANCE}",
    "CNBSU-106400:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:O${B=106400}",
    "CNBSU-106400:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=106410}",
    "CNBSU-106410:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "CNBSU-106410:Reconciliation:C${B=BOOK BALANCE}",
    "CNBSU-106410:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:O${B=106410}",
    "CNBSU-106410:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=106420}",
    "CNBSU-106420:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "CNBSU-106420:Reconciliation:C${B=BOOK BALANCE}",
    "CNBSU-106420:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:O${B=106420}",
    "CNBSU-106420:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=106600}",
    "CNBSU-106600:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "CNBSU-106600:Reconciliation:C${B=BOOK BALANCE}",
    "CNBSU-106600:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:O${B=106600}",
    "CNBSU-106600:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=112000}",
    "CNBSU-112000:112000:C${B=Totals}",
    "CNBSU-112000:112000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-112000:112000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=119000}",
    "CNBSU-119000:119000:C${B=Totals}",
    "CNBSU-119000:119000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-119000:119000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=120000}",
    "CNBSU-120000:120000:AX${B=Totals}",
    "CNBSU-120000:120000:AX${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-120000:120000:AX${B=Check}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=190500}",
    "CNBSU-190500:190500:AC${D=Totals}",
    "CNBSU-190500:190500:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-190500:190500:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=191000}",
    "CNBSU-191000:191000:C${B=Totals}",
    "CNBSU-191000:191000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-191000:191000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=196000}",
    "CNBSU-196000:196000:C${B=Totals}",
    "CNBSU-196000:196000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-196000:196000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=200000}",
    "CNBSU-200000:200000:C${B=Balance per subledger}:isNegative",
    "CNBSU-200000:200000:C${B=Balance per GL}",
    "CNBSU-200000:200000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=205000}",
    "CNBSU-205000:205000:AC${D=Totals}",
    "CNBSU-205000:205000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-205000:205000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=210000}",
    "CNBSU-210000:210000:AC${D=Totals}",
    "CNBSU-210000:210000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-210000:210000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=214000}",
    "CNBSU-214000:214000:AC${D=Totals}",
    "CNBSU-214000:214000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-214000:214000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=215000}",
    "CNBSU-215000:215000:AC${D=Totals}",
    "CNBSU-215000:215000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-215000:215000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=215100}",
    "CNBSU-215100:215100:AC${D=Totals}",
    "CNBSU-215100:215100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-215100:215100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=215300}",
    "CNBSU-215300:215300:AC${D=Totals}",
    "CNBSU-215300:215300:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-215300:215300:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=215600}",
    "CNBSU-215600:215600:AC${D=Totals}",
    "CNBSU-215600:215600:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-215600:215600:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=216000}",
    "CNBSU-216000:216000:AC${D=Totals}",
    "CNBSU-216000:216000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-216000:216000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=217500}",
    "CNBSU-217500:217500:AC${D=Totals}",
    "CNBSU-217500:217500:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-217500:217500:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=221850}",
    "CNBSU-221850:221850:AC${D=Totals}",
    "CNBSU-221850:221850:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-221850:221850:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=230000}",
    "CNBSU-230000:230000:AC${D=Totals}",
    "CNBSU-230000:230000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-230000:230000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=233000}",
    "CNBSU-233000:233000:AC${D=Totals}",
    "CNBSU-233000:233000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-233000:233000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=235000}",
    "CNBSU-235000:235000:AC${D=Totals}",
    "CNBSU-235000:235000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-235000:235000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=235100}",
    "CNBSU-235100:235100:AC${D=Totals}",
    "CNBSU-235100:235100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-235100:235100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=274000}",
    "CNBSU-274000:274000:AC${D=Totals}",
    "CNBSU-274000:274000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-274000:274000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=897100}",
    "CNBSU-897100:897100:AC${D=Totals}",
    "CNBSU-897100:897100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-897100:897100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:O${B=897600}",
    "CNBSU-897600:897600:AC${D=Totals}",
    "CNBSU-897600:897600:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNBSU-897600:897600:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=106430}",
    "CNSZU-106430:Reconciliation:C${B=BANK BALANCE (AS PER BANK STATEMENT)}",
    "CNSZU-106430:Reconciliation:C${B=BOOK BALANCE}",
    "CNSZU-106430:Reconciliation:C${B=VARIANCE}",
    "TB_BU:Sheet1:P${B=106430}",
    "CNSZU-106430:Reconciliation:C${B=BALANCE PER THE GL (FUNCTIONAL CURRENCY)}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=112000}",
    "CNSZU-112000:112000:C${B=Totals}",
    "CNSZU-112000:112000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-112000:112000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=120000}",
    "CNSZU-120000:120000:AX${G=Totals}",
    "CNSZU-120000:120000:AX${G=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-120000:120000:AX${G=Check}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=140000}",
    "CNSZU-140000:140000:C${B=Totals}",
    "CNSZU-140000:140000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-140000:140000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=191000}",
    "CNSZU-191000:191000:C${B=Totals}",
    "CNSZU-191000:191000:C${B=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-191000:191000:C${B=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=214000}",
    "CNSZU-214000:214000:AC${D=Totals}",
    "CNSZU-214000:214000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-214000:214000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=215000}",
    "CNSZU-215000:215000:AC${D=Totals}",
    "CNSZU-215000:215000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-215000:215000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=215100}",
    "CNSZU-215100:215100:AC${D=Totals}",
    "CNSZU-215100:215100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-215100:215100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=215300}",
    "CNSZU-215300:215300:AC${D=Totals}",
    "CNSZU-215300:215300:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-215300:215300:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=215600}",
    "CNSZU-215600:215600:AC${D=Totals}",
    "CNSZU-215600:215600:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-215600:215600:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=216000}",
    "CNSZU-216000:216000:AC${D=Totals}",
    "CNSZU-216000:216000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-216000:216000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=217500}",
    "CNSZU-217500:217500:AC${D=Totals}",
    "CNSZU-217500:217500:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-217500:217500:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=230000}",
    "CNSZU-230000:230000:AC${D=Totals}",
    "CNSZU-230000:230000:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-230000:230000:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=235100}",
    "CNSZU-235100:235100:AC${D=Totals}",
    "CNSZU-235100:235100:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-235100:235100:AC${D=Variance}"
  ],
  [
    "TB_BU_FN:Sheet1:P${B=897600}",
    "CNSZU-897600:897600:AC${D=Totals}",
    "CNSZU-897600:897600:AC${D=CNCDU|CNXAU|CNBSU|CNWHU|CNSZU}",
    "CNSZU-897600:897600:AC${D=Variance}"
  ]
];
