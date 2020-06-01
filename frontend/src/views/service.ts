import Excel from "exceljs";
import moment from "moment";

export const SHEET_NAME = {
  _196000: "196000",
  _106700: "106700",
};

const CELL_VALUE_TYPE = {
  IMAGE: "image",
  DATE: "date",
  VALUE: "value"
};

const CONVERSION_RELATIONS = {
  [SHEET_NAME._106700]: {
    outputName: "202004-CNCDU-106700.xlsx",
    relation: {
      "Reconciliation": [
        {from: "A3", to: "A4", type: CELL_VALUE_TYPE.DATE},
        {from: "C6", to: "C10"},
        {
          range: {
            tl: { col: 4.5, row: 10.5 },
            br: { col: 11, row: 26.5 }
          },
          type: CELL_VALUE_TYPE.IMAGE
        }
      ]
    }
  },
  [SHEET_NAME._196000]: {
    outputName: "202004-CNCDU-196000.xlsx",
    relation: {
      "196000": [
        { from: "C6", to: "C8" },
        { from: "A3", to: "A4", type: CELL_VALUE_TYPE.DATE }
      ],
      TB: [
        { from: "C6", to: "C6" },
        { from: "C6", to: "C7" },
        {
          range: {
            tl: { col: 0, row: 10.5 },
            br: { col: 6.5, row: 15.5 }
          },
          type: CELL_VALUE_TYPE.IMAGE
        }
      ]
    }
  }
};

function hasSheet(workbook: Excel.Workbook, sheetName: string) {
  return workbook && workbook.getWorksheet(sheetName);
}

export function validateTB(workbook: Excel.Workbook) {
  const missingSheet: any = [];
  Object.values(SHEET_NAME).forEach(sheetName => {
      if(!hasSheet(workbook, sheetName)) {
        missingSheet.push(sheetName);
      }
  })
  return missingSheet.length > 0 ?  `tb sheet ${missingSheet.join(",")} is not found` : '';
}

export function loadTemplate(sheetName: string) {
  return new Promise(resolve => {
    fetch(`http://localhost:3000/static/${CONVERSION_RELATIONS[sheetName].outputName}`)
      .then(async response => {
        response.arrayBuffer().then(async buffer => {
          const workbook = new Excel.Workbook();
          await workbook.xlsx.load(buffer);
          resolve(workbook);
        });
      })
      .catch(err => console.log("err", err));
  });
}

export function fillData(
  from: Excel.Workbook,
  to: Excel.Workbook,
  sheetName: string
) {
  const conversionRelation = CONVERSION_RELATIONS[sheetName].relation;
  const fromWorkSheet = from.getWorksheet(sheetName);
  Object.keys(conversionRelation).forEach(rsn => {
    const sheetMetaArray = conversionRelation[rsn];
    const toWorkSheet = to.getWorksheet(rsn);
    sheetMetaArray.forEach((meta: any) => {
      if (meta.type === CELL_VALUE_TYPE.IMAGE) {
        const fromImageId = Number(fromWorkSheet.getImages()[0].imageId);
        const imageId = to.addImage({
          buffer: from.getImage(fromImageId).buffer,
          extension: "png"
        });
        toWorkSheet.addImage(imageId, meta.range);
      } else if (meta.type === CELL_VALUE_TYPE.DATE) {
        const cellValue: any = fromWorkSheet.getCell(meta.from).value;
        toWorkSheet.getCell(meta.to).value = moment(cellValue.result).format(
          "YYYY-MM-DD"
        );
      } else {
        toWorkSheet.getCell(meta.to).value = fromWorkSheet.getCell(
          meta.from
        ).value;
      }
    });
  });
  return to;
}
