import Excel from "exceljs";
export const SHEET_NAME = {
  _196000: "196000"
};

const SHEET_META_INDEX = {
  COLUM : 0,
  ROW: 1,
}

const SHEET_META = {
  [SHEET_NAME._196000]: {
   TARGET: [["C","6"]],
   RESULT: {
     "196000" : [{}],
     "TB" : []
   }
  }
}

const templateFile = {
  [SHEET_NAME._196000] : "202004-CNCDU-196000.xlsx",
} 

function hasSheet(workbook: Excel.Workbook, sheetName: string) {
  return workbook && workbook.getWorksheet(sheetName);
}

export function validateTB(workbook: Excel.Workbook) {
  if (!hasSheet(workbook, SHEET_NAME._196000)) {
    return `tb sheet ${SHEET_NAME._196000} is not found`;
  }
  return ``;
}

export function parse(workbook: Excel.Workbook, sheetName: string) {
  const worksheet = workbook.getWorksheet(sheetName);
  if (worksheet) {
    console.log('workbook', workbook);
    return SHEET_META[sheetName].TARGET.map(colAndRow => {
      const colum = worksheet.getColumn(colAndRow[SHEET_META_INDEX.COLUM]);
      if(!colum || !colum.values) throw Error(`Parse ${sheetName} err`)
      return colum.values[Number(colAndRow[SHEET_META_INDEX.ROW])];
    })
  }
  return [];
}

export async function loadTemplate(sheetName: string) {
  return new Promise((resolve) => {
    fetch(`http://localhost:3000/static/${templateFile[sheetName]}`).then(async (response) => {
    response.arrayBuffer().then(async buffer=> {
      const workbook = new Excel.Workbook();
      await workbook.xlsx.load(buffer);
      resolve(workbook);
    })
   }).catch(err => console.log('err', err))
  })
}

export async function fillData(workbook: Workbook, data: object) {
  return null;
}
