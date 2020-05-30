import Excel from "exceljs";
export const SHEET_NAME = {
  _196000: "196000"
};

const CELL_VALUE_TYPE = {
  IMAGE: "image",
  VALUE: "value",
}

const CONVERSION_RELATIONS = {
  [SHEET_NAME._196000]: {
    "196000" : [
      {from: "C6", to: "C8"}
     ],
    "TB" : [
       {from: "A3", to: "A3"},
       {from: "C6", to: "C6"},
       {from: "C6", to: "C7"},
       {range: {
         tl: { col: 0, row: 10.5 },
         br: { col: 6.5, row: 15.5 }
       }, type: CELL_VALUE_TYPE.IMAGE},

     ]
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

export function loadTemplate(sheetName: string){
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

export function fillData(from: Excel.Workbook, to: Excel.Workbook, sheetName: string) {
  const resultMeta = CONVERSION_RELATIONS[sheetName];
  const fromWorkSheet = from.getWorksheet(sheetName);
  Object.keys(resultMeta).forEach(rsn => {
    console.log(rsn);
    const sheetMetaArray = resultMeta[rsn];
    const toWorkSheet = to.getWorksheet(rsn);
    sheetMetaArray.forEach((meta:any) => {
      if(meta.type === CELL_VALUE_TYPE.IMAGE){
        const fromImageId = Number(fromWorkSheet.getImages()[0].imageId);
        var imageId = to.addImage({
          buffer:  from.getImage(fromImageId).buffer,
          extension: 'png',
        });
        toWorkSheet.addImage(imageId, meta.range)
      }else{
        toWorkSheet.getCell(meta.to).value = fromWorkSheet.getCell(meta.from).value
      }
    })
  })
  return to;
}
