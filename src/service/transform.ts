import Excel from "exceljs";
import moment from "moment";
import * as _ from "lodash";

export const CELL_VALUE_TYPE = {
  IMAGE: "image",
  DATE: "date",
  STRING: "string",
  SHEET: "sheet",
  SHEET_CAPTURE: "sheet-capture",
  VALUE: "value"
};

function hasSheet(workbook: Excel.Workbook, sheetName: string) {
  return workbook && workbook.getWorksheet(sheetName);
}

function getCellSheet(mp: string): string {
  return mp.split(":").length === 2 ? mp.split(":")[0] : mp;
}

function getCellPosition(mp: string): string {
  return mp.split(":").length === 2 ? mp.split(":")[1] : mp;
}

export function getUsingSheet(
  mappings: Array<object> = [],
  cellKey: string
): Array<string> {
  const metaCells = _.flatten(
    mappings.map(mp =>
      _.get(mp, "values", []).map((value: object) => _.get(value, cellKey))
    )
  );
  return Array.from(new Set(metaCells.map(getCellSheet)));
}

interface Mapping {
  valuse: MappingValue[];
  templateName: string;
}

interface MappingValue {
  from: string;
  to: string;
  type: string;
  range?: object | undefined;
}

export function fillData(
  from: Excel.Workbook,
  to: Excel.Workbook,
  mapping: Mapping
) {
  const values = _.get(mapping, "values");
  _.forEach(values, value => {
    const type = _.get(value, "type");
    if (type === CELL_VALUE_TYPE.SHEET_CAPTURE) {
      const fromSheetNames: string[] = _.map(
        _.get(from, "worksheets"),
        sheet => sheet.name
      );
      const matchedSheets = _.filter(
        fromSheetNames,
        name => name && name.match(value.from)
      );
      _.each(matchedSheets, sheetName => {
        sheetName = `${sheetName}`;
        const fromSheet = from.getWorksheet(sheetName);
        const toSheetName = _.get(sheetName.match(new RegExp(value.from)), 1, "");
        let toSheet = to.getWorksheet(toSheetName);
        if (!toSheet) {
          toSheet = to.addWorksheet(toSheetName);
        }
        toSheet.model = fromSheet.model;
        toSheet.name = toSheetName;
        _.each(fromSheet.getImages(), image => {
          const fromImageId = Number(_.get(image, "imageId"));
          const imageId = to.addImage({
            buffer: from.getImage(fromImageId).buffer,
            extension: "png"
          });
          toSheet.addImage(imageId, {
            tl: {
              col: Number(_.get(image, "range.tl.col")),
              row: Number(_.get(image, "range.tl.row"))
            },
            br: {
              col: Number(_.get(image, "range.br.col")),
              row: Number(_.get(image, "range.br.row"))
            }
          });
        });
      });
      return true;
    }
    const fromSheet = from.getWorksheet(getCellSheet(_.get(value, "from")));
    if (fromSheet) {
      const toSheetName = getCellSheet(_.get(value, "to"));
      let toSheet = to.getWorksheet(toSheetName);
      if (!toSheet) {
        toSheet = to.addWorksheet(toSheetName);
      }
      if (type === CELL_VALUE_TYPE.SHEET) {
        toSheet.model = fromSheet.model;
        toSheet.name = toSheetName;
        _.each(fromSheet.getImages(), image => {
          const fromImageId = Number(_.get(image, "imageId"));
          const imageId = to.addImage({
            buffer: from.getImage(fromImageId).buffer,
            extension: "png"
          });
          toSheet.addImage(imageId, {
            tl: {
              col: Number(_.get(image, "range.tl.col")),
              row: Number(_.get(image, "range.tl.row"))
            },
            br: {
              col: Number(_.get(image, "range.br.col")),
              row: Number(_.get(image, "range.br.row"))
            }
          });
        });
      } else if (type === CELL_VALUE_TYPE.IMAGE) {
        const fromImageId = Number(
          _.get(_.first(fromSheet.getImages()), "imageId")
        );
        const imageId = to.addImage({
          buffer: from.getImage(fromImageId).buffer,
          extension: "png"
        });
        toSheet.addImage(imageId, _.get(value, "range"));
      } else if (type === CELL_VALUE_TYPE.DATE) {
        const cellValue = fromSheet.getCell(
          getCellPosition(_.get(value, "from"))
        ).value;
        toSheet.getCell(getCellPosition(_.get(value, "to"))).value = moment(
          _.get(cellValue, "result")
        ).format("YYYY-MM-DD");
      } else {
        toSheet.getCell(
          getCellPosition(_.get(value, "to"))
        ).value = fromSheet.getCell(
          getCellPosition(_.get(value, "from"))
        ).value;
      }
    }
  });
  return to;
}
