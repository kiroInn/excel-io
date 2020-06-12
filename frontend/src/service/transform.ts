import Excel from "exceljs";
import moment from "moment";
import * as _ from "lodash";

export const CELL_VALUE_TYPE = {
  IMAGE: "image",
  DATE: "date",
  STRING: "string",
  SHEET: "sheet",
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

export function validateFrom(
  workbook: Excel.Workbook,
  mappings: Array<object>
) {
  const missingSheet: Array<string> = [];
  const useSheets: Array<string> = getUsingSheet(mappings, "from");
  useSheets.forEach(sheetName => {
    if (!hasSheet(workbook, sheetName)) {
      missingSheet.push(sheetName);
    }
  });
  return missingSheet.length > 0
    ? `from sheet ${missingSheet.join(",")} is not found`
    : "";
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
    const fromSheet = from.getWorksheet(getCellSheet(_.get(value, "from")));
    let toSheet = to.getWorksheet(getCellSheet(_.get(value, "to")));
    if (!toSheet) {
      toSheet = to.addWorksheet(getCellSheet(_.get(value, "to")));
    }
    const type = _.get(value, "type");
    if (type === CELL_VALUE_TYPE.VALUE) {
      //todo
    }
    if (type === CELL_VALUE_TYPE.SHEET) {
      toSheet.model = fromSheet.model;
      toSheet.name = getCellSheet(_.get(value, "to"));
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
      const cellValue = fromSheet.getCell(getCellPosition(_.get(value, "from")))
        .value;
      toSheet.getCell(getCellPosition(_.get(value, "to"))).value = moment(
        _.get(cellValue, "result")
      ).format("YYYY-MM-DD");
    } else {
      toSheet.getCell(
        getCellPosition(_.get(value, "to"))
      ).value = fromSheet.getCell(getCellPosition(_.get(value, "from"))).value;
    }
  });
  return to;
}
