import { CELL_VALUE_TYPE } from "@/service/transform";
import _ from "lodash";

export const DEFAULT_MAPPING = [
  {
    templateName: "196000.xlsx",
    fileName: "${date}-${companyName}-196000.xlsx",
    values: [
      {
        from: "196000:C6",
        to: "196000:C8",
        type: CELL_VALUE_TYPE.STRING
      },
      { from: "196000:A3", to: "196000:A4", type: CELL_VALUE_TYPE.DATE },
      { from: "TB:C6", to: "TB:C6", type: CELL_VALUE_TYPE.STRING },
      { from: "TB:C6", to: "TB:C7", type: CELL_VALUE_TYPE.STRING },
      {
        from: "196000",
        to: "TB",
        range: {
          tl: { col: 0, row: 10.5 },
          br: { col: 6.5, row: 15.5 }
        },
        type: CELL_VALUE_TYPE.IMAGE
      }
    ]
  },
  {
    templateName: "106700.xlsx",
    fileName: "${date}-${companyName}-106700.xlsx",
    values: [
      {
        from: "106700:A3",
        to: "Reconciliation:A4",
        type: CELL_VALUE_TYPE.DATE
      },
      {
        from: "106700:C6",
        to: "Reconciliation:C10",
        type: CELL_VALUE_TYPE.STRING
      },
      {
        from: "106700",
        to: "Reconciliation",
        range: {
          tl: { col: 4.5, row: 10.5 },
          br: { col: 11, row: 26.5 }
        },
        type: CELL_VALUE_TYPE.IMAGE
      }
    ]
  }
];

export function transformMappings(mappings: Array<object>): Array<object> {
  let result: Array<object> = [];
  _.forEach(mappings, mapping => {
    const toFile = _.get(mapping, "templateName");
    result = _.concat(
      result,
      _.map(_.get(mapping, "values"), mp => ({ ...mp, toFile }))
    );
  });
  return result;
}
