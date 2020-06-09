import { CELL_VALUE_TYPE } from "@/service/transform";
import _ from "lodash";

export const DEFAULT_MAPPING = [
  {
    templateName: "106700.xlsx",
    values: [
      { from: "106700-Reconciliation", to: "Reconciliation", type: "sheet" },
      {
        from: "106700-Outstanding Checks",
        to: "Outstanding Checks",
        type: "sheet"
      },
      {
        from: "106700-Deposits in Transit",
        to: "Deposits in Transit",
        type: "sheet"
      },
      { from: "106700-Instructions", to: "Instructions", type: "sheet" }
    ]
  },
  {
    templateName: "106710.xlsx",
    values: [
      { from: "106710-Reconciliation", to: "Reconciliation", type: "sheet" },
      { from: "106710-Instructions", to: "Instructions", type: "sheet" },
      {
        from: "106710-Outstanding Checks",
        to: "Outstanding Checks",
        type: "sheet"
      },
      {
        from: "106710-Deposits in Transit",
        to: "Deposits in Transit",
        type: "sheet"
      }
    ]
  },
  {
    templateName: "106750.xlsx",
    values: [
      { from: "106750-Reconciliation", to: "Reconciliation", type: "sheet" },
      { from: "106750-Instructions", to: "Instructions", type: "sheet" },
      {
        from: "106750-Outstanding Checks",
        to: "Outstanding Checks",
        type: "sheet"
      },
      {
        from: "106750-Deposits in Transit",
        to: "Deposits in Transit",
        type: "sheet"
      }
    ]
  },
  {
    templateName: "112000.xlsx",
    values: [
      { from: "112000", to: "112000", type: "sheet" },
      { from: "112000-1", to: "1", type: "sheet" }
    ]
  },
  {
    templateName: "119000.xlsx",
    values: [{ from: "119000", to: "119000", type: "sheet" }]
  },
  {
    templateName: "190500.xlsx",
    values: [
      { from: "190500", to: "190500", type: "sheet" },
      { from: "190500-1", to: "1", type: "sheet" }
    ]
  },
  {
    templateName: "196000.xlsx",
    values: [
      { from: "196000", to: "196000", type: "sheet" },
      { from: "196000-1", to: "1", type: "sheet" }
    ]
  },
  {
    templateName: "200000.xlsx",
    values: [
      { from: "200000", to: "200000", type: "sheet" },
      { from: "200000-1", to: "1", type: "sheet" },
      { from: "200000-2", to: "2", type: "sheet" }
    ]
  },
  {
    templateName: "205000.xlsx",
    values: [
      { from: "205000", to: "205000", type: "sheet" },
      { from: "205000-1", to: "1", type: "sheet" }
    ]
  },
  {
    templateName: "215000.xlsx",
    values: [
      { from: "215000", to: "215000", type: "sheet" },
      { from: "215000-1", to: "1", type: "sheet" },
      { from: "215000-Jan", to: "Jan", type: "sheet" },
      { from: "215000-Feb", to: "Feb", type: "sheet" },
      { from: "215000-Mar", to: "Mar", type: "sheet" },
      { from: "215000-Apr", to: "Apr", type: "sheet" },
      { from: "215000-May", to: "May", type: "sheet" }
    ]
  },
  {
    templateName: "215100.xlsx",
    values: [
      { from: "215100-1", to: "1", type: "sheet" },
      { from: "215100", to: "215100", type: "sheet" }
    ]
  },
  {
    templateName: "215300.xlsx",
    values: [
      { from: "215300", to: "215300", type: "sheet" },
      { from: "215300-1", to: "1", type: "sheet" }
    ]
  },
  {
    templateName: "215600.xlsx",
    values: [
      { from: "215600", to: "215600", type: "sheet" },
      { from: "215600-1", to: "1", type: "sheet" },
      { from: "215600-Jan", to: "Jan", type: "sheet" },
      { from: "215600-Feb", to: "Feb", type: "sheet" },
      { from: "215600-Mar", to: "Mar", type: "sheet" },
      { from: "215600-Apr", to: "Apr", type: "sheet" },
      { from: "215600-May", to: "May", type: "sheet" }
    ]
  },
  {
    templateName: "216000.xlsx",
    values: [
      { from: "216000", to: "216000", type: "sheet" },
      { from: "216000-1", to: "1", type: "sheet" }
    ]
  },
  {
    templateName: "217500.xlsx",
    values: [
      { from: "217500", to: "217500", type: "sheet" },
      { from: "217500-1", to: "1", type: "sheet" },
      { from: "217500-Jan", to: "Jan", type: "sheet" },
      { from: "217500-Feb", to: "Feb", type: "sheet" },
      { from: "217500-Mar", to: "Mar", type: "sheet" },
      { from: "217500-Apr", to: "Apr", type: "sheet" },
      { from: "217500-May", to: "May", type: "sheet" }
    ]
  },
  {
    templateName: "274000.xlsx",
    values: [
      { from: "274000", to: "274000", type: "sheet" },
      { from: "274000-1", to: "1", type: "sheet" }
    ]
  },
  {
    templateName: "897100.xlsx",
    values: [{ from: "897100", to: "897100", type: "sheet" }]
  },
  {
    templateName: "897600.xlsx",
    values: [{ from: "897600", to: "897600", type: "sheet" }]
  }
];

export function transformMappings(mappings: Array<object>): Array<object> {
  let result: Array<object> = [];
  _.forEach(mappings, mapping => {
    const toFile = _.get(mapping, "templateName");
    const range = {
      tl: { col: 0, row: 0 },
      br: { col: 1, row: 1 }
    };
    result = _.concat(
      result,
      _.map(_.get(mapping, "values"), mp => ({ range, ...mp, toFile }))
    );
  });
  return result;
}

export function reverseTransformMappings(mappings: object[]): object[] {
  return _.reduce(
    mappings,
    (result: object[], curr: object) => {
      const item = _.find(
        result,
        item => _.get(item, "templateName") === _.get(curr, "toFile")
      );
      const value = {
        from: _.get(curr, "from"),
        to: _.get(curr, "to"),
        type: _.get(curr, "type")
      };
      if (_.get(curr, "type") === CELL_VALUE_TYPE.IMAGE)
        _.assign(value, { range: _.get(curr, "range") });
      if (item) {
        _.get(item, "values", []).push(value);
      } else {
        result.push({
          templateName: _.get(curr, "toFile"),
          values: [value]
        });
      }
      return result;
    },
    []
  );
}
