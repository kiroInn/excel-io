import { CELL_VALUE_TYPE } from "@/service/transform";
import _ from "lodash";

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
