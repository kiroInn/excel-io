import { getUsingSheet, CELL_VALUE_TYPE } from "@/service/transform";
import { transformMappings } from "@/service/mapping";

describe("service", () => {
  let mappings: Array<object>;
  beforeEach(() => {
    mappings = [
      {
        templateName: "196000.xlsx",
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
            from: "TB",
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
  });
  it("Given sheet cell metadata When pass sheet metadata Then should return transform sheetName", () => {
    expect(transformMappings(mappings)).toEqual([
      {
        from: "196000:C6",
        to: "196000:C8",
        toFile: "196000.xlsx",
        type: "string"
      },
      {
        from: "196000:A3",
        to: "196000:A4",
        toFile: "196000.xlsx",
        type: "date"
      },
      {
        from: "TB:C6",
        to: "TB:C6",
        toFile: "196000.xlsx",
        type: "string"
      },
      {
        from: "TB:C6",
        to: "TB:C7",
        toFile: "196000.xlsx",
        type: "string"
      },
      {
        from: "TB",
        range: {
          br: {
            col: 6.5,
            row: 15.5
          },
          tl: {
            col: 0,
            row: 10.5
          }
        },
        to: "TB",
        toFile: "196000.xlsx",
        type: "image"
      },
      {
        from: "106700:A3",
        to: "Reconciliation:A4",
        toFile: "106700.xlsx",
        type: "date"
      },
      {
        from: "106700:C6",
        to: "Reconciliation:C10",
        toFile: "106700.xlsx",
        type: "string"
      },
      {
        from: "106700",
        range: {
          br: {
            col: 11,
            row: 26.5
          },
          tl: {
            col: 4.5,
            row: 10.5
          }
        },
        to: "Reconciliation",
        toFile: "106700.xlsx",
        type: "image"
      }
    ]);
  });
  it("Given sheet cell metadata When pass sheet metadata Then should return include sheetName", () => {
    expect(getUsingSheet(mappings, "from")).toEqual(["196000", "TB", "106700"]);
    expect(getUsingSheet(mappings, "to")).toEqual([
      "196000",
      "TB",
      "Reconciliation"
    ]);
  });
});
