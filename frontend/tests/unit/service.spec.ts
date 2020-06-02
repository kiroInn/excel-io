import { getUsingSheet, CELL_VALUE_TYPE } from "@/service/transform";

describe("service", () => {
  it("Given sheet cell metadata When pass sheet metadata Then should return include sheetName", () => {
    const mappings = [
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
    expect(getUsingSheet(mappings, "from")).toEqual(["196000", "TB", "106700"]);
    expect(getUsingSheet(mappings, "to")).toEqual([
      "196000",
      "TB",
      "Reconciliation"
    ]);
  });
});
