import { isXlsx } from "@/util/file";

describe("file util", () => {
  it("Given Determine file type When pass valid fileName Then should return true", () => {
    expect(isXlsx("xxx.xlsx")).toBe(true);
    expect(isXlsx("有用的.xlsx")).toBe(true);
    expect(isXlsx("2020-04-25.xlsx")).toBe(true);
  });

  it("Given Determine file type When pass inValid fileName Then should return false", () => {
    expect(isXlsx("xxx.xx")).toBe(false);
    expect(isXlsx("有用的.json")).toBe(false);
    expect(isXlsx("2020-04-25.json")).toBe(false);
    expect(isXlsx("xxxx")).toBe(false);
  })
});
