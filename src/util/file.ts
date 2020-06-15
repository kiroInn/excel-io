import _ from "lodash";

export function isXlsx(fileName: string) {
  const parts = fileName.split(".");
  const suffix = parts[parts.length - 1];
  return _.includes(["xlsx", "xls"], suffix.toLowerCase());
}
