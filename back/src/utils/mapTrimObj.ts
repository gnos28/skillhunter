import { ExtRow, BaseRow } from "../interfaces";

type MapTrimObjProps = (row: ExtRow, colsToKeep: string[]) => BaseRow;

export const mapTrimObj: MapTrimObjProps = (row, colsToKeep) => {
  const filteredRow: BaseRow = {};

  colsToKeep.forEach(
    (colName) =>
      (filteredRow[colName] = Object.keys(row).includes(colName)
        ? row[colName]
        : "")
  );

  // if (Object.keys(row).includes("rowIndex"))
  // return {
  //   ...filteredRow,
  //   rowIndex: row.rowIndex,
  //   a1Range: row.a1Range,
  // } as ExtRow;
  return filteredRow;
};
