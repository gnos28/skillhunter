import { BaseRow } from "../interfaces";

export const getValuesFromBaseRow = (allData: BaseRow[][]) => {
  const maxRows = allData.reduce(
    (acc, val) => (val.length > acc ? val.length : acc),
    0
  );

  return Array(maxRows)
    .fill(undefined)
    .map((_, rowIndex) => {
      return allData
        .map((data) => {
          if (rowIndex > data.length - 1)
            return Array(Object.keys(data[0]).length).fill("");

          return Object.keys(data[rowIndex]).map((key) => data[rowIndex][key]);
        })
        .flat();
    });
};
