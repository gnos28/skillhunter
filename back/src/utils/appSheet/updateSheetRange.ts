import AlphanumericEncoder from "alphanumeric-encoder";
import { appSheet } from "../google";

type UpdateSheetRange = {
  sheetId: string;
  tabName: string;
  startCoords: [number, number];
  data: any[][];
};

export const updateSheetRange = async ({
  sheetId,
  tabName,
  startCoords,
  data,
}: UpdateSheetRange) => {
  const sheetApp = appSheet();

  const encoder = new AlphanumericEncoder();

  const encodedStartCol = encoder.encode(startCoords[1] || 1);
  const encodedEndCol = encoder.encode(
    (startCoords[1] || 1) - 1 + data[0].length
  );

  const rangeA1notation = `'${tabName}'!${encodedStartCol}${
    startCoords[0] || 1
  }:${encodedEndCol}${data.length + (startCoords[0] || 1) - 1}`;

  await sheetApp.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: rangeA1notation,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: data,
    },
  });
};
