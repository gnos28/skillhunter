import { sheets_v4 } from "googleapis";
import { appSheet } from "../google";

export type AddProtectedRangeProps = {
  spreadsheetId: string;
  editors: string[];
  namedRangeId?: string;
  sheetId: number;
  startColumnIndex: number;
  startRowIndex: number;
  endColumnIndex: number;
  endRowIndex: number;
};

const protectedRangeBatchBuffer: {
  [key: string]: sheets_v4.Schema$Request[];
} = {};

export const batchUpdate = {
  addProtectedRange: ({
    spreadsheetId,
    editors,
    namedRangeId,
    sheetId,
    startColumnIndex,
    startRowIndex,
    endColumnIndex,
    endRowIndex,
  }: AddProtectedRangeProps) => {
    if (!protectedRangeBatchBuffer[spreadsheetId])
      protectedRangeBatchBuffer[spreadsheetId] = [];

    const namedRange = namedRangeId
      ? namedRangeId
      : Math.random().toString().split(".")[1];

    protectedRangeBatchBuffer[spreadsheetId].push({
      addProtectedRange: {
        protectedRange: {
          editors: { users: editors },
          description: namedRange,
          range: {
            sheetId,
            startColumnIndex,
            startRowIndex,
            endColumnIndex: endColumnIndex + 1,
            endRowIndex: endRowIndex + 1,
          },
        },
      },
    });
  },

  runProtectedRange: async (spreadsheetId: string) => {
    const sheetApp = appSheet();

    const requests = protectedRangeBatchBuffer[spreadsheetId];
    if (requests) {
      console.log("requests count : ", requests.length);
      // console.log("requests", requests);
      await sheetApp.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests,
        },
      });
      protectedRangeBatchBuffer[spreadsheetId] = [];
    }
  },
};
