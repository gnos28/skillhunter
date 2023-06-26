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

  clearBuffer: () => {
    Object.keys(protectedRangeBatchBuffer).forEach((key) => {
      protectedRangeBatchBuffer[key] = [];
    });
  },

  getBatchProtectedRange: () => {
    return protectedRangeBatchBuffer;
  },

  runProtectedRange: async (spreadsheetId: string) => {
    const sheetApp = appSheet();

    const requests = protectedRangeBatchBuffer[spreadsheetId];
    if (requests.length > 0) {
      console.log("[runProtectedRange] requests count : ", requests.length);
      // console.log("requests", requests);
      const startTime = new Date().getTime();
      await sheetApp.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests,
        },
      });
      console.log(
        "[runProtectedRange] time taken (ms) : ",
        new Date().getTime() - startTime
      );
      protectedRangeBatchBuffer[spreadsheetId] = [];
    }
  },

  getProtectedRangeIds: async (spreadsheetId: string, sheetId: number) => {
    const sheetApp = appSheet();

    const getResult = await sheetApp.spreadsheets.get({ spreadsheetId });

    const sheets = getResult.data.sheets;
    if (sheets !== undefined) {
      const sheet = sheets.filter(
        (sheet) => sheet.properties?.sheetId === sheetId
      )[0];

      const protectedRanges = sheet.protectedRanges;

      if (protectedRanges !== undefined) {
        return protectedRanges
          .map((protectedRange) => protectedRange.protectedRangeId)
          .filter((id) => id !== null && id !== undefined) as number[];
      }
    }

    return [];
  },

  deleteProtectedRange: async (
    spreadsheetId: string,
    protectedRangeIds: number[]
  ) => {
    const sheetApp = appSheet();

    const requests: sheets_v4.Schema$Request[] = protectedRangeIds.map(
      (protectedRangeId) => ({
        deleteProtectedRange: { protectedRangeId },
      })
    );
    console.log("[deleteProtectedRange] requests count : ", requests.length);
    // console.log("requests", requests);
    const startTime = new Date().getTime();

    await sheetApp.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests,
      },
    });

    console.log(
      "[deleteProtectedRange] time taken (ms) : ",
      new Date().getTime() - startTime
    );
  },
};
