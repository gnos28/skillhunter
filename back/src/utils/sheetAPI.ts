import { AddProtectedRangeProps, batchUpdate } from "./appSheet/batchUpdate";
import { getSheetTabIds } from "./appSheet/getSheetTabIds";
import { appSheet } from "./google";
import { importSheetData } from "./appSheet/importSheetData";
import { updateSheetRange } from "./appSheet/updateSheetRange";
import { clearTabData } from "./appSheet/clearSheetRows";
import { TabListItem } from "../interfaces";

type TabCache = {
  [key: string]: ({
    [key: string]: string;
  } & {
    rowIndex: number;
    a1Range: string;
  })[];
};

type TabIdsCache = {
  [key: string]: {
    sheetId: string;
    sheetName: string;
  }[];
};

type GetTabMetaDataProps = {
  spreadsheetId: string;
  fields: string;
  ranges: string[];
};

type UpdateSheetRangeProps = {
  sheetId: string;
  tabName: string;
  startCoords: [number, number];
  data: any[][];
};

type GetProtectedRangeIdsProps = {
  spreadsheetId: string;
  sheetId: number;
};

type DeleteProtectedRangeProps = {
  spreadsheetId: string;
  protectedRangeIds: number[];
};

type ClearTabDataProps = {
  sheetId: string;
  tabList: TabListItem[];
  tabName: string;
  headerRowIndex?: number;
};

let tabCache: TabCache = {};
let tabIdsCache: TabIdsCache = {};
let lastReadRequestTime: number | undefined = undefined;
let lastWriteRequestTime: number | undefined = undefined;
let nbInQueueRead = 0;
let nbInQueueWrite = 0;
let readCatchCount = 0;
let writeCatchCount = 0;

const DELAY = 300; // in ms
const CATCH_DELAY_MULTIPLIER = 15;
const MAX_CATCH_COUNT = 30;
const MAX_AWAITING_TIME = 120_000;

class MaxAwaitingTimeError extends Error {}

const handleReadTryCatch = async <T>(
  callback: () => Promise<T>,
  readCatchCount: number,
  delayMultiplier?: number
) => {
  let res: T | undefined = undefined;
  let timeout: NodeJS.Timeout | undefined = undefined;

  try {
    timeout = setTimeout(() => {
      console.log("[READ] MAX_AWAITING_TIME reached 💀");
      throw new MaxAwaitingTimeError();
    }, MAX_AWAITING_TIME);

    res = await callback();

    clearTimeout(timeout);
    lastReadRequestTime = new Date().getTime();
    nbInQueueRead -= delayMultiplier || 1;
  } catch (e: any) {
    console.log(
      `inside catch 💩 READ#${readCatchCount}`,
      callback.name,
      e.message
    );
    readCatchCount++;
    lastReadRequestTime = new Date().getTime();
    nbInQueueRead -= delayMultiplier || 1;
    clearTimeout(timeout);
    if (e instanceof MaxAwaitingTimeError && writeCatchCount > 1)
      readCatchCount = MAX_CATCH_COUNT;

    if (readCatchCount < MAX_CATCH_COUNT)
      res = await handleReadDelay(
        callback,
        readCatchCount,
        CATCH_DELAY_MULTIPLIER
      );
  } finally {
    readCatchCount = 0;
    return res as T;
  }
};

const handleReadDelay = async <T>(
  callback: () => Promise<T>,
  readCatchCount: number = 0,
  delayMultiplier?: number
) => {
  const currentTime = new Date().getTime();
  nbInQueueRead += delayMultiplier || 1;

  if (
    lastReadRequestTime &&
    currentTime < lastReadRequestTime + DELAY * nbInQueueRead
  ) {
    console.log("*** force DELAY [READ] ", {
      nbInQueueRead: nbInQueueRead / (delayMultiplier || 1),
      timeout: lastReadRequestTime
        ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
        : 0,
    });
    await new Promise((resolve) =>
      setTimeout(
        () => resolve(null),
        lastReadRequestTime
          ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
          : 0
      )
    );
  }

  const res: T = await handleReadTryCatch(
    callback,
    readCatchCount,
    delayMultiplier
  );

  return res;
};

const handleWriteTryCatch = async <T>(
  callback: () => Promise<T>,
  writeCatchCount: number,
  delayMultiplier?: number
) => {
  let res: T | undefined = undefined;
  let timeout: NodeJS.Timeout | undefined = undefined;

  try {
    timeout = setTimeout(() => {
      console.log("[READ] MAX_AWAITING_TIME reached 💀");
      throw new MaxAwaitingTimeError();
    }, MAX_AWAITING_TIME);

    res = await callback();

    clearTimeout(timeout);
    lastWriteRequestTime = new Date().getTime();
    nbInQueueWrite -= delayMultiplier || 1;
  } catch (e: any) {
    console.log(
      `inside catch 💩 WRITE#${writeCatchCount}`,
      callback.name,
      e.message
    );
    writeCatchCount++;
    lastWriteRequestTime = new Date().getTime();
    nbInQueueWrite -= delayMultiplier || 1;
    clearTimeout(timeout);
    if (e instanceof MaxAwaitingTimeError && writeCatchCount > 1)
      writeCatchCount = MAX_CATCH_COUNT;

    if (writeCatchCount < MAX_CATCH_COUNT)
      res = await handleWriteDelay(
        callback,
        writeCatchCount,
        CATCH_DELAY_MULTIPLIER
      );
  } finally {
    writeCatchCount = 0;
    return res as T;
  }
};

const handleWriteDelay = async <T>(
  callback: () => Promise<T>,
  writeCatchCount: number = 0,
  delayMultiplier?: number
) => {
  const currentTime = new Date().getTime();
  nbInQueueWrite += delayMultiplier || 1;

  if (
    lastWriteRequestTime &&
    currentTime < lastWriteRequestTime + DELAY * nbInQueueWrite
  ) {
    console.log("*** force DELAY [WRITE]", {
      nbInQueueWrite: nbInQueueWrite / (delayMultiplier || 1),
      timeout: lastWriteRequestTime
        ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
        : 0,
    });
    await new Promise((resolve) =>
      setTimeout(
        () => resolve(null),
        lastWriteRequestTime
          ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
          : 0
      )
    );
  }

  const res: T = await handleWriteTryCatch(
    callback,
    writeCatchCount,
    delayMultiplier
  );

  return res;
};

export const sheetAPI = {
  getTabIds: async (sheetId: string | undefined) => {
    console.log("*** sheetAPI.getTabIds", sheetId);
    if (sheetId) {
      const cacheKey = sheetId;
      if (tabIdsCache[cacheKey] === undefined) {
        await handleReadDelay(async () => {
          tabIdsCache[cacheKey] = await getSheetTabIds(sheetId);
        });
      } else console.log("*** using cache 👍");

      return tabIdsCache[cacheKey];
    }
    return [];
  },

  getTabData: async (
    sheetId: string,
    tabList: TabListItem[],
    tabName: string,
    headerRowIndex?: number
  ) => {
    console.log("*** sheetAPI.getTabData", tabName);

    const tabId = tabList.filter((tab) => tab.sheetName === tabName)[0]
      ?.sheetId;
    if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

    const cacheKey = sheetId + ":" + tabId;

    if (tabCache[cacheKey] === undefined) {
      await handleReadDelay(async () => {
        tabCache[cacheKey] = await importSheetData(
          sheetId,
          tabId,
          headerRowIndex
        );
      });
    } else console.log("*** using cache 👍");

    return tabCache[cacheKey];
  },

  getTabMetaData: async ({
    spreadsheetId,
    fields,
    ranges,
  }: GetTabMetaDataProps) => {
    console.log("*** sheetAPI.getTabMetaData");

    const metaData = await handleReadDelay(async () => {
      const sheetApp = appSheet();

      return await sheetApp.spreadsheets.get({
        spreadsheetId,
        fields,
        ranges,
      });
    });

    return metaData;
  },

  clearCache: () => {
    tabCache = {};
    tabIdsCache = {};
  },

  updateRange: async ({
    sheetId,
    tabName,
    startCoords,
    data,
  }: UpdateSheetRangeProps) => {
    console.log("*** sheetAPI.updateRange");

    await handleWriteDelay(async () => {
      await updateSheetRange({
        sheetId,
        tabName,
        startCoords,
        data,
      });
    });
  },

  getProtectedRangeIds: async ({
    spreadsheetId,
    sheetId,
  }: GetProtectedRangeIdsProps) => {
    return await handleReadDelay(async () => {
      return await batchUpdate.getProtectedRangeIds(spreadsheetId, sheetId);
    });
  },

  deleteProtectedRange: async ({
    spreadsheetId,
    protectedRangeIds,
  }: DeleteProtectedRangeProps) => {
    await handleWriteDelay(async () => {
      await batchUpdate.deleteProtectedRange(spreadsheetId, protectedRangeIds);
    });
  },

  logBatchProtectedRange: () => {
    const batchProtectedRange = batchUpdate.getBatchProtectedRange();

    console.log("%%%%% logBatchProtectedRange");

    Object.keys(batchProtectedRange).forEach((key) => {
      console.log(key, batchProtectedRange[key].length);
    });

    console.log("%%%%% logBatchProtectedRange");
  },

  clearProtectedRangeBuffer: () => {
    batchUpdate.clearBuffer();
  },

  addBatchProtectedRange: ({
    spreadsheetId,
    editors,
    namedRangeId,
    sheetId,
    startColumnIndex,
    startRowIndex,
    endColumnIndex,
    endRowIndex,
  }: AddProtectedRangeProps) => {
    batchUpdate.addProtectedRange({
      spreadsheetId,
      editors,
      namedRangeId,
      sheetId,
      startColumnIndex,
      startRowIndex,
      endColumnIndex,
      endRowIndex,
    });
  },

  runBatchProtectedRange: async (spreadsheetId: string) => {
    console.log("*** sheetAPI.runBatchProtectedRange");

    await handleWriteDelay(async () => {
      await batchUpdate.runProtectedRange(spreadsheetId);
    });
  },

  clearTabData: async ({
    sheetId,
    tabList,
    tabName,
    headerRowIndex,
  }: ClearTabDataProps) => {
    console.log("*** sheetAPI.clearTabData");

    await handleWriteDelay(async () => {
      await clearTabData(sheetId, tabList, tabName, headerRowIndex);
    });
  },
};
