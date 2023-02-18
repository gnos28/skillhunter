import { sheets_v4 } from "googleapis";
import { GaxiosResponse } from "googleapis-common";
import { TabListItem } from "./clearSheetRows";
import { getSheetTabIds } from "./getSheetTabIds";
import { appSheet } from "./google";
import { importSheetData } from "./importSheetData";
import { updateSheetRange } from "./updateSheetRange";

type TabCache = {
  [key: string]: {
    [key: string]: string;
  }[];
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

let tabCache: TabCache = {};
let tabIdsCache: TabIdsCache = {};
let lastReadRequestTime: number | undefined = undefined;
let lastWriteRequestTime: number | undefined = undefined;
let nbInQueueRead = 0;
let nbInQueueWrite = 0;

const DELAY = 2000; // in ms

const handleReadDelay = async <T>(callback: () => Promise<T>) => {
  const currentTime = new Date().getTime();
  nbInQueueRead++;

  if (
    lastReadRequestTime &&
    currentTime < lastReadRequestTime + DELAY * nbInQueueRead
  ) {
    console.log(
      "*** force DELAY",
      nbInQueueRead,
      lastReadRequestTime
        ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
        : 0
    );
    await new Promise((resolve) =>
      setTimeout(
        () => resolve(null),
        lastReadRequestTime
          ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
          : 0
      )
    );
  }

  const res = await callback();

  lastReadRequestTime = new Date().getTime();
  nbInQueueRead--;

  return res;
};

const handleWriteDelay = async <T>(callback: () => Promise<T>) => {
  const currentTime = new Date().getTime();
  nbInQueueWrite++;

  if (
    lastWriteRequestTime &&
    currentTime < lastWriteRequestTime + DELAY * nbInQueueWrite
  ) {
    console.log(
      "*** force DELAY",
      nbInQueueWrite,
      lastWriteRequestTime
        ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
        : 0
    );
    await new Promise((resolve) =>
      setTimeout(
        () => resolve(null),
        lastWriteRequestTime
          ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
          : 0
      )
    );
  }

  const res = await callback();

  lastWriteRequestTime = new Date().getTime();
  nbInQueueWrite--;

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
    console.log("*** sheetAPI.getTabData", sheetId, tabName);

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
    nbInQueueRead++;

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
    // clear cache
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

    // const currentTime = new Date().getTime();

    // if (lastWriteRequestTime && currentTime < lastWriteRequestTime + DELAY) {
    //   console.log(
    //     "*** force DELAY",
    //     lastWriteRequestTime ? lastWriteRequestTime + DELAY - currentTime : 0
    //   );
    //   await new Promise((resolve) =>
    //     setTimeout(
    //       () => resolve(null),
    //       lastWriteRequestTime ? lastWriteRequestTime + DELAY - currentTime : 0
    //     )
    //   );
    // }

    // await updateSheetRange({
    //   sheetId,
    //   tabName,
    //   startCoords,
    //   data,
    // });

    // lastWriteRequestTime = new Date().getTime();
  },
};
