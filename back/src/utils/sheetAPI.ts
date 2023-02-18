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

const DELAY = 2000; // in ms

export const sheetAPI = {
  getTabIds: async (sheetId: string | undefined) => {
    console.log("*** sheetAPI.getTabIds", sheetId);
    if (sheetId) {
      const cacheKey = sheetId;
      if (tabIdsCache[cacheKey] === undefined) {
        const currentTime = new Date().getTime();

        if (lastReadRequestTime && currentTime < lastReadRequestTime + DELAY) {
          console.log(
            "*** force DELAY",
            lastReadRequestTime ? lastReadRequestTime + DELAY - currentTime : 0
          );
          await new Promise((resolve) =>
            setTimeout(
              () => resolve(null),
              lastReadRequestTime
                ? lastReadRequestTime + DELAY - currentTime
                : 0
            )
          );
        }

        tabIdsCache[cacheKey] = await getSheetTabIds(sheetId);
        lastReadRequestTime = new Date().getTime();
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
      const currentTime = new Date().getTime();

      if (lastReadRequestTime && currentTime < lastReadRequestTime + DELAY) {
        console.log(
          "*** force DELAY",
          lastReadRequestTime ? lastReadRequestTime + DELAY - currentTime : 0
        );
        await new Promise((resolve) =>
          setTimeout(
            () => resolve(null),
            lastReadRequestTime ? lastReadRequestTime + DELAY - currentTime : 0
          )
        );
      }

      tabCache[cacheKey] = await importSheetData(
        sheetId,
        tabId,
        headerRowIndex
      );

      lastReadRequestTime = new Date().getTime();
    } else console.log("*** using cache 👍");

    return tabCache[cacheKey];
  },

  getTabMetaData: async ({
    spreadsheetId,
    fields,
    ranges,
  }: GetTabMetaDataProps) => {
    console.log("*** sheetAPI.getTabMetaData");

    const sheetApp = appSheet();

    const currentTime = new Date().getTime();

    if (lastReadRequestTime && currentTime < lastReadRequestTime + DELAY) {
      console.log(
        "*** force DELAY",
        lastReadRequestTime ? lastReadRequestTime + DELAY - currentTime : 0
      );
      await new Promise((resolve) =>
        setTimeout(
          () => resolve(null),
          lastReadRequestTime ? lastReadRequestTime + DELAY - currentTime : 0
        )
      );
    }

    const metaData = await sheetApp.spreadsheets.get({
      spreadsheetId,
      fields,
      ranges,
    });

    lastReadRequestTime = new Date().getTime();

    return metaData;
  },

  clearCache: () => {
    // clear cache
    tabCache = {};
  },

  updateRange: async ({
    sheetId,
    tabName,
    startCoords,
    data,
  }: UpdateSheetRangeProps) => {
    console.log("*** sheetAPI.updateRange");

    const currentTime = new Date().getTime();

    if (lastWriteRequestTime && currentTime < lastWriteRequestTime + DELAY) {
      console.log(
        "*** force DELAY",
        lastWriteRequestTime ? lastWriteRequestTime + DELAY - currentTime : 0
      );
      await new Promise((resolve) =>
        setTimeout(
          () => resolve(null),
          lastWriteRequestTime ? lastWriteRequestTime + DELAY - currentTime : 0
        )
      );
    }

    await updateSheetRange({
      sheetId,
      tabName,
      startCoords,
      data,
    });

    lastWriteRequestTime = new Date().getTime();
  },
};
