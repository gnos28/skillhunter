import { TabListItem } from "./clearSheetRows";
import { getSheetTabIds } from "./getSheetTabIds";
import { importSheetData } from "./importSheetData";

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

let tabCache: TabCache = {};
let tabIdsCache: TabIdsCache = {};
let lastRequestTime: number | undefined = undefined;

const DELAY = 5000; // in ms

export const tabData = {
  getTabIds: async (sheetId: string | undefined) => {
    console.log("*** tabData.getTabIds", sheetId);
    if (sheetId) {
      const cacheKey = sheetId;
      if (tabIdsCache[cacheKey] === undefined) {
        const currentTime = new Date().getTime();

        if (lastRequestTime && currentTime < lastRequestTime + DELAY) {
          console.log(
            "*** force DELAY",
            lastRequestTime ? lastRequestTime + DELAY - currentTime : 0
          );
          await new Promise((resolve) =>
            setTimeout(
              () => resolve(null),
              lastRequestTime ? lastRequestTime + DELAY - currentTime : 0
            )
          );
        }

        tabIdsCache[cacheKey] = await getSheetTabIds(sheetId);
        lastRequestTime = new Date().getTime();
      } else console.log("*** using cache 👍");

      return tabIdsCache[cacheKey];
    }
    return [];
  },

  get: async (
    sheetId: string,
    tabList: TabListItem[],
    tabName: string,
    headerRowIndex?: number
  ) => {
    console.log("*** tabData.get", sheetId, tabName);

    const tabId = tabList.filter((tab) => tab.sheetName === tabName)[0]
      ?.sheetId;
    if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

    const cacheKey = sheetId + ":" + tabId;

    if (tabCache[cacheKey] === undefined) {
      const currentTime = new Date().getTime();

      if (lastRequestTime && currentTime < lastRequestTime + DELAY) {
        console.log(
          "*** force DELAY",
          lastRequestTime ? lastRequestTime + DELAY - currentTime : 0
        );
        await new Promise((resolve) =>
          setTimeout(
            () => resolve(null),
            lastRequestTime ? lastRequestTime + DELAY - currentTime : 0
          )
        );
      }

      tabCache[cacheKey] = await importSheetData(
        sheetId,
        tabId,
        headerRowIndex
      );

      lastRequestTime = new Date().getTime();
    } else console.log("*** using cache 👍");

    return tabCache[cacheKey];
  },
  clearCache: () => {
    // clear cache
    tabCache = {};
  },
};
