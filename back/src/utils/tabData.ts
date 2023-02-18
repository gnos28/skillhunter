import { TabListItem } from "./clearSheetRows";
import { importSheetData } from "./importSheetData";

type TabCache = {
  [key: string]: {
    [key: string]: string;
  }[];
};

let tabCache: TabCache = {};

export const tabData = {
  get: async (
    sheetId: string,
    tabList: TabListItem[],
    tabName: string,
    headerRowIndex?: number
  ) => {
    const tabId = tabList.filter((tab) => tab.sheetName === tabName)[0]
      ?.sheetId;
    if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

    const cacheKey = sheetId + ":" + tabId;

    if (tabCache[cacheKey] === undefined)
      tabCache[cacheKey] = await importSheetData(
        sheetId,
        tabId,
        headerRowIndex
      );

    return tabCache[cacheKey];
  },
  clearCache: () => {
    // clear cache
    tabCache = {};
  },
};
