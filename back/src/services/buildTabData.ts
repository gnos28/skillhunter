import { TabListItem, ExtRow } from "../interfaces";
import { mapTrimObj } from "../utils/mapTrimObj";
import { sheetAPI } from "../utils/sheetAPI";

type BuildTabDataProps = {
  mainSpreadsheetId: string;
  tabList: TabListItem[];
  tabName: string;
  collabName?: string;
  colToKeep: string[];
  filterByCol?: string;
  headerRowIndex?: number;
};

export const buildTabData = async ({
  mainSpreadsheetId,
  tabList,
  tabName,
  collabName,
  filterByCol,
  headerRowIndex,
  colToKeep,
}: BuildTabDataProps) => [
  ...(
    await sheetAPI.getTabData(
      mainSpreadsheetId,
      tabList,
      tabName,
      headerRowIndex
    )
  )
    .filter((row) => (filterByCol ? row[filterByCol] === collabName : true))
    .map((row) => mapTrimObj(row, colToKeep)),
  mapTrimObj({ rowIndex: -1, a1Range: "" } as ExtRow, colToKeep),
];
