import { ExtRow, TabListItem } from "../interfaces";
import {
  COL2KEEP_CLIENTS,
  COL2KEEP_COLLAB,
  COL2KEEP_CONTRATS,
  COL2KEEP_VARIABLE,
  COL2KEEP_VERSEMENT,
  TAB_COLLAB_COL_COLLAB,
  TAB_CONTRATS_COL_COLLAB,
  TAB_IMPORT_DATA,
  TAB_NAME_CLIENTS,
  TAB_NAME_COLLAB,
  TAB_NAME_CONTRATS,
  TAB_NAME_VARIABLE,
  TAB_NAME_VERSEMENT,
  TAB_VARIABLE_COL_COLLAB,
  TAB_VERSEMENT_COL_COLLAB,
} from "../interfaces/const";
import { mapTrimObj } from "../utils/mapTrimObj";
import { sheetAPI } from "../utils/sheetAPI";
import { buildTabData } from "./buildTabData";
import { handleContratUpdate } from "./handleContratUpdateService";

type UpdateWholeDatasProps = {
  tabList: TabListItem[];
  collabName: string;
  collabFileId: string;
  mainSpreadsheetId: string;
  forceContratUpdate?: boolean;
};

export const updateWholeDatas = async ({
  tabList,
  collabName,
  collabFileId,
  mainSpreadsheetId,
  forceContratUpdate,
}: UpdateWholeDatasProps) => {
  console.log("updateWholeDatas", collabName);

  const contratData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    collabName,
    colToKeep: COL2KEEP_CONTRATS,
    tabName: TAB_NAME_CONTRATS,
    filterByCol: TAB_CONTRATS_COL_COLLAB,
    headerRowIndex: 2,
  });

  const versementData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    collabName,
    colToKeep: COL2KEEP_VERSEMENT,
    tabName: TAB_NAME_VERSEMENT,
    filterByCol: TAB_VERSEMENT_COL_COLLAB,
  });

  const clientsData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    collabName,
    colToKeep: COL2KEEP_CLIENTS,
    tabName: TAB_NAME_CLIENTS,
  });

  const variableData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    collabName,
    colToKeep: COL2KEEP_VARIABLE,
    tabName: TAB_NAME_VARIABLE,
    filterByCol: TAB_VARIABLE_COL_COLLAB,
    headerRowIndex: 2,
  });

  const collabData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    collabName,
    colToKeep: COL2KEEP_COLLAB,
    tabName: TAB_NAME_COLLAB,
    filterByCol: TAB_COLLAB_COL_COLLAB,
  });

  const newFileTabList = await sheetAPI.getTabIds(collabFileId);

  //   const importData = await sheetAPI.getTabData(
  //     collabFileId,
  //     newFileTabList,
  //     TAB_IMPORT_DATA,
  //     2
  //   );

  const allData = [versementData, clientsData, collabData, variableData];

  const maxRows = allData.reduce(
    (acc, val) => (val.length > acc ? val.length : acc),
    0
  );

  const values = Array(maxRows)
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

  // effacer les précédentes données
  await sheetAPI.clearTabData({
    sheetId: collabFileId,
    tabList: newFileTabList,
    tabName: TAB_IMPORT_DATA,
    headerRowIndex: 2,
  });

  // udpate data in IMPORT_DATAS sheet
  await sheetAPI.updateRange({
    sheetId: collabFileId,
    tabName: TAB_IMPORT_DATA,
    startCoords: [3, 1],
    data: values,
  });

  if (forceContratUpdate) {
    await handleContratUpdate({
      collabFileId,
      newFileTabList,
      contratData,
      collabName,
    });
  }
};
