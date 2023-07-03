import { TAB_COLLAB_COL_SHEET_ID, TAB_NAME_COLLAB } from "../interfaces/const";
import { appDrive } from "../utils/google";
import { sheetAPI } from "../utils/sheetAPI";
import { getBodyFromFs } from "./getBodyFromFs";

export const deleteSheet = async (sheetId: string) => {
  const driveApp = appDrive();

  await driveApp.files.delete({ fileId: sheetId });
};

export const resetCollab = async (collabId: string) => {
  console.log(
    "🗑️ 🗑️ 🗑️ 🗑️ resetCollab",
    collabId.substring(0, 8),
    " 🗑️ 🗑️ 🗑️ 🗑️"
  );

  if (!collabId) return;

  const storedFile = await getBodyFromFs();

  const mainSpreadsheetId = storedFile.mainSpreadsheetId;

  const tabList = await sheetAPI.getTabIds(mainSpreadsheetId);

  const collabData = await sheetAPI.getTabData(
    mainSpreadsheetId,
    tabList,
    TAB_NAME_COLLAB
  );

  let lineIndex: number = 2;

  collabData.forEach((line, forEachIndex) => {
    const lineCollabId = line[TAB_COLLAB_COL_SHEET_ID];
    if (lineCollabId === collabId) lineIndex += forEachIndex;
  });

  console.log("lineIndex", lineIndex);

  await sheetAPI.updateRange({
    sheetId: mainSpreadsheetId,
    tabName: TAB_NAME_COLLAB,
    startCoords: [lineIndex, 4],
    data: [[""]],
  });

  await deleteSheet(collabId);
};
