import * as dotenv from "dotenv";
import { BaseRow } from "../interfaces";
import {
  COL2KEEP_CLIENTS,
  COL2KEEP_CONTRATS,
  COL2KEEP_CONTRATS_IMPORT,
  TAB_COLLAB_COL_COLLAB,
  TAB_COLLAB_COL_EMAIL,
  TAB_COLLAB_COL_SHEET_ID,
  TAB_CONTRATS_COL_CANDIDAT,
  TAB_CONTRATS_COL_CLIENT,
  TAB_CONTRATS_COL_COLLAB,
  TAB_CONTRATS_COL_DATE_DEBUT,
  TAB_CONTRATS_COL_ID,
  TAB_CONTRATS_COL_IMPORT_ID,
  TAB_IMPORT_DATA,
  TAB_NAME_CLIENTS,
  TAB_NAME_COLLAB,
  TAB_NAME_CONTRATS,
} from "../interfaces/const";
import { getValuesFromBaseRow } from "../utils/getValuesFromBaseRow";
import { appDrive } from "../utils/google";
import { sheetAPI } from "../utils/sheetAPI";
import { buildTabData } from "./buildTabData";
import { exportProgression } from "./exportProgression";
dotenv.config();

type ExportCorrectionsProps = {
  mainSpreadsheetId: string | undefined;
};

export const exportCorrections = async ({
  mainSpreadsheetId,
}: ExportCorrectionsProps) => {
  if (mainSpreadsheetId === undefined) throw new Error("missing id");

  await exportProgression.init({
    spreadsheetId: mainSpreadsheetId,
    actionName: "exportCorrections",
    nbIncrement: 1,
  });

  const tabList = await sheetAPI.getTabIds(mainSpreadsheetId);

  const collabData = await sheetAPI.getTabData(
    mainSpreadsheetId,
    tabList,
    TAB_NAME_COLLAB
  );

  const contratData = await sheetAPI.getTabData(
    mainSpreadsheetId,
    tabList,
    TAB_NAME_CONTRATS,
    2
  );

  const contratDataBuilded = await buildTabData({
    mainSpreadsheetId,
    tabList,
    colToKeep: COL2KEEP_CONTRATS,
    tabName: TAB_NAME_CONTRATS,
    headerRowIndex: 2,
  });

  const contratImportData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    colToKeep: COL2KEEP_CONTRATS_IMPORT,
    tabName: TAB_NAME_CONTRATS,
    headerRowIndex: 2,
  });

  const contratsImportValues = getValuesFromBaseRow([contratImportData]);

  const clientsData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    colToKeep: COL2KEEP_CLIENTS,
    tabName: TAB_NAME_CLIENTS,
  });

  const clientValues = getValuesFromBaseRow([clientsData]);

  const allContratsImportIdIndex = Object.keys(contratData[0]).findIndex(
    (col) => col === TAB_CONTRATS_COL_IMPORT_ID
  );

  const nbCollab = collabData.reduce(
    (acc, val) => (val[TAB_COLLAB_COL_SHEET_ID] ? acc + 1 : acc),
    0
  );

  exportProgression.updateNbIncrement({
    actionName: "exportCorrections",
    nbIncrement: nbCollab + 2,
  });
  await exportProgression.increment({
    actionName: "exportCorrections",
  });

  for await (const line of collabData) {
    const collabName = line[TAB_COLLAB_COL_COLLAB];
    const collabEmail = line[TAB_COLLAB_COL_EMAIL];
    let collabId = line[TAB_COLLAB_COL_SHEET_ID];
    const driveApp = appDrive();

    if (collabName && collabEmail) {
      let sheetFound = false;
      let collabSheet = null;

      if (collabId) {
        const fileInfo = await driveApp.files.get({
          fileId: collabId,
          fields: "*",
        });

        const isTrashed = fileInfo.data.trashed;

        if (!isTrashed) {
          sheetFound = true;
          console.log(`sheet ${collabName} found 😀`);
        } else console.log(`sheet ${collabName} is trashed 🗑️`);
      }

      if (!sheetFound) {
        console.log(`sheet ${collabName} not found 😱`);
      } // si existe >> exporter corrections
      else {
        // MAJ listing clients
        const collabTabList = await sheetAPI.getTabIds(collabId);

        // udpate clientsData in IMPORT_DATAS sheet
        await sheetAPI.updateRange({
          sheetId: collabId,
          tabName: TAB_IMPORT_DATA,
          startCoords: [3, 6],
          data: clientValues,
        });

        const collabContratData = await buildTabData({
          mainSpreadsheetId: collabId,
          tabList: collabTabList,
          collabName,
          colToKeep: COL2KEEP_CONTRATS,
          tabName: TAB_NAME_CONTRATS,
          filterByCol: TAB_CONTRATS_COL_COLLAB,
        });

        const collabContratIdIndex = Object.keys(
          collabContratData[0]
        ).findIndex((col) => col === TAB_CONTRATS_COL_ID);

        const collabContratDateDebutIndex = Object.keys(
          collabContratData[0]
        ).findIndex((col) => col === TAB_CONTRATS_COL_DATE_DEBUT);

        const collabContratClientIndex = Object.keys(
          collabContratData[0]
        ).findIndex((col) => col === TAB_CONTRATS_COL_CLIENT);

        const collabContratCandidatIndex = Object.keys(
          collabContratData[0]
        ).findIndex((col) => col === TAB_CONTRATS_COL_CANDIDAT);

        const collabContratValues = getValuesFromBaseRow([collabContratData]);

        const updatedCollabContratValues = collabContratValues.map(
          (collabContratRow) => {
            if (
              !collabContratRow[collabContratDateDebutIndex] ||
              !collabContratRow[collabContratClientIndex] ||
              !collabContratRow[collabContratCandidatIndex]
            )
              return collabContratRow;

            const filteredContratData = contratData.filter((allContratRow) => {
              return (
                allContratRow[TAB_CONTRATS_COL_ID] ===
                collabContratRow[collabContratIdIndex]
              );
            });

            if (!filteredContratData.length) return collabContratRow;

            const rowIndex = filteredContratData[0].rowIndex - 3;

            contratsImportValues[rowIndex] = Object.values(
              contratDataBuilded[rowIndex]
            );

            return Object.values(contratDataBuilded[rowIndex]);
          }
        );

        await sheetAPI.updateRange({
          sheetId: collabId,
          tabName: TAB_NAME_CONTRATS,
          startCoords: [2, 1],
          data: updatedCollabContratValues,
        });
      }
      await exportProgression.increment({
        actionName: "exportCorrections",
      });
    }
  }

  await sheetAPI.updateRange({
    sheetId: mainSpreadsheetId,
    tabName: TAB_NAME_CONTRATS,
    startCoords: [3, allContratsImportIdIndex + 1],
    data: contratsImportValues,
  });

  await exportProgression.increment({
    actionName: "exportCorrections",
  });

  console.log("****** END OF exportCorrections FUNCTION ******");

  return "";
};
