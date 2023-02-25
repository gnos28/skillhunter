import * as dotenv from "dotenv";
dotenv.config();
import { appDrive } from "../utils/google";
import { sheetAPI } from "../utils/sheetAPI";
import {
  TAB_COLLAB_COL_COLLAB,
  TAB_COLLAB_COL_EMAIL,
  TAB_COLLAB_COL_SHEET_ID,
  TAB_NAME_COLLAB,
} from "../interfaces/const";
import { createNewSheet } from "./createNewSheetService";
import { importDatas } from "./importDatasService";
import { updateWholeDatas } from "./updateWholeDatasService";
import { getBodyFromFs } from "./getBodyFromFs";
import { exportProgression } from "./exportProgression";

const encodeBase64 = (data: string) => {
  return Buffer.from(data).toString("base64");
};

type BuildCollabProps = {
  mainSpreadsheetId?: string | undefined;
  folderId?: string | undefined;
  trameId?: string | undefined;
};

export const buildCollab = async ({
  mainSpreadsheetId: argMainSpreadsheetId,
  folderId: argFolderId,
  trameId: argTrameId,
}: BuildCollabProps) => {
  let mainSpreadsheetId: string | undefined = undefined;
  let folderId: string | undefined = undefined;
  let trameId: string | undefined = undefined;

  if (
    argMainSpreadsheetId === undefined ||
    argFolderId === undefined ||
    argTrameId === undefined
  ) {
    const storedFile = await getBodyFromFs();

    mainSpreadsheetId = storedFile.mainSpreadsheetId;
    folderId = storedFile.folderId;
    trameId = storedFile.trameId;
  } else {
    mainSpreadsheetId = argMainSpreadsheetId;
    folderId = argFolderId;
    trameId = argTrameId;
  }

  if (
    mainSpreadsheetId === undefined ||
    folderId === undefined ||
    trameId === undefined
  )
    throw new Error("missing id");

  await exportProgression.init({
    spreadsheetId: mainSpreadsheetId,
    actionName: "buildCollab",
    nbIncrement: 1,
  });

  // clear cache
  sheetAPI.clearCache();

  const driveApp = appDrive();

  const tabList = await sheetAPI.getTabIds(mainSpreadsheetId);

  const collabData = await sheetAPI.getTabData(
    mainSpreadsheetId,
    tabList,
    TAB_NAME_COLLAB
  );

  // vérifier existence du fichier collaborateur
  let nbCreatedFiles = 0;
  let lineIndex = 2;

  const nbCollab = collabData.reduce(
    (acc, val) => (val[TAB_COLLAB_COL_SHEET_ID] ? acc + 1 : acc),
    0
  );
  exportProgression.updateNbIncrement({
    actionName: "buildCollab",
    nbIncrement: nbCollab * 2 + 2,
  });
  await exportProgression.increment({
    actionName: "buildCollab",
  });

  for await (const line of collabData) {
    const collabName = line[TAB_COLLAB_COL_COLLAB];
    const collabEmail = line[TAB_COLLAB_COL_EMAIL];
    let collabId = line[TAB_COLLAB_COL_SHEET_ID];

    if (collabName && collabEmail) {
      let sheetFound = false;
      let collabSheet = null;

      if (collabId) {
        // try {
        // console.log({ collabName, collabEmail, collabId });
        const fileInfo = await driveApp.files.get({
          fileId: collabId,
          fields: "*",
        });

        const isTrashed = fileInfo.data.trashed;

        // console.log(collabName, "fileInfo", fileInfo.data);

        //   collabSheet = SpreadsheetApp.openById(collabId);
        // const isTrashed = DriveApp.getFileById(collabId).isTrashed();
        if (!isTrashed) {
          sheetFound = true;
          console.log(`sheet ${collabName} found 😀`);
        } else console.log(`sheet ${collabName} is trashed 🗑️`);
        // } catch {
        //   console.log(`sheet ${collabName} not found 😱`);
        // }
      }

      if (!sheetFound) {
        console.log(`sheet ${collabName} not found 😱`);
        // si existe pas >> créer le fichier
        collabId = await createNewSheet({
          collabName,
          collabEmail,
          mainSpreadsheetId,
          trameId,
          folderId,
          tabList,
        });
        console.log("collabId", collabId);

        await sheetAPI.updateRange({
          sheetId: mainSpreadsheetId,
          tabName: TAB_NAME_COLLAB,
          startCoords: [lineIndex, 4],
          data: [[collabId]],
        });

        // collabData[lineIndex][3] = collabId;
        // collabListRange.setValues(collabData);
        nbCreatedFiles++;
      } // si existe >> mettre à jour data
      else
        await updateWholeDatas({
          collabName,
          tabList,
          collabFileId: collabId,
          mainSpreadsheetId,
          forceContratUpdate: false,
        });

      await exportProgression.increment({
        actionName: "buildCollab",
      });
    }
    lineIndex++;
  }

  await importDatas({ emailAlert: true, mainSpreadsheetId, tabList });

  console.log("****** END OF buildCollab FUNCTION ******");

  return collabData;
};
