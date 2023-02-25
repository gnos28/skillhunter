import { TabListItem } from "../interfaces";
import {
  TAB_NAME_PARAMETRES,
  TAB_PARAMETRES_COL_EMAIL,
} from "../interfaces/const";
import { appDrive } from "../utils/google";
import { sheetAPI } from "../utils/sheetAPI";
import { updateWholeDatas } from "./updateWholeDatasService";

type CreateNewSheetProps = {
  collabName: string;
  collabEmail: string;
  mainSpreadsheetId: string;
  trameId: string;
  folderId: string;
  tabList: TabListItem[];
};

export const createNewSheet = async ({
  collabName,
  collabEmail,
  mainSpreadsheetId,
  trameId,
  folderId,
  tabList,
}: CreateNewSheetProps) => {
  console.log("createNewSheet", collabName);

  const driveApp = appDrive();

  // créer copie trame
  const trameCopy = await driveApp.files.copy({
    fileId: trameId,
    fields: "*",
    requestBody: {},
  });

  const fileId = trameCopy.data.id;
  if (fileId) {
    await driveApp.files.update({
      fileId,
      addParents: folderId,
      requestBody: { name: collabName },
      fields: "*",
    });
    // mise à jour des datas
    await updateWholeDatas({
      tabList,
      collabName,
      collabFileId: fileId,
      mainSpreadsheetId,
      forceContratUpdate: true,
    });

    const paramsData = await sheetAPI.getTabData(
      mainSpreadsheetId,
      tabList,
      TAB_NAME_PARAMETRES
    );

    // donner accès
    await driveApp.permissions.create({
      fileId,
      requestBody: { role: "writer", type: "user", emailAddress: collabEmail },
      sendNotificationEmail: false,
    });

    for await (const params of paramsData) {
      await driveApp.permissions.create({
        fileId,
        requestBody: {
          role: "writer",
          type: "user",
          emailAddress: params[TAB_PARAMETRES_COL_EMAIL],
        },
        sendNotificationEmail: false,
      });
    }
  }

  return fileId || "";
};
