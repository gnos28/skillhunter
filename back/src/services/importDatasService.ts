import { TabListItem } from "../interfaces";
import {
  TAB_NAME_COLLAB,
  TAB_NAME_CONTRATS,
  TAB_NAME_PARAMETRES,
  TAB_PARAMETRES_COL_EMAIL,
  TAB_COLLAB_COL_COLLAB,
  TAB_COLLAB_COL_EMAIL,
  TAB_COLLAB_COL_SHEET_ID,
  TAB_CONTRATS_COL_ID,
  TAB_CONTRATS_COL_DATE_DEBUT,
  TAB_CONTRATS_COL_NB_SEMAINE_GARANTIE,
  TAB_CONTRATS_COL_CLIENT,
  TAB_CONTRATS_COL_TYPE,
  TAB_CONTRATS_COL_CANDIDAT,
  TAB_CONTRATS_COL_DESCRIPTION,
  TAB_CONTRATS_COL_SALAIRE,
  TAB_CONTRATS_COL_PERCENT,
  TAB_CONTRATS_COL_IMPORT_ID,
  COL2KEEP_CONTRATS,
} from "../interfaces/const";
import { appDrive, appGmail } from "../utils/google";
import { mapTrimObj } from "../utils/mapTrimObj";
import { sheetAPI } from "../utils/sheetAPI";
import { exportProgression } from "./exportProgression";
import { lockContrat } from "./lockContratService";

type ImportDatasProps = {
  emailAlert: boolean;
  mainSpreadsheetId: string;
  tabList?: TabListItem[];
  initProgress?: boolean;
};

export const importDatas = async ({
  emailAlert = true,
  mainSpreadsheetId,
  tabList: argTabList,
  initProgress = false,
}: ImportDatasProps) => {
  console.log("****** importDatas ******");

  sheetAPI.clearCache();

  if (initProgress)
    await exportProgression.init({
      spreadsheetId: mainSpreadsheetId,
      actionName: "importDatas",
      nbIncrement: 1,
    });

  let tabList: TabListItem[] = [];
  if (!argTabList) tabList = await sheetAPI.getTabIds(mainSpreadsheetId);
  else tabList = argTabList;

  const driveApp = appDrive();

  const collabData = await sheetAPI.getTabData(
    mainSpreadsheetId,
    tabList,
    TAB_NAME_COLLAB
  );
  const allContratData = await sheetAPI.getTabData(
    mainSpreadsheetId,
    tabList,
    TAB_NAME_CONTRATS,
    2
  );

  const alertList: string[] = [];
  if (emailAlert) {
    const paramsData = await sheetAPI.getTabData(
      mainSpreadsheetId,
      tabList,
      TAB_NAME_PARAMETRES
    );

    paramsData.forEach((line) => {
      if (line[TAB_PARAMETRES_COL_EMAIL])
        alertList.push(line[TAB_PARAMETRES_COL_EMAIL]);
    });
  }

  if (initProgress) {
    const nbCollab = collabData.reduce(
      (acc, val) => (val[TAB_COLLAB_COL_SHEET_ID] ? acc + 1 : acc),
      0
    );
    exportProgression.updateNbIncrement({
      actionName: "importDatas",
      nbIncrement: nbCollab + 1,
    });
    await exportProgression.increment({
      actionName: "importDatas",
    });
  } else
    await exportProgression.increment({
      actionName: "buildCollab",
    });

  // vérifier existence du fichier collaborateur

  for await (const collabLine of collabData) {
    const collabName = collabLine[TAB_COLLAB_COL_COLLAB];
    const collabEmail = collabLine[TAB_COLLAB_COL_EMAIL];
    let collabId = collabLine[TAB_COLLAB_COL_SHEET_ID];

    if (collabName && collabEmail) {
      let sheetFound = false;
      let collabSheet = null;

      if (collabId) {
        try {
          const fileInfo = await driveApp.files.get({
            fileId: collabId,
            fields: "*",
          });

          const isTrashed = fileInfo.data.trashed;

          if (!isTrashed) {
            sheetFound = true;
            console.log(`sheet ${collabName} found 😀`);
          } else console.log(`sheet ${collabName} is trashed 🗑️`);
        } catch {
          console.log(`sheet ${collabName} not found 😱`);
        }
      }

      if (sheetFound) {
        // récupérer les infos contrats dans fichier collaborateur
        const collabTabList = await sheetAPI.getTabIds(collabId);

        const collabContratData = await sheetAPI.getTabData(
          collabId,
          collabTabList,
          TAB_NAME_CONTRATS
        );

        for await (const line of collabContratData) {
          // collabContratData.forEach(async (line, lineIndex) => {
          const id = line[TAB_CONTRATS_COL_ID];
          const dateDebut = line[TAB_CONTRATS_COL_DATE_DEBUT];
          const nbGarantyWeeks = line[TAB_CONTRATS_COL_NB_SEMAINE_GARANTIE];
          const client = line[TAB_CONTRATS_COL_CLIENT];
          const type = line[TAB_CONTRATS_COL_TYPE];
          const candidat = line[TAB_CONTRATS_COL_CANDIDAT];
          const description = line[TAB_CONTRATS_COL_DESCRIPTION];
          const salaire = line[TAB_CONTRATS_COL_SALAIRE];
          const percent = line[TAB_CONTRATS_COL_PERCENT];

          if (
            id &&
            dateDebut &&
            nbGarantyWeeks &&
            client &&
            type &&
            candidat &&
            description &&
            salaire &&
            percent
          ) {
            console.log("contrat found !", id);
            // si contrat correctement rempli
            // rechercher contrat dans fichier chapeau
            let allContratLineIndex: number | undefined = undefined;

            const filteredAllContratData = allContratData.filter(
              (contrat) => contrat[TAB_CONTRATS_COL_ID] === id
            );

            let colIndex: number | undefined = undefined;

            if (filteredAllContratData.length)
              allContratLineIndex = filteredAllContratData[0].rowIndex;

            colIndex =
              Object.keys(allContratData[0]).findIndex(
                (col) => col === TAB_CONTRATS_COL_IMPORT_ID
              ) + 1;

            const trimedLine = mapTrimObj(line, COL2KEEP_CONTRATS);

            const trimedArray = Object.values(trimedLine);

            // si trouvé > mettre à jour les datas dans chapeau
            if (allContratLineIndex) {
              console.log("update data line ", allContratLineIndex);

              await sheetAPI.updateRange({
                sheetId: mainSpreadsheetId,
                tabName: TAB_NAME_CONTRATS,
                startCoords: [allContratLineIndex, colIndex],
                data: [trimedArray],
              });
            } // si pas trouvé > ajouter nouveau contrat dans chapeau
            else {
              console.log("add new contrat");

              const allContratIndex = allContratData.map(
                (line) => line.rowIndex
              );

              let emptyLineIndex: number | false = false; // recherche premiere ligne vide [sans ID]
              let indexToCheck = 3;

              while (emptyLineIndex === false) {
                if (allContratIndex.includes(indexToCheck)) indexToCheck++;
                else emptyLineIndex = indexToCheck;
              }

              await sheetAPI.updateRange({
                sheetId: mainSpreadsheetId,
                tabName: TAB_NAME_CONTRATS,
                startCoords: [emptyLineIndex, 1],
                data: [trimedArray],
              });

              await sheetAPI.updateRange({
                sheetId: mainSpreadsheetId,
                tabName: TAB_NAME_CONTRATS,
                startCoords: [emptyLineIndex, colIndex],
                data: [trimedArray],
              });

              // envoyer un email au owner du chapeau
              if (emailAlert) {
                const gMailApp = appGmail();

                const dateDebutDate = new Date(dateDebut);

                // const data =
                //   base64url.encode(`date début ${dateDebutDate.getDate()}/${
                //     dateDebutDate.getMonth() + 1
                //   }/${dateDebutDate.getFullYear()}<br>
                // client ${client}<br>
                // type ${type}<br>
                // réalisé par ${collabName}<br>`);

                // await gMailApp.users.messages.send({
                //   userId: "me",
                //   requestBody: {
                //     payload: {
                //       body: {
                //         data,
                //       },
                //       headers: [
                //         {
                //           name: "To",
                //           value: alertList.join(","),
                //         },
                //         {
                //           name: "Subject",
                //           value: `*** NOUVEAU CONTRAT ${candidat.toUpperCase()} ***`,
                //         },
                //       ],
                //     },
                //   },
                // });
              }
            }

            await lockContrat({
              collabSheetId: collabId,
              users: alertList,
              collabTabList,
              contratLine: line,
            });
          }
        }
      }
      if (initProgress)
        await exportProgression.increment({
          actionName: "importDatas",
        });
      else
        await exportProgression.increment({
          actionName: "buildCollab",
        });
    }
  }
  console.log("****** END OF importDatas FUNCTION ******");
};
