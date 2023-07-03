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
import { appDrive } from "../utils/google";
import { mapTrimObj } from "../utils/mapTrimObj";
import { MaxAwaitingTimeError, sheetAPI } from "../utils/sheetAPI";
import { exportProgression } from "./exportProgression";
import { lockContrat } from "./lockContratService";
import { resetCollab } from "./resetCollabService";

type ImportDatasProps = {
  emailAlert: boolean;
  mainSpreadsheetId: string;
  tabList?: TabListItem[];
  initProgress?: boolean;
};

let importDatasRunning = false;
let collabId: string = "";

export const importDatas = async ({
  emailAlert = true,
  mainSpreadsheetId,
  tabList: argTabList,
  initProgress = false,
}: ImportDatasProps) => {
  if (importDatasRunning === true) {
    console.log("💥 importDatas already running 💥");
    return;
  }

  try {
    importDatasRunning = true;

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
      collabId = collabLine[TAB_COLLAB_COL_SHEET_ID];

      let nbContratsFound = 0;

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
              nbContratsFound++;
              // console.log("contrat found !", id);
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
                // console.log("update data line ", allContratLineIndex);

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
                  // const gMailApp = appGmail();

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

      // run lockContratBatch ici
      // run batch
      if (nbContratsFound > 0) {
        console.log("nbContratsFound > 0 🥶🥶🥶");

        const sheetId = parseInt(
          (await sheetAPI.getTabIds(collabId)).filter(
            (tabId) => tabId.sheetName === TAB_NAME_CONTRATS
          )[0].sheetId,
          10
        );

        const protectedRangeIds = await sheetAPI.getProtectedRangeIds({
          spreadsheetId: collabId,
          sheetId,
        });

        await sheetAPI.deleteProtectedRange({
          spreadsheetId: collabId,
          protectedRangeIds,
        });

        sheetAPI.logBatchProtectedRange();
        await sheetAPI.runBatchProtectedRange(collabId); // volontary missing await here
      }
    }
    importDatasRunning = false;
    console.log("****** END OF importDatas FUNCTION ******");
  } catch (error: any) {
    console.log(
      "catch inside importDatas",
      error instanceof MaxAwaitingTimeError
    );

    if (error instanceof MaxAwaitingTimeError) {
      console.log("collabId", collabId);
      await resetCollab(collabId);
    }
    console.error(error.message);
    importDatasRunning = false;
  }
};
