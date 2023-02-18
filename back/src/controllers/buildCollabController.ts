import * as dotenv from "dotenv";
dotenv.config();
import { getSheetTabIds } from "../utils/getSheetTabIds";
import { clearTabData, TabListItem } from "../utils/clearSheetRows";
import { appDrive, appSheet } from "../utils/google";
import { updateSheetRange } from "../utils/updateSheetRange";
import { tabData } from "../utils/tabData";
import { ControllerType } from "../interfaces";

const TAB_NAME_CONTRATS = "CONTRATS";
const TAB_NAME_VERSEMENT = "VERSEMENTS VARIABLE";
const TAB_NAME_CLIENTS = "CLIENTS";
const TAB_NAME_VARIABLE = "VARIABLE / COLLABORATEURS";
const TAB_NAME_COLLAB = "COLLABORATEURS";
const TAB_IMPORT_DATA = "IMPORT_DATAS";

const TAB_CONTRATS_COL_ID = "ID CONTRAT";
const TAB_CONTRATS_COL_COLLAB = "RÉALISÉ PAR";

const COL2KEEP_CONTRATS = [
  TAB_CONTRATS_COL_ID,
  TAB_CONTRATS_COL_COLLAB,
  "DATE DEBUT",
  "NB SEMAINES GARANTIE",
  "(RUPTURE GARANTIE)",
  "DATE PAIEMENT CLIENT",
  "CLIENT",
  "TYPE CONTRAT",
  "CANDIDAT",
  "DESCRIPTION",
  "SALAIRE CANDIDAT",
  "% CONTRAT",
];

const TAB_VERSEMENT_COL_COLLAB = "NOM PRENOM COLLABORATEUR";

const COL2KEEP_VERSEMENT = [
  TAB_VERSEMENT_COL_COLLAB,
  "DATE",
  "MONTANT VERSE",
  "AJOUT CONTRAT",
  "ALL CONTRATS",
];

const TAB_VARIABLE_COL_COLLAB = "NOM PRENOM";

const COL2KEEP_VARIABLE = [
  TAB_VARIABLE_COL_COLLAB,
  "DEBUT",
  "FIN",
  "FREQ VARIABLE",
  "T1 MINI",
  "T1 %",
  "T2 MINI",
  "T2 %",
  "T3 MINI",
  "T3 %",
  "T4 MINI",
  "T4 %",
];

const TAB_COLLAB_COL_COLLAB = "NOM PRENOM";

const COL2KEEP_COLLAB = [TAB_COLLAB_COL_COLLAB, "CONTRAT", "EMAIL"];

const COL2KEEP_CLIENTS = ["NOM CLIENT", "NB WEEKS GARANTIE"];

const mapTrimObj = (
  row: {
    [key: string]: string;
  },
  colsToKeep: string[]
) => {
  const filteredRow: {
    [key: string]: string;
  } = {};

  colsToKeep.forEach(
    (colName) =>
      (filteredRow[colName] = Object.keys(row).includes(colName)
        ? row[colName]
        : "")
  );

  return filteredRow;
};

type HandleContratUpdateProps = {
  collabFileId: string;
  newFileTabList: {
    sheetId: string;
    sheetName: string;
  }[];
  contratData: {
    [key: string]: string;
  }[];
  collabName: string;
};

const handleContratUpdate = async ({
  collabFileId,
  newFileTabList,
  contratData,
  collabName,
}: HandleContratUpdateProps) => {
  const sheetApp = appSheet();

  // onglet "CONTRATS" dans copie de la trame
  const contratsByCollabValues = await tabData.get(
    collabFileId,
    newFileTabList,
    TAB_NAME_CONTRATS
  );

  const values = Array(contratData.length)
    .fill(undefined)
    .map((_, rowIndex) => {
      if (rowIndex > contratData.length - 1)
        return Array(Object.keys(contratData[0]).length).fill("");

      return Object.keys(contratData[rowIndex]).map(
        (key) => contratData[rowIndex][key]
      );
    });

  const spreadsheetsData = await sheetApp.spreadsheets.get({
    spreadsheetId: collabFileId,
    fields: "*",
    ranges: ["A:A"],
  });

  const sheetData = spreadsheetsData.data.sheets?.filter(
    (sheet) => sheet.properties?.title === TAB_NAME_CONTRATS
  );

  const nbRows = sheetData && sheetData[0].properties?.gridProperties?.rowCount;

  if (nbRows) {
    const fullValues = Array(nbRows - 1)
      .fill(undefined)
      .map(
        (_, rowIndex) => values[rowIndex] || Array(values[0].length).fill("")
      );

    const contratsCollabIndex = COL2KEEP_CONTRATS.findIndex(
      (val) => val === TAB_CONTRATS_COL_COLLAB
    );

    const contratsIdIndex = COL2KEEP_CONTRATS.findIndex(
      (val) => val === TAB_CONTRATS_COL_ID
    );

    // pré-remplir colonne "réalisé par"
    for (let i = 0; i < fullValues.length; i++)
      fullValues[i][contratsCollabIndex] = collabName;

    console.log("contratsByCollabValues", contratsByCollabValues);

    // générer les IDs de contrat
    const idList: string[] = [];
    contratsByCollabValues.forEach((line) => {
      if (line[TAB_CONTRATS_COL_ID]) idList.push(line[contratsIdIndex]);
    });

    console.log("idList", idList);

    for (let i = 0; i < fullValues.length; i++)
      if (!fullValues[i][contratsIdIndex]) {
        let newId = "";
        let j = i;
        while (newId === "" || idList.includes(newId)) {
          newId =
            collabName.replace(" ", "").toLowerCase() +
            "_" +
            (j + 1).toString().padStart(4, "0");
          j++;
        }
        idList.push(newId);

        fullValues[i][contratsIdIndex] = newId;
      }

    // contratsByCollabRange.setValues(contratsByCollabValues);
    // contratsByCollabRange.setDataValidations(dataValidationRules); // remettre data validation

    // effacer les précédentes données ?

    updateSheetRange({
      sheetId: collabFileId,
      tabName: TAB_NAME_CONTRATS,
      startCoords: [1, 2],
      data: fullValues,
    });
  }
};

type BuildTabDataProps = {
  mainSpreadsheetId: string;
  tabList: TabListItem[];
  tabName: string;
  collabName: string;
  colToKeep: string[];
  filterByCol?: string;
  headerRowIndex?: number;
};

const buildTabData = async ({
  mainSpreadsheetId,
  tabList,
  tabName,
  collabName,
  filterByCol,
  headerRowIndex,
  colToKeep,
}: BuildTabDataProps) => [
  ...(await tabData.get(mainSpreadsheetId, tabList, tabName, headerRowIndex))
    .filter((row) => (filterByCol ? row[filterByCol] === collabName : true))
    .map((row) => mapTrimObj(row, colToKeep)),
  mapTrimObj({}, colToKeep),
];

type UpdateWholeDatasProps = {
  tabList: TabListItem[];
  collabName: string;
  collabFileId: string;
  mainSpreadsheetId: string;
  forceContratUpdate?: boolean;
};

const updateWholeDatas = async ({
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

  const newFileTabList = await getSheetTabIds(collabFileId);

  //   const importData = await tabData.get(
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
  await clearTabData(collabFileId, newFileTabList, TAB_IMPORT_DATA, 2);

  // udpate data in IMPORT_DATAS sheet
  updateSheetRange({
    sheetId: collabFileId,
    tabName: TAB_IMPORT_DATA,
    startCoords: [1, 3],
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

type CreateNewSheetProps = {
  collabName: string;
  collabEmail: string;
  mainSpreadsheetId: string;
  trameId: string;
  folderId: string;
  tabList: TabListItem[];
};

const createNewSheet = async ({
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

    // donner accès
    await driveApp.permissions.create({
      fileId,
      requestBody: { role: "writer", type: "user", emailAddress: collabEmail },
    });
  }

  return fileId || "";
};

type BuildCollabProps = {
  mainSpreadsheetId: string | undefined;
  folderId: string | undefined;
  trameId: string | undefined;
};

const buildCollab = async ({
  mainSpreadsheetId,
  folderId,
  trameId,
}: BuildCollabProps) => {
  if (
    mainSpreadsheetId === undefined ||
    folderId === undefined ||
    trameId === undefined
  )
    throw new Error("missing id");

  // clear cache
  tabData.clearCache();

  const today = new Date();

  const driveApp = appDrive();

  const tabList = await getSheetTabIds(mainSpreadsheetId);

  const collabData = await tabData.get(
    mainSpreadsheetId,
    tabList,
    TAB_NAME_COLLAB
  );

  // vérifier existence du fichier collaborateur
  let nbCreatedFiles = 0;

  let lineIndex = 2;
  for await (const line of collabData) {
    const collabName = line["NOM PRENOM"];
    const collabEmail = line["EMAIL"];
    let collabId = line["SHEET ID"];

    if (collabName && collabEmail) {
      let sheetFound = false;
      let collabSheet = null;

      if (collabId) {
        try {
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
        } catch {
          console.log(`sheet ${collabName} not found 😱`);
        }
      }

      if (!sheetFound) {
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

        updateSheetRange({
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
    }
    lineIndex++;
  }

  //   importDatas(true)

  return collabData;
};

const buildCollabController: ControllerType = {};

buildCollabController.buildCollab = async (req, res) => {
  try {
    const { mainSpreadsheetId, folderId, trameId } = req.body;

    const buildResult = await buildCollab({
      mainSpreadsheetId,
      folderId,
      trameId,
    });
    res.send(buildResult);
  } catch (err: unknown) {
    console.error(err);
    res.sendStatus(500);
  }
};

export default buildCollabController;
