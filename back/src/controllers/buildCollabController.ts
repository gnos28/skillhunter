import { Request, Response } from "express";
import { calendar_v3, google } from "googleapis";
import fs from "fs";

import * as dotenv from "dotenv";
import { importSheetData } from "../utils/importSheetData";
import { getSheetTabIds } from "../utils/getSheetTabIds";
import AlphanumericEncoder from "alphanumeric-encoder";
import { clearSheetRows } from "../utils/clearSheetRows";
dotenv.config();

export type ControllerType = {
  [key: string]: (req: Request, res: Response) => Promise<void>;
};

type BuildCollabProps = {
  mainSpreadsheetId: string | undefined;
  folderId: string | undefined;
  trameId: string | undefined;
};

const getAuth = () =>
  new google.auth.GoogleAuth({
    keyFile: "./auth.json",
    scopes: ["https://www.googleapis.com/auth/drive"],
  });

const getDrive = () => {
  const auth = getAuth();

  const drive = google.drive({
    version: "v3",
    auth,
  });

  return drive;
};

const getSheet = () => {
  const auth = getAuth();

  const sheets = google.sheets({
    version: "v4",
    auth,
  });

  return sheets;
};

const TAB_NAME_CONTRATS = "CONTRATS";
const TAB_NAME_VERSEMENT = "VERSEMENTS VARIABLE";
const TAB_NAME_CLIENTS = "CLIENTS";
const TAB_NAME_VARIABLE = "VARIABLE / COLLABORATEURS";
const TAB_NAME_COLLAB = "COLLABORATEURS";
const TAB_IMPORT_DATA = "IMPORT_DATAS";

type TabListItem = {
  sheetId: string;
  sheetName: string;
};

type TabCache = {
  [key: string]: {
    [key: string]: string;
  }[];
};

const tabCache: TabCache = {};

const getTabData = async (
  sheetId: string,
  tabList: TabListItem[],
  tabName: string,
  headerRowIndex?: number
) => {
  const tabId = tabList.filter((tab) => tab.sheetName === tabName)[0]?.sheetId;
  if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

  const cacheKey = sheetId + ":" + tabId;

  if (tabCache[cacheKey] === undefined)
    tabCache[cacheKey] = await importSheetData(sheetId, tabId, headerRowIndex);

  return tabCache[cacheKey];
};

const clearTabData = async (
  sheetId: string,
  tabList: TabListItem[],
  tabName: string,
  headerRowIndex?: number
) => {
  const tabId = tabList.filter((tab) => tab.sheetName === tabName)[0]?.sheetId;
  if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

  return await clearSheetRows(sheetId, tabId, headerRowIndex);
};

type UpdateWholeDatasProps = {
  tabList: TabListItem[];
  collabName: string;
  collabFileId: string;
  mainSpreadsheetId: string;
  forceContratUpdate?: boolean;
};

const contratColsToKeep = [
  "ID CONTRAT",
  "RÉALISÉ PAR",
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

const versementColsToKeep = [
  "NOM PRENOM COLLABORATEUR",
  "DATE",
  "MONTANT VERSE",
  "AJOUT CONTRAT",
  "ALL CONTRATS",
];

const variableColsToKeep = [
  "NOM PRENOM",
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

const collabColsToKeep = ["NOM PRENOM", "CONTRAT", "EMAIL"];

const clientColsToKeep = ["NOM CLIENT", "NB WEEKS GARANTIE"];

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

type UpdateSheetRange = {
  sheetId: string;
  tabName: string;
  startCoords: [number, number];
  data: any[][];
};

const updateSheetRange = async ({
  sheetId,
  tabName,
  startCoords,
  data,
}: UpdateSheetRange) => {
  const sheetApp = getSheet();

  const encoder = new AlphanumericEncoder();

  const encodedStartCol = encoder.encode(startCoords[1] || 1);
  const encodedEndCol = encoder.encode(
    (startCoords[1] || 1) - 1 + data[0].length
  );

  const rangeA1notation = `'${tabName}'!${encodedStartCol}${
    startCoords[0] || 1
  }:${encodedEndCol}${data.length + (startCoords[0] || 1) - 1}`;

  await sheetApp.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: rangeA1notation,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: data,
    },
  });
};

const handleContratUpdate = async ({
  collabFileId,
  newFileTabList,
  contratData,
  collabName,
}: HandleContratUpdateProps) => {
  const sheetApp = getSheet();

  // onglet "CONTRATS" dans copie de la trame
  const contratsByCollabValues = await getTabData(
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

    const contratsCollabIndex = contratColsToKeep.findIndex(
      (val) => val === "RÉALISÉ PAR"
    );

    const contratsIdIndex = contratColsToKeep.findIndex(
      (val) => val === "ID CONTRAT"
    );

    // pré-remplir colonne "réalisé par"
    for (let i = 0; i < fullValues.length; i++)
      fullValues[i][contratsCollabIndex] = collabName;

    console.log("contratsByCollabValues", contratsByCollabValues);

    // générer les IDs de contrat
    const idList: string[] = [];
    contratsByCollabValues.forEach((line) => {
      if (line["ID CONTRAT"]) idList.push(line[contratsIdIndex]);
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
    const encoder = new AlphanumericEncoder();
    const encodedCol = encoder.encode(fullValues[0].length);

    const rangeA1notation = `'${TAB_NAME_CONTRATS}'!A2:${encodedCol}${
      fullValues.length + 1
    }`;

    // effacer les précédentes données ?

    await sheetApp.spreadsheets.values.update({
      spreadsheetId: collabFileId,
      range: rangeA1notation,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: fullValues,
      },
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
  ...(await getTabData(mainSpreadsheetId, tabList, tabName, headerRowIndex))
    .filter((row) => (filterByCol ? row[filterByCol] === collabName : true))
    .map((row) => mapTrimObj(row, colToKeep)),
  mapTrimObj({}, colToKeep),
];

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
    colToKeep: contratColsToKeep,
    tabName: TAB_NAME_CONTRATS,
    filterByCol: "RÉALISÉ PAR",
    headerRowIndex: 2,
  });

  const versementData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    collabName,
    colToKeep: versementColsToKeep,
    tabName: TAB_NAME_VERSEMENT,
    filterByCol: "NOM PRENOM COLLABORATEUR",
  });

  const clientsData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    collabName,
    colToKeep: clientColsToKeep,
    tabName: TAB_NAME_CLIENTS,
  });

  const variableData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    collabName,
    colToKeep: variableColsToKeep,
    tabName: TAB_NAME_VARIABLE,
    filterByCol: "NOM PRENOM",
    headerRowIndex: 2,
  });

  const collabData = await buildTabData({
    mainSpreadsheetId,
    tabList,
    collabName,
    colToKeep: collabColsToKeep,
    tabName: TAB_NAME_COLLAB,
    filterByCol: "NOM PRENOM",
  });

  const newFileTabList = await getSheetTabIds(collabFileId);

  //   const importData = await getTabData(
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
  const sheetApp = getSheet();

  const encoder = new AlphanumericEncoder();
  const encodedCol = encoder.encode(values[0].length);

  const rangeA1notation = `'${TAB_IMPORT_DATA}'!A3:${encodedCol}${
    values.length + 2
  }`;

  await sheetApp.spreadsheets.values.update({
    spreadsheetId: collabFileId,
    range: rangeA1notation,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values,
    },
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

  const driveApp = getDrive();

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

  const today = new Date();

  const driveApp = getDrive();

  const tabList = await getSheetTabIds(mainSpreadsheetId);

  const collabData = await getTabData(
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
          //   console.log({ collabName, collabEmail, collabId });
          const fileInfo = await driveApp.files.get({
            fileId: collabId,
            fields: "*",
          });

          const isTrashed = fileInfo.data.trashed;

          //   console.log(collabName, "fileInfo", fileInfo.data.trashed);

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
