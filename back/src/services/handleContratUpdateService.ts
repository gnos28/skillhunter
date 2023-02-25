import { BaseRow } from "../interfaces";
import {
  TAB_NAME_CONTRATS,
  COL2KEEP_CONTRATS,
  TAB_CONTRATS_COL_COLLAB,
  TAB_CONTRATS_COL_ID,
} from "../interfaces/const";
import { sheetAPI } from "../utils/sheetAPI";

type HandleContratUpdateProps = {
  collabFileId: string;
  newFileTabList: {
    sheetId: string;
    sheetName: string;
  }[];
  contratData: BaseRow[];
  collabName: string;
};

export const handleContratUpdate = async ({
  collabFileId,
  newFileTabList,
  contratData,
  collabName,
}: HandleContratUpdateProps) => {
  // retirer datavalidation

  // const sheetApp = appSheet();

  // const test = await sheetApp.spreadsheets.get({
  //   spreadsheetId: collabFileId,
  //   ranges: ["'CONTRATS'!C:L"],
  //   // includeGridData: true,
  //   fields:
  //     "sheets(data/rowData/values/dataValidation,properties(sheetId,title))",
  // });

  // const rowData = test.data.sheets
  //   ?.filter((sheet) => sheet.properties?.title === "CONTRATS")[0]
  //   .data?.map((data) => data.rowData);

  // const rowVals =
  //   rowData && rowData.map((row) => row && row.map((val) => val.values));

  // const dataValidations =
  //   rowVals &&
  //   rowVals
  //     .map(
  //       (cellData) =>
  //         cellData &&
  //         cellData
  //           .filter((dv) => dv !== undefined)
  //           .map((cell) => cell && cell.map((c) => c && c.dataValidation))
  //           .flat()
  //           .filter((dv) => dv !== undefined)
  //     )
  //     .flat()
  //     .filter((dv) => dv !== undefined);

  // console.log("dataValidation", dataValidation);

  // console.log("test", JSON.stringify(test));

  // onglet "CONTRATS" dans copie de la trame
  const contratsByCollabValues = await sheetAPI.getTabData(
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

  const spreadsheetsData = await sheetAPI.getTabMetaData({
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

    await sheetAPI.updateRange({
      sheetId: collabFileId,
      tabName: TAB_NAME_CONTRATS,
      startCoords: [2, 1],
      data: fullValues,
    });
  }
};
