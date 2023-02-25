import { ExtRow } from "../interfaces";
import {
  TAB_NAME_CONTRATS,
  TAB_CONTRATS_COL_DATE_DEBUT,
  TAB_CONTRATS_COL_NB_SEMAINE_GARANTIE,
  TAB_CONTRATS_COL_CLIENT,
  TAB_CONTRATS_COL_PERCENT,
  TAB_CONTRATS_COL_DATE_FIN_GARANTIE,
  TAB_CONTRATS_COL_RUPTURE,
} from "../interfaces/const";
import { sheetAPI } from "../utils/sheetAPI";

type LockContratProps = {
  collabSheetId: string;
  users: string[];
  collabTabList: {
    sheetId: string;
    sheetName: string;
  }[];
  contratLine: ExtRow;
};

export const lockContrat = async ({
  collabSheetId,
  users,
  collabTabList,
  contratLine,
}: LockContratProps) => {
  // const sheetApp = appSheet();
  const today = new Date();

  const sheetId = parseInt(
    collabTabList.filter((tab) => tab.sheetName === TAB_NAME_CONTRATS)[0]
      .sheetId,
    10
  );

  // console.log("contratLine.rowIndex", contratLine.rowIndex);
  const rowIndex = contratLine.rowIndex - 1;

  const contratLineKeys = Object.keys(contratLine);

  const dateDebutIndex = contratLineKeys.findIndex(
    (key) => key === TAB_CONTRATS_COL_DATE_DEBUT
  );
  const garantieIndex = contratLineKeys.findIndex(
    (key) => key === TAB_CONTRATS_COL_NB_SEMAINE_GARANTIE
  );
  const clientIndex = contratLineKeys.findIndex(
    (key) => key === TAB_CONTRATS_COL_CLIENT
  );
  const percentIndex = contratLineKeys.findIndex(
    (key) => key === TAB_CONTRATS_COL_PERCENT
  );

  sheetAPI.addBatchProtectedRange({
    spreadsheetId: collabSheetId,
    editors: users,
    namedRangeId: `lock-${rowIndex}-CD`,
    sheetId,
    startColumnIndex: dateDebutIndex,
    startRowIndex: rowIndex,
    endColumnIndex: garantieIndex,
    endRowIndex: rowIndex,
  });

  sheetAPI.addBatchProtectedRange({
    spreadsheetId: collabSheetId,
    editors: users,
    namedRangeId: `lock-${rowIndex}-GL`,
    sheetId,
    startColumnIndex: clientIndex,
    startRowIndex: rowIndex,
    endColumnIndex: percentIndex,
    endRowIndex: rowIndex,
  });

  const garantieDate = contratLine[TAB_CONTRATS_COL_DATE_FIN_GARANTIE];
  const rupture = contratLine[TAB_CONTRATS_COL_RUPTURE];

  if (
    rupture ||
    (garantieDate && today.getTime() > new Date(garantieDate).getTime())
  ) {
    // bloquer les cellules "rupture" dont la date de garantie est dépassée

    // console.log("locking rupture date of ", rowIndex);

    const ruptureIndex = contratLineKeys.findIndex(
      (key) => key === TAB_CONTRATS_COL_RUPTURE
    );

    sheetAPI.addBatchProtectedRange({
      spreadsheetId: collabSheetId,
      editors: users,
      namedRangeId: `lock-${rowIndex}-E`,
      sheetId,
      startColumnIndex: ruptureIndex,
      startRowIndex: rowIndex,
      endColumnIndex: ruptureIndex,
      endRowIndex: rowIndex,
    });
  }

  // run batch
  await sheetAPI.runBatchProtectedRange(collabSheetId);
};
