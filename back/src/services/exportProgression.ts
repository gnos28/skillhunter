import { TAB_NAME_SERVER_STATUS } from "../interfaces/const";
import { sheetAPI } from "../utils/sheetAPI";

let actionName = "";
let nbIncrement: undefined | number = undefined;
let incrementState = 0;
let spreadsheetId = "";

type InitProps = {
  spreadsheetId: string;
  actionName: string;
  nbIncrement: number;
};

type UpdateNbIncrementProps = {
  actionName: string;
  nbIncrement: number;
};

type IncrementProps = {
  actionName: string;
};

//TAB_NAME_SERVER_STATUS

export const exportProgression = {
  init: async ({
    spreadsheetId: argSpreadsheetId,
    actionName: argActionName,
    nbIncrement: argNbIncrement,
  }: InitProps) => {
    spreadsheetId = argSpreadsheetId;

    actionName = argActionName;
    nbIncrement = argNbIncrement || 1;
    incrementState = 0;

    const dataToStore = [[actionName, 0, new Date().toLocaleString("fr-FR")]];

    await sheetAPI.updateRange({
      sheetId: spreadsheetId,
      tabName: TAB_NAME_SERVER_STATUS,
      startCoords: [2, 1],
      data: dataToStore,
    });
  },
  updateNbIncrement: ({
    actionName: argActionName,
    nbIncrement: argNbIncrement,
  }: UpdateNbIncrementProps) => {
    actionName = argActionName;
    nbIncrement = argNbIncrement || 1;
  },
  increment: async ({ actionName: argActionName }: IncrementProps) => {
    incrementState++;

    if (nbIncrement) {
      const dataToStore = [
        [argActionName, incrementState / nbIncrement, new Date().toLocaleString("fr-FR")],
      ];

      await sheetAPI.updateRange({
        sheetId: spreadsheetId,
        tabName: TAB_NAME_SERVER_STATUS,
        startCoords: [2, 1],
        data: dataToStore,
      });
    }
  },
};
