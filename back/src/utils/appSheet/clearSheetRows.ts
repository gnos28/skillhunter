import { GoogleSpreadsheet } from "google-spreadsheet";
import * as dotenv from "dotenv";
import { TabListItem } from "../../interfaces";
dotenv.config();

export const clearSheetRows = async (
  sheetId: string,
  tabId: string,
  headerRowIndex?: number
) => {
  console.log("clearSheetRows");

  const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY } = process.env;

  if (GOOGLE_SERVICE_ACCOUNT_EMAIL && GOOGLE_PRIVATE_KEY) {
    const doc = new GoogleSpreadsheet(sheetId);

    await doc.useServiceAccountAuth({
      // env var values are copied from service account credentials generated by google
      // see "Authentication" section in docs for more info
      client_email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: GOOGLE_PRIVATE_KEY,
    });

    await doc.loadInfo(); // loads document properties and worksheets

    const sheet = doc.sheetsById[tabId];

    if (headerRowIndex) sheet.loadHeaderRow(headerRowIndex);

    await sheet.clearRows();
  }
};

export const clearTabData = async (
  sheetId: string,
  tabList: TabListItem[],
  tabName: string,
  headerRowIndex?: number
) => {
  const tabId = tabList.filter((tab) => tab.sheetName === tabName)[0]?.sheetId;
  if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

  return await clearSheetRows(sheetId, tabId, headerRowIndex);
};