import { GoogleSpreadsheet } from "google-spreadsheet";
import * as dotenv from "dotenv";
import AlphanumericEncoder from "alphanumeric-encoder";

dotenv.config();

export const getSheetTabIds = async (sheetId: string) => {
  console.log("*** getSheetTabIds 👎");

  const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY } = process.env;

  if (GOOGLE_SERVICE_ACCOUNT_EMAIL && GOOGLE_PRIVATE_KEY && sheetId) {
    const doc = new GoogleSpreadsheet(sheetId);

    await doc.useServiceAccountAuth({
      // env var values are copied from service account credentials generated by google
      // see "Authentication" section in docs for more info
      client_email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: GOOGLE_PRIVATE_KEY,
    });

    await doc.loadInfo(); // loads document properties and worksheets

    return doc.sheetsByIndex.map((tab) => ({
      sheetId: tab.sheetId,
      sheetName: tab.title,
    }));
  }
  return [];
};