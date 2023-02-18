import { GoogleSpreadsheet } from "google-spreadsheet";
import * as dotenv from "dotenv";

dotenv.config();

type Data = {
  id: number | string;
  [key: string]: string | number | undefined;
};

export const importSheetData = async (
  sheetId: string | undefined,
  tabId: string | undefined,
  headerRowIndex?: number
) => {
  console.log("*** importSheetData 👎");
  const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY } = process.env;

  // let rawdatas = fs.readFileSync("jobseasons-oauth.json", "utf8");

  // let datas = JSON.parse(rawdatas);

  if (GOOGLE_SERVICE_ACCOUNT_EMAIL && GOOGLE_PRIVATE_KEY && sheetId && tabId) {
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

    const rows = await sheet.getRows();

    const { headerValues } = sheet;

    const datas = rows.map((row) => {
      const newObj: { [key: string]: string } = {};

      headerValues.forEach((header) => (newObj[header] = row[header]));

      return newObj;
    });

    return datas;
  }
  return [];
};
