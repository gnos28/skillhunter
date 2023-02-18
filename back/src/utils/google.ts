import { google } from "googleapis";

const getAuth = (scopes: string[]) =>
  new google.auth.GoogleAuth({
    keyFile: "./auth.json",
    scopes,
  });

export const appCalendar = () => {
  const scopes = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/calendar.events",
  ];

  const auth = getAuth(scopes);

  const agenda = google.calendar({
    version: "v3",
    auth,
  });

  return agenda;
};

export const appDrive = () => {
  const scopes = ["https://www.googleapis.com/auth/drive"];
  const auth = getAuth(scopes);

  const drive = google.drive({
    version: "v3",
    auth,
  });

  return drive;
};

export const appSheet = () => {
  const scopes = ["https://www.googleapis.com/auth/spreadsheets"];
  const auth = getAuth(scopes);

  const sheets = google.sheets({
    version: "v4",
    auth,
  });

  return sheets;
};

export const appGmail = () => {
  const scopes = ["https://www.googleapis.com/auth/gmail.send"];
  const auth = getAuth(scopes);

  const gmail = google.gmail({
    version: "v1",
    auth,
  });

  return gmail;
};
