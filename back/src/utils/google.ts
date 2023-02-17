import { google } from "googleapis";

const getAuth = () =>
  new google.auth.GoogleAuth({
    keyFile: "./auth.json",
    scopes: [
      "https://www.googleapis.com/auth/calendar",
      "https://www.googleapis.com/auth/calendar.events",
    ],
  });

export const getAgenda = () => {
  const auth = getAuth();

  const agenda = google.calendar({
    version: "v3",
    auth,
  });

  return agenda;
};
