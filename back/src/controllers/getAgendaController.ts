import { Request, Response } from "express";
import { calendar_v3, google } from "googleapis";
import fs from "fs";

import * as dotenv from "dotenv";
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

  const driveApp = getDrive();

  driveApp.files.copy();
};

const getAgendaController: ControllerType = {};

getAgendaController.buildCollab = async (req, res) => {
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

export default getAgendaController;
