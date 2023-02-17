import { Request, Response } from "express";
import { importSheet } from "../utils/importSheet";

export type ControllerType = {
  [key: string]: (req: Request, res: Response) => Promise<void>;
};

const getPrestationsController: ControllerType = {};

getPrestationsController.getAll = async (req, res) => {
  try {
    const { IMPORT_PRESTATIONS_SHEET_ID } = process.env;

    const prestations = (await importSheet(IMPORT_PRESTATIONS_SHEET_ID)).map(
      (prestation) => ({
        ...prestation,
        price: parseInt(prestation.price),
        duration: parseInt(prestation.duration),
      })
    );
    res.send({ prestations });
  } catch (err: unknown) {
    console.error(err);
    res.sendStatus(500);
  }
};

export default getPrestationsController;
