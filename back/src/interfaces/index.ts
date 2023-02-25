import { Request, Response } from "express";

export type ControllerType = {
  [key: string]: (req: Request, res: Response) => Promise<void>;
};

export type BaseRow = {
  [key: string]: string;
};

export type ExtRow = BaseRow & {
  rowIndex: number;
  a1Range: string;
};

export type TabListItem = {
  sheetId: string;
  sheetName: string;
};
