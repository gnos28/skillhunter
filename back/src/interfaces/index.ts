import { Request, Response } from "express";

export type ControllerType = {
  [key: string]: (req: Request, res: Response) => Promise<void>;
};
