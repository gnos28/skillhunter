import { ControllerType } from "../interfaces";
import { storeBodyToFs } from "../services/storeBodyToFs";

const initVariablesController: ControllerType = {};

initVariablesController.init = async (req, res) => {
  try {
    const { mainSpreadsheetId, folderId, trameId } = req.body;

    await storeBodyToFs({ mainSpreadsheetId, folderId, trameId });

    res.send("processing request");
  } catch (err: unknown) {
    console.error(err);
    res.sendStatus(500);
  }
};

export default initVariablesController;
