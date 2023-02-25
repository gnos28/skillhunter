import { ControllerType } from "../interfaces";
import { buildCollab } from "../services/buildCollabService";
import { storeBodyToFs } from "../services/storeBodyToFs";

const buildCollabController: ControllerType = {};

buildCollabController.buildCollab = async (req, res) => {
  try {
    const { mainSpreadsheetId, folderId, trameId } = req.body;

    await storeBodyToFs({ mainSpreadsheetId, folderId, trameId });

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

export default buildCollabController;
