import { ControllerType } from "../interfaces";
import { importDatas } from "../services/importDatasService";

const importDatasController: ControllerType = {};

importDatasController.importDatas = async (req, res) => {
  try {
    const { emailAlert, mainSpreadsheetId } = req.body;

    const buildResult = await importDatas({ emailAlert, mainSpreadsheetId });
    res.send(buildResult);
  } catch (err: unknown) {
    console.error(err);
    res.sendStatus(500);
  }
};

export default importDatasController;
