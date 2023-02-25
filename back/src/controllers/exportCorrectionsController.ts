import { ControllerType } from "../interfaces";
import { exportCorrections } from "../services/exportCorrectionsService";

const exportCorrectionsController: ControllerType = {};

exportCorrectionsController.exportCorrections = async (req, res) => {
  try {
    const { mainSpreadsheetId } = req.body;

    exportCorrections({
      mainSpreadsheetId,
    });

    res.send("processing request");
  } catch (err: unknown) {
    console.error(err);
    res.sendStatus(500);
  }
};

export default exportCorrectionsController;
