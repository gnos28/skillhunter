import express from "express";
import exportCorrectionsController from "../controllers/exportCorrectionsController";

const router = express.Router();

router.post("/", exportCorrectionsController.exportCorrections);

export default router;
