import express from "express";
import importDatasController from "../controllers/importDatasController";

const router = express.Router();

router.post("/", importDatasController.importDatas);

export default router;
