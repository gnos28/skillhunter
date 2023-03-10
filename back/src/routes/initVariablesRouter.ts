import express from "express";
import initVariablesController from "../controllers/initVariablesController";

const router = express.Router();

router.post("/", initVariablesController.init);

export default router;
