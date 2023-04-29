import express from "express";
import buildCollabController from "../controllers/buildCollabController";

const router = express.Router();

router.post("/", buildCollabController.buildCollab);
router.get("/", buildCollabController.buildCollabNoParams);

export default router;
