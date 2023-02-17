import express from "express";
import getAgendaController from "../controllers/getAgendaController";

const router = express.Router();

router.get("/", getAgendaController.getAll);
router.post("/", getAgendaController.getByDate);

export default router;
