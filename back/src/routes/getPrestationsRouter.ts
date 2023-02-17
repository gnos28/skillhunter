import express from "express";
import getPrestationsController from "../controllers/getPrestationsController";

const router = express.Router();

router.get("/", getPrestationsController.getAll);

export default router;
