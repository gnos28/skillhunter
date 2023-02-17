import express from "express";
import storePaypalController from "../controllers/storePaypalController";

const router = express.Router();

router.post("/", storePaypalController.store);

export default router;
