import express from "express";
import cors from "cors";
import cookieParser from "cookie-parser";
import * as dotenv from "dotenv";

dotenv.config();

const app = express();

app.use(cookieParser());

app.use(
  cors({
    // origin: process.env.FRONTEND_URL ?? "http://localhost:3000",
    origin: "*",
    credentials: false,
    optionsSuccessStatus: 200,
  })
);

app.use(express.json());

// CRUD API routes
const router = express.Router();

import getAgenda from "./routes/getAgendaRouter";
import storePaypal from "./routes/storePaypalRouter";
import getPrestations from "./routes/getPrestationsRouter";

router.use("/getAgenda", getAgenda);
router.use("/getPrestations", getPrestations);
router.use("/storePaypal", storePaypal);

app.use("/api", router);

export default app;
