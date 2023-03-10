import express from "express";
import cors from "cors";
import cookieParser from "cookie-parser";
import tasks from "./tasks";
import * as dotenv from "dotenv";

dotenv.config();

const app = express();

// CRON SCHEDULED TASK
tasks.initScheduledJobs();

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

import buildCollabRouter from "./routes/buildCollabRouter";
import importDatasRouter from "./routes/importDatasRouter";
import exportCorrectionsRouter from "./routes/exportCorrectionsRouter";
import initVariablesRouter from "./routes/initVariablesRouter";

router.use("/buildCollab", buildCollabRouter);
router.use("/importDatas", importDatasRouter);
router.use("/exportCorrections", exportCorrectionsRouter);
router.use("/init", initVariablesRouter);

app.use("/api", router);

export default app;
