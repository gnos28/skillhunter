import CronJob from "node-cron";
import { buildCollab } from "../services/buildCollabService";

const initScheduledJobs = () => {
  if (process.env.NODE_ENV !== "test") console.log("initScheduledJobs...");

  const every30min = "*/30 * * * *";
  const every6min = "*/2 * * * *";

  const cronBuildCollab = CronJob.schedule(every30min, () => buildCollab({}));

  cronBuildCollab.start();
};

export default { initScheduledJobs };
