import CronJob from "node-cron";
import { buildCollab } from "../services/buildCollabService";

type Callback = () => void;

const initScheduledJobs = () => {
  if (process.env.NODE_ENV !== "test") console.log("initScheduledJobs...");

  const every2hour = "0 */2 * * *";
  const every30min = "*/30 * * * *";
  const every2min = "*/2 * * * *";

  const cronBuildCollab = CronJob.schedule(every30min, () => {
    console.log("⏰⏰⏰ launching scheduled task buildCollab");
    buildCollab({});
  });

  cronBuildCollab.start();
};

export default { initScheduledJobs };
