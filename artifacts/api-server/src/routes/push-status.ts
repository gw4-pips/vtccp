import { Router, type IRouter } from "express";
import { readFile, access } from "node:fs/promises";
import path from "node:path";

const router: IRouter = Router();

const REPO_ROOT = path.resolve(process.cwd(), "..", "..");
const FAILED_FILE = path.join(REPO_ROOT, ".github-push-failed");
const STATUS_FILE = path.join(REPO_ROOT, ".github-sync-status");

router.get("/push-status", async (_req, res) => {
  try {
    let failed = false;
    let failedAt: string | null = null;
    let failedMessage: string | null = null;
    let lastStatusLine: string | null = null;

    try {
      await access(FAILED_FILE);
      const raw = await readFile(FAILED_FILE, "utf8");
      failed = true;
      const lines = raw.trim().split("\n");
      for (const line of lines) {
        if (line.startsWith("timestamp=")) failedAt = line.slice("timestamp=".length);
        if (line.startsWith("message=")) failedMessage = line.slice("message=".length);
      }
    } catch {
      failed = false;
    }

    try {
      const statusRaw = await readFile(STATUS_FILE, "utf8");
      const statusLines = statusRaw.trim().split("\n").filter(Boolean);
      if (statusLines.length > 0) {
        lastStatusLine = statusLines[statusLines.length - 1];
      }
    } catch {
    }

    res.json({
      failed,
      failedAt,
      failedMessage,
      lastStatusLine,
    });
  } catch (err) {
    res.status(500).json({ error: "Failed to read push status" });
  }
});

export default router;
