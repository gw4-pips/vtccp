import { Router, type IRouter } from "express";
import { readFile, access } from "node:fs/promises";
import { watch, type FSWatcher } from "node:fs";
import path from "node:path";

const router: IRouter = Router();

const REPO_ROOT = path.resolve(process.cwd(), "..", "..");
const FAILED_FILE = path.join(REPO_ROOT, ".github-push-failed");
const STATUS_FILE = path.join(REPO_ROOT, ".github-sync-status");

interface PushStatusData {
  failed: boolean;
  failedAt: string | null;
  failedMessage: string | null;
  lastStatusLine: string | null;
}

async function readPushStatus(): Promise<PushStatusData> {
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

  return { failed, failedAt, failedMessage, lastStatusLine };
}

router.get("/push-status", async (_req, res) => {
  try {
    const data = await readPushStatus();
    res.json(data);
  } catch {
    res.status(500).json({ error: "Failed to read push status" });
  }
});

router.get("/push-status/stream", async (req, res) => {
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");
  res.flushHeaders();

  let debounceTimer: ReturnType<typeof setTimeout> | null = null;
  const watchers: FSWatcher[] = [];
  let closed = false;

  async function sendStatus() {
    if (closed) return;
    try {
      const data = await readPushStatus();
      res.write(`data: ${JSON.stringify(data)}\n\n`);
    } catch {
      // swallow — client will reconnect
    }
  }

  function scheduleUpdate() {
    if (debounceTimer) clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => { void sendStatus(); }, 150);
  }

  function tryWatch(filePath: string) {
    try {
      const w = watch(filePath, () => scheduleUpdate());
      watchers.push(w);
    } catch {
      // file may not exist yet; directory watcher handles creation
    }
  }

  // Watch the directory so we catch file creation and deletion too
  try {
    const dirWatcher = watch(REPO_ROOT, (_event, filename) => {
      const base = filename ?? "";
      if (
        base === path.basename(FAILED_FILE) ||
        base === path.basename(STATUS_FILE)
      ) {
        scheduleUpdate();
      }
    });
    watchers.push(dirWatcher);
  } catch {
    // best-effort
  }

  // Also watch the files directly for in-place modifications
  tryWatch(FAILED_FILE);
  tryWatch(STATUS_FILE);

  // Heartbeat every 20 s keeps the connection alive through proxies
  const heartbeatId = setInterval(() => {
    if (!closed) res.write(": heartbeat\n\n");
  }, 20_000);

  // Send initial status right away
  await sendStatus();

  req.on("close", () => {
    closed = true;
    clearInterval(heartbeatId);
    if (debounceTimer) clearTimeout(debounceTimer);
    for (const w of watchers) {
      try { w.close(); } catch { /* ignore */ }
    }
  });
});

export default router;
