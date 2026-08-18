import { useState, useEffect, useRef } from "react";
import scriptText from "./v137.txt?raw";

const LINE_COUNT = scriptText.split("\n").length;
const BYTE_COUNT = new Blob([scriptText]).size;

interface OverwrittenCommit {
  sha: string;
  message: string;
}

interface ParsedPushLog {
  overwrittenCommits: OverwrittenCommit[];
  commitCount: number;
  pushSha: string | null;
  raw: string;
}

function parsePushLog(log: string): ParsedPushLog {
  const lines = log.split("\n");
  const overwrittenCommits: OverwrittenCommit[] = [];
  let commitCount = 0;
  let inWarningBlock = false;
  let pushSha: string | null = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Extract Replit HEAD SHA: "==> Replit HEAD: a1b2c3d"
    const shaMatch = line.match(/Replit HEAD:\s*([0-9a-f]{6,12})/i);
    if (shaMatch) {
      pushSha = shaMatch[1];
    }

    const warningMatch = line.match(
      /WARNING:\s+GitHub has (\d+) commit\(s\) that Replit does not/
    );
    if (warningMatch) {
      commitCount = parseInt(warningMatch[1], 10);
      inWarningBlock = true;
      continue;
    }

    if (inWarningBlock) {
      // Commit lines are indented with 4 spaces: "    <sha> <message>"
      const commitMatch = line.match(/^\s{4}([0-9a-f]{6,12})\s+(.+)/);
      if (commitMatch) {
        overwrittenCommits.push({
          sha: commitMatch[1],
          message: commitMatch[2],
        });
      } else if (line.trim() === "" && overwrittenCommits.length > 0) {
        // blank line after commits — end of the commit list
        inWarningBlock = false;
      } else if (line.trim() !== "") {
        // non-commit, non-blank line ends the block too
        inWarningBlock = false;
      }
    }
  }

  return { overwrittenCommits, commitCount, pushSha, raw: log };
}

// ── Push log history (localStorage) ──────────────────────────────────────────

const PUSH_LOG_HISTORY_KEY = "push-log-history";
const PUSH_LOG_HISTORY_MAX = 10;

interface PushLogEntry {
  id: string;
  savedAt: string; // ISO date string
  raw: string;
}

function loadHistory(): PushLogEntry[] {
  try {
    const raw = localStorage.getItem(PUSH_LOG_HISTORY_KEY);
    if (!raw) return [];
    return JSON.parse(raw) as PushLogEntry[];
  } catch {
    return [];
  }
}

function persistHistory(entries: PushLogEntry[]) {
  try {
    localStorage.setItem(PUSH_LOG_HISTORY_KEY, JSON.stringify(entries));
  } catch {}
}

function upsertEntry(
  raw: string,
  existing: PushLogEntry[],
  currentId: string | null
): { entries: PushLogEntry[]; id: string } {
  // If there's a currently selected (unsaved) draft, update it in-place.
  // Otherwise create a new entry.
  if (currentId) {
    const idx = existing.findIndex((e) => e.id === currentId);
    if (idx !== -1) {
      const updated = existing.map((e, i) =>
        i === idx ? { ...e, raw, savedAt: new Date().toISOString() } : e
      );
      // Move updated entry to front
      const entry = updated.splice(idx, 1)[0];
      const next = [entry, ...updated].slice(0, PUSH_LOG_HISTORY_MAX);
      persistHistory(next);
      return { entries: next, id: entry.id };
    }
  }
  // New entry
  const entry: PushLogEntry = {
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    savedAt: new Date().toISOString(),
    raw,
  };
  const next = [entry, ...existing].slice(0, PUSH_LOG_HISTORY_MAX);
  persistHistory(next);
  return { entries: next, id: entry.id };
}

function formatEntryDate(iso: string): string {
  const d = new Date(iso);
  const now = new Date();
  const sameDay =
    d.getFullYear() === now.getFullYear() &&
    d.getMonth() === now.getMonth() &&
    d.getDate() === now.getDate();
  if (sameDay) {
    return d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
  }
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric" }) +
    " " + d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
}

function EntrySummary({ entry }: { entry: PushLogEntry }) {
  const p = entry.raw.trim() ? parsePushLog(entry.raw) : null;
  const hasWarning = p && p.overwrittenCommits.length > 0;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
      <span style={{ fontSize: 11, fontWeight: 600, color: "#e6edf3" }}>
        {formatEntryDate(entry.savedAt)}
      </span>
      <span style={{ fontSize: 10, display: "flex", alignItems: "center", gap: 4 }}>
        {hasWarning && p ? (
          <>
            <span style={{ color: "#e3b341" }}>⚠</span>
            <span style={{ color: "#9e7c26" }}>
              {p.commitCount} overwritten
            </span>
          </>
        ) : (
          <>
            <span style={{ color: "#3fb950" }}>✓</span>
            <span style={{ color: "#57ab5a" }}>clean</span>
          </>
        )}
        {p?.pushSha && (
          <span style={{ color: "#484f58", marginLeft: 2 }}>
            · {p.pushSha.slice(0, 7)}
          </span>
        )}
      </span>
    </div>
  );
}

function PushLogTab() {
  // Load history once at mount; don't call loadHistory() multiple times
  const initialHistory = loadHistory();
  const [history, setHistory] = useState<PushLogEntry[]>(initialHistory);
  // selectedId — which sidebar entry is highlighted (view mode)
  const [selectedId, setSelectedId] = useState<string | null>(
    initialHistory[0]?.id ?? null
  );
  // logText — what's currently in the textarea
  const [logText, setLogText] = useState<string>(initialHistory[0]?.raw ?? "");

  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // draftIdRef: ID of the entry currently being composed in the textarea.
  //   null  → next save must create a NEW entry (compose mode / fresh start)
  //   <id>  → next save updates this entry in-place (ongoing edit of current draft)
  // This is separate from selectedId so that loading a history entry for viewing
  // does NOT make the next paste overwrite that entry — it creates a new one.
  // All ref mutations happen synchronously in handlers, never via useEffect.
  const draftIdRef = useRef<string | null>(null);

  // logTextRef mirrors logText for use in async flush callbacks.
  // Updated synchronously in handleTextChange, not via effect, to avoid races
  // where a paste + immediate switch would still capture stale text.
  const logTextRef = useRef<string>(logText);

  const parsed = logText.trim() ? parsePushLog(logText) : null;
  const hasWarning = parsed && parsed.overwrittenCommits.length > 0;

  function cancelPendingSave() {
    if (saveTimerRef.current) {
      clearTimeout(saveTimerRef.current);
      saveTimerRef.current = null;
    }
  }

  /**
   * Persist any pending draft right now, synchronously.
   * Returns the authoritative post-flush entries so callers can chain further
   * mutations without touching stale React state. Does NOT update selectedId.
   */
  function flushPendingSave(currentText?: string): PushLogEntry[] | null {
    if (!saveTimerRef.current) return null;
    cancelPendingSave();
    // Accept an explicit text value (e.g. passed directly from a handler before
    // React state has been committed) or fall back to the synchronously-kept ref.
    const text = currentText ?? logTextRef.current;
    if (!text.trim()) return null;
    const capturedDraftId = draftIdRef.current;
    const snapshot = loadHistory();
    const idStillExists =
      capturedDraftId !== null && snapshot.some((e) => e.id === capturedDraftId);
    const { entries, id } = upsertEntry(
      text,
      snapshot,
      idStillExists ? capturedDraftId : null
    );
    draftIdRef.current = id; // keep draft pointer consistent after flush
    setHistory(entries);
    return entries;
  }

  function scheduleSave(text: string) {
    cancelPendingSave();
    if (!text.trim()) {
      // Clearing the textarea — no save queued. The entry in history (if any)
      // keeps its last-saved content; the user can reload or delete it explicitly.
      return;
    }
    // Capture both at schedule time so the async callback is independent of
    // any state changes that happen before it fires.
    const capturedDraftId = draftIdRef.current;
    saveTimerRef.current = setTimeout(() => {
      saveTimerRef.current = null; // clear ref once executing
      // Guard: if the draft entry was deleted between schedule and fire,
      // create a new entry rather than resurrecting or corrupting a sibling.
      const snapshot = loadHistory();
      const idStillExists =
        capturedDraftId !== null && snapshot.some((e) => e.id === capturedDraftId);
      const { entries, id } = upsertEntry(
        text,
        snapshot,
        idStillExists ? capturedDraftId : null
      );
      draftIdRef.current = id; // update draft pointer to wherever save landed
      setHistory(entries);
      setSelectedId(id);
    }, 1200);
  }

  function handleTextChange(text: string) {
    logTextRef.current = text; // synchronous — must happen before any flush
    setLogText(text);
    scheduleSave(text);
  }

  function selectEntry(entry: PushLogEntry) {
    // Flush any in-progress draft before switching so edits are preserved.
    flushPendingSave();
    // Clear draft pointer: viewing a history entry does NOT make the next paste
    // update that entry — it creates a new one.
    draftIdRef.current = null;
    logTextRef.current = entry.raw;
    setSelectedId(entry.id);
    setLogText(entry.raw);
  }

  function deleteEntry(id: string) {
    const isDraft = id === draftIdRef.current;
    let base: PushLogEntry[];
    if (isDraft) {
      // Deleting the entry currently being composed — discard unsaved changes.
      cancelPendingSave();
      draftIdRef.current = null;
      base = history;
    } else {
      // Deleting a different entry — flush the current draft first, then filter
      // from the authoritative post-flush entries (not stale React state).
      const flushed = flushPendingSave();
      base = flushed ?? history;
    }
    const updated = base.filter((e) => e.id !== id);
    persistHistory(updated);
    setHistory(updated);
    if (selectedId === id) {
      // The currently shown entry was deleted — switch to the next one.
      const next = updated[0] ?? null;
      draftIdRef.current = null; // reset draft — next text change = new entry
      logTextRef.current = next?.raw ?? "";
      setSelectedId(next?.id ?? null);
      setLogText(next?.raw ?? "");
    }
  }

  function handleNewEntry() {
    // Flush any in-progress draft so it's not lost.
    flushPendingSave();
    draftIdRef.current = null; // next paste = new entry
    logTextRef.current = "";
    setSelectedId(null);
    setLogText("");
  }

  return (
    <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
      {/* ── History sidebar ── */}
      <div style={{
        width: 180,
        flexShrink: 0,
        borderRight: "1px solid #21262d",
        background: "#0d1117",
        display: "flex",
        flexDirection: "column",
        overflowY: "auto",
      }}>
        <div style={{
          padding: "8px 10px",
          fontSize: 10,
          fontWeight: 700,
          color: "#484f58",
          textTransform: "uppercase",
          letterSpacing: "0.08em",
          borderBottom: "1px solid #21262d",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
        }}>
          <span>History ({history.length}/{PUSH_LOG_HISTORY_MAX})</span>
          <button
            onClick={handleNewEntry}
            title="New entry"
            style={{
              background: "transparent",
              border: "1px solid #30363d",
              borderRadius: 4,
              color: "#7ee787",
              cursor: "pointer",
              fontSize: 14,
              lineHeight: 1,
              padding: "1px 5px",
              fontFamily: "inherit",
            }}
          >+</button>
        </div>

        {history.length === 0 && (
          <div style={{
            padding: "12px 10px",
            fontSize: 11,
            color: "#484f58",
            lineHeight: 1.5,
          }}>
            Paste a log and it will be saved automatically.
          </div>
        )}

        {history.map((entry) => {
          const isSelected = entry.id === selectedId;
          return (
            <div
              key={entry.id}
              onClick={() => selectEntry(entry)}
              style={{
                padding: "8px 10px",
                cursor: "pointer",
                borderBottom: "1px solid #21262d",
                background: isSelected ? "#161b22" : "transparent",
                borderLeft: isSelected ? "2px solid #7ee787" : "2px solid transparent",
                position: "relative",
                transition: "background 0.1s",
              }}
              onMouseEnter={(e) => {
                if (!isSelected) (e.currentTarget as HTMLElement).style.background = "#111418";
                (e.currentTarget.querySelector(".del-btn") as HTMLElement | null)?.style.setProperty("opacity", "1");
              }}
              onMouseLeave={(e) => {
                if (!isSelected) (e.currentTarget as HTMLElement).style.background = "transparent";
                (e.currentTarget.querySelector(".del-btn") as HTMLElement | null)?.style.setProperty("opacity", "0");
              }}
            >
              <EntrySummary entry={entry} />
              <button
                className="del-btn"
                onClick={(e) => { e.stopPropagation(); deleteEntry(entry.id); }}
                title="Delete"
                style={{
                  position: "absolute",
                  top: 6,
                  right: 6,
                  background: "transparent",
                  border: "none",
                  color: "#6e7681",
                  cursor: "pointer",
                  fontSize: 13,
                  lineHeight: 1,
                  padding: "1px 3px",
                  opacity: 0,
                  transition: "opacity 0.1s",
                }}
              >×</button>
            </div>
          );
        })}
      </div>

      {/* ── Main panel ── */}
      <div style={{ flex: 1, padding: "16px", overflowY: "auto", maxWidth: 720 }}>
        <div style={{ marginBottom: 12, fontSize: 12, color: "#8b949e", lineHeight: 1.6 }}>
          Paste the output from <code style={{ color: "#7ee787" }}>bash scripts/sync-github.sh</code> below.
          Logs are saved automatically and kept in history for quick review.
        </div>

        <textarea
          value={logText}
          onChange={(e) => handleTextChange(e.target.value)}
          placeholder={"Paste push log output here...\n\nExample:\n==> Replit HEAD: a1b2c3d\n==> Fetching remote ref...\nWARNING: GitHub has 2 commit(s) that Replit does not:\n    e4f5g6h Fixed a bug on GitHub\n    i7j8k9l Another GitHub-only change\n\nA force-push will permanently overwrite these commits on GitHub."}
          spellCheck={false}
          style={{
            width: "100%",
            minHeight: 220,
            background: "#0d1117",
            color: "#e6edf3",
            border: "1px solid #30363d",
            borderRadius: 8,
            padding: "12px",
            fontSize: 12,
            lineHeight: 1.6,
            fontFamily: "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace",
            resize: "vertical",
            boxSizing: "border-box",
            outline: "none",
          }}
        />

        {parsed && !hasWarning && logText.trim() && (
          <div style={{
            marginTop: 16,
            padding: "12px 16px",
            background: "#0d1117",
            border: "1px solid #238636",
            borderRadius: 8,
            display: "flex",
            alignItems: "center",
            gap: 10,
            fontSize: 13,
            color: "#3fb950",
          }}>
            <span style={{ fontSize: 16 }}>✓</span>
            <span>
              No overwritten commits detected — push was clean.
              {parsed.pushSha && (
                <span style={{ color: "#484f58", fontSize: 11, marginLeft: 8 }}>
                  HEAD {parsed.pushSha}
                </span>
              )}
            </span>
          </div>
        )}

        {hasWarning && parsed && (
          <div style={{
            marginTop: 16,
            background: "#1f1300",
            border: "1px solid #d29922",
            borderRadius: 8,
            overflow: "hidden",
          }}>
            <div style={{
              padding: "10px 16px",
              background: "#2d1f00",
              borderBottom: "1px solid #d29922",
              display: "flex",
              alignItems: "center",
              gap: 10,
            }}>
              <span style={{ fontSize: 16 }}>⚠️</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: 13, color: "#e3b341" }}>
                  GitHub had {parsed.commitCount} commit{parsed.commitCount !== 1 ? "s" : ""} that Replit did not
                </div>
                <div style={{ fontSize: 11, color: "#9e7c26", marginTop: 2 }}>
                  These were permanently overwritten by the force-push and are no longer on GitHub.
                  {parsed.pushSha && (
                    <span style={{ marginLeft: 8 }}>HEAD {parsed.pushSha}</span>
                  )}
                </div>
              </div>
            </div>

            <div style={{ padding: "8px 0" }}>
              {parsed.overwrittenCommits.map((commit, i) => (
                <div
                  key={i}
                  style={{
                    display: "flex",
                    alignItems: "baseline",
                    gap: 10,
                    padding: "6px 16px",
                    borderBottom: i < parsed.overwrittenCommits.length - 1
                      ? "1px solid #2d1f00"
                      : "none",
                  }}
                >
                  <code style={{
                    fontSize: 11,
                    color: "#d29922",
                    background: "#2d1f00",
                    padding: "1px 6px",
                    borderRadius: 4,
                    flexShrink: 0,
                    letterSpacing: "0.03em",
                  }}>
                    {commit.sha}
                  </code>
                  <span style={{ fontSize: 12, color: "#cdb87a", lineHeight: 1.5 }}>
                    {commit.message}
                  </span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

interface PushStatus {
  failed: boolean;
  failedAt: string | null;
  failedMessage: string | null;
  lastStatusLine: string | null;
}

const DISMISSED_KEY = "push-failure-dismissed-timestamps";

function getDismissedSet(): Set<string> {
  try {
    const raw = localStorage.getItem(DISMISSED_KEY);
    if (!raw) return new Set();
    return new Set(JSON.parse(raw) as string[]);
  } catch {
    return new Set();
  }
}

const DISMISSED_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
const DISMISSED_MAX_COUNT = 100;

function saveDismissedSet(set: Set<string>) {
  try {
    const cutoff = Date.now() - DISMISSED_MAX_AGE_MS;
    const pruned = [...set]
      .filter((ts) => new Date(ts).getTime() >= cutoff)
      .sort()
      .slice(-DISMISSED_MAX_COUNT);
    localStorage.setItem(DISMISSED_KEY, JSON.stringify(pruned));
  } catch {
  }
}

function usePushStatus() {
  const [status, setStatus] = useState<PushStatus | null>(null);
  const [reachable, setReachable] = useState<boolean | null>(null);
  const [dismissedTimestamps, setDismissedTimestamps] = useState<Set<string>>(getDismissedSet);

  useEffect(() => {
    let es: EventSource | null = null;
    let retryDelay = 1_000;
    let retryTimer: ReturnType<typeof setTimeout> | null = null;
    let unmounted = false;

    function connect() {
      if (unmounted) return;
      es = new EventSource("/api/push-status/stream");

      es.onmessage = (e) => {
        retryDelay = 1_000;
        try {
          const data: PushStatus = JSON.parse(e.data as string);
          setStatus(data);
          setReachable(true);
        } catch {
          // ignore malformed events
        }
      };

      es.onerror = () => {
        setReachable(false);
        es?.close();
        es = null;
        if (!unmounted) {
          retryTimer = setTimeout(() => {
            retryDelay = Math.min(retryDelay * 2, 30_000);
            connect();
          }, retryDelay);
        }
      };
    }

    connect();

    return () => {
      unmounted = true;
      if (retryTimer) clearTimeout(retryTimer);
      es?.close();
    };
  }, []);

  const failedAt = status?.failedAt ?? null;
  const dismissed = failedAt !== null && dismissedTimestamps.has(failedAt);

  function dismiss() {
    if (!failedAt) return;
    setDismissedTimestamps((prev) => {
      const next = new Set(prev);
      next.add(failedAt);
      saveDismissedSet(next);
      return next;
    });
  }

  return { status, reachable, dismissed, dismiss };
}

function SyncBadge({ status, reachable, onClick }: {
  status: PushStatus | null;
  reachable: boolean | null;
  onClick: () => void;
}) {
  const [tooltipVisible, setTooltipVisible] = useState(false);

  let dotColor: string;
  let label: string;
  let sublabel: string | null = null;
  let tooltipText: string | null = null;

  if (reachable === null) {
    dotColor = "#484f58";
    label = "Checking…";
  } else if (!reachable || status === null) {
    dotColor = "#484f58";
    label = "API unreachable";
  } else if (status.failed) {
    dotColor = "#f85149";
    label = "Push failed";
    sublabel = status.failedAt ?? null;
    tooltipText = status.failedMessage ?? status.failedAt ?? null;
  } else {
    dotColor = "#3fb950";
    label = "GitHub ✓";
    if (status.lastStatusLine) {
      const ts = status.lastStatusLine.replace(/^.*?(\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(?::\d{2})?).*$/, "$1");
      sublabel = ts !== status.lastStatusLine ? ts : status.lastStatusLine.slice(0, 40);
      tooltipText = status.lastStatusLine;
    }
  }

  return (
    <div style={{ position: "relative", display: "inline-flex" }}>
      <button
        onClick={onClick}
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          background: "transparent",
          border: "1px solid #30363d",
          borderRadius: 6,
          padding: "4px 10px",
          cursor: "pointer",
          fontFamily: "inherit",
          color: "#e6edf3",
          textAlign: "left",
          transition: "border-color 0.15s",
        }}
        onMouseEnter={e => {
          e.currentTarget.style.borderColor = "#58a6ff";
          setTooltipVisible(true);
        }}
        onMouseLeave={e => {
          e.currentTarget.style.borderColor = "#30363d";
          setTooltipVisible(false);
        }}
      >
        <span style={{
          width: 8,
          height: 8,
          borderRadius: "50%",
          background: dotColor,
          flexShrink: 0,
          boxShadow: reachable && status && !status.failed ? `0 0 6px ${dotColor}88` : undefined,
        }} />
        <span style={{ display: "flex", flexDirection: "column", gap: 0 }}>
          <span style={{ fontSize: 11, fontWeight: 700, lineHeight: 1.3 }}>{label}</span>
          {sublabel && (
            <span style={{ fontSize: 10, color: "#8b949e", lineHeight: 1.2 }}>{sublabel}</span>
          )}
        </span>
      </button>

      {tooltipVisible && tooltipText && (
        <div style={{
          position: "absolute",
          top: "calc(100% + 6px)",
          right: 0,
          background: "#1c2128",
          border: "1px solid #30363d",
          borderRadius: 6,
          padding: "6px 10px",
          fontSize: 11,
          color: "#e6edf3",
          whiteSpace: "pre-wrap",
          wordBreak: "break-word",
          maxWidth: 320,
          lineHeight: 1.5,
          zIndex: 100,
          boxShadow: "0 4px 12px rgba(0,0,0,0.5)",
          pointerEvents: "none",
        }}>
          {tooltipText}
        </div>
      )}
    </div>
  );
}

function PushFailureBanner({ status, onDismiss }: { status: PushStatus; onDismiss: () => void }) {
  return (
    <div style={{
      background: "#2d0a0a",
      border: "1px solid #da3633",
      borderLeft: "4px solid #f85149",
      borderRadius: 0,
      padding: "10px 16px",
      display: "flex",
      alignItems: "flex-start",
      gap: 12,
      fontSize: 12,
      lineHeight: 1.6,
    }}>
      <span style={{ fontSize: 16, flexShrink: 0, marginTop: 1 }}>✗</span>
      <div style={{ flex: 1 }}>
        <div style={{ fontWeight: 700, color: "#f85149", marginBottom: 2 }}>
          GitHub push failed
        </div>
        <div style={{ color: "#ffa198" }}>
          {status.failedMessage ?? "Push failed — check .github-sync-status for details."}
        </div>
        {status.failedAt && (
          <div style={{ color: "#8b949e", marginTop: 4, fontSize: 11 }}>
            {status.failedAt} · Run <code style={{ color: "#7ee787" }}>bash scripts/sync-github.sh</code> to retry
          </div>
        )}
      </div>
      <button
        onClick={onDismiss}
        title="Dismiss"
        style={{
          background: "transparent",
          border: "none",
          color: "#8b949e",
          cursor: "pointer",
          fontSize: 16,
          lineHeight: 1,
          padding: "0 2px",
          flexShrink: 0,
        }}
      >
        ×
      </button>
    </div>
  );
}

function App() {
  const [copied, setCopied] = useState(false);
  const [activeTab, setActiveTab] = useState<"script" | "push-log">("script");
  const { status, reachable, dismissed, dismiss } = usePushStatus();
  const showBanner = status?.failed === true && !dismissed;

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(scriptText);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      const ta = document.createElement("textarea");
      ta.value = scriptText;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  const download = () => {
    const blob = new Blob([scriptText], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "DmstPushScript_v1.37.txt";
    a.click();
    URL.revokeObjectURL(url);
  };

  const tabStyle = (tab: "script" | "push-log"): React.CSSProperties => ({
    padding: "6px 14px",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
    borderRadius: "6px 6px 0 0",
    border: "none",
    fontFamily: "inherit",
    background: activeTab === tab ? "#0b1020" : "transparent",
    color: activeTab === tab ? "#e6edf3" : "#8b949e",
    borderBottom: activeTab === tab ? "2px solid #7ee787" : "2px solid transparent",
    transition: "color 0.15s",
  });

  return (
    <div style={{
      minHeight: "100vh",
      background: "#0b1020",
      color: "#e6edf3",
      fontFamily: "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace",
      display: "flex",
      flexDirection: "column",
    }}>
      <header style={{
        padding: "12px 16px",
        borderBottom: "1px solid #21262d",
        background: "#161b22",
        display: "flex",
        alignItems: "center",
        gap: 12,
        flexWrap: "wrap",
        position: "sticky",
        top: 0,
        zIndex: 10,
      }}>
        <div style={{ fontWeight: 600, fontSize: 14, color: "#7ee787" }}>
          DmstPushScript v1.37
        </div>
        <div style={{ fontSize: 12, color: "#8b949e" }}>
          {LINE_COUNT.toLocaleString()} lines · {(BYTE_COUNT / 1024).toFixed(1)} KB · production build · probes: ApplicationStdArray/BarcodeAssignment · 2026-06-10
        </div>
        <div style={{ flex: 1 }} />
        <SyncBadge
          status={status}
          reachable={reachable}
          onClick={() => setActiveTab("push-log")}
        />
        {activeTab === "script" && (
          <>
            <button
              onClick={copy}
              style={{
                background: copied ? "#2ea043" : "#238636",
                color: "white",
                border: "none",
                padding: "6px 14px",
                borderRadius: 6,
                fontSize: 12,
                fontWeight: 600,
                cursor: "pointer",
                fontFamily: "inherit",
              }}
            >
              {copied ? "✓ Copied" : "Copy all"}
            </button>
            <button
              onClick={download}
              style={{
                background: "#21262d",
                color: "#e6edf3",
                border: "1px solid #30363d",
                padding: "6px 14px",
                borderRadius: 6,
                fontSize: 12,
                fontWeight: 600,
                cursor: "pointer",
                fontFamily: "inherit",
              }}
            >
              Download .txt
            </button>
          </>
        )}
      </header>

      <div style={{
        background: "#161b22",
        borderBottom: "1px solid #21262d",
        display: "flex",
        gap: 4,
        padding: "0 16px",
      }}>
        <button style={tabStyle("script")} onClick={() => setActiveTab("script")}>
          Script
        </button>
        <button style={tabStyle("push-log")} onClick={() => setActiveTab("push-log")}>
          Push Log
        </button>
      </div>

      {showBanner && status && (
        <PushFailureBanner status={status} onDismiss={dismiss} />
      )}

      {activeTab === "script" && (
        <>
          <div style={{
            padding: "8px 16px",
            background: "#0d1117",
            borderBottom: "1px solid #21262d",
            fontSize: 11,
            color: "#8b949e",
            lineHeight: 1.5,
          }}>
            <strong style={{ color: "#e6edf3" }}>Install:</strong> DMST → Format Data → Scripting tab → Open Script → paste → Save → Write Settings to verifier. Confirm <code style={{ color: "#7ee787" }}>&lt;PushScriptDiag&gt;v1.37 q=r.trucheck m=found&lt;/PushScriptDiag&gt;</code> appears in the next scan output.
          </div>

          <pre style={{
            margin: 0,
            padding: "16px",
            fontSize: 12,
            lineHeight: 1.5,
            overflow: "auto",
            flex: 1,
            whiteSpace: "pre",
            tabSize: 4,
          }}>
            {scriptText}
          </pre>
        </>
      )}

      {activeTab === "push-log" && <PushLogTab />}
    </div>
  );
}

export default App;
