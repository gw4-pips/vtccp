import { useState, useEffect } from "react";
import scriptText from "./v134.txt?raw";

const LINE_COUNT = scriptText.split("\n").length;
const BYTE_COUNT = new Blob([scriptText]).size;

interface OverwrittenCommit {
  sha: string;
  message: string;
}

interface ParsedPushLog {
  overwrittenCommits: OverwrittenCommit[];
  commitCount: number;
  raw: string;
}

function parsePushLog(log: string): ParsedPushLog {
  const lines = log.split("\n");
  const overwrittenCommits: OverwrittenCommit[] = [];
  let commitCount = 0;
  let inWarningBlock = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

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

  return { overwrittenCommits, commitCount, raw: log };
}

function PushLogTab() {
  const [logText, setLogText] = useState("");
  const parsed = logText.trim() ? parsePushLog(logText) : null;
  const hasWarning = parsed && parsed.overwrittenCommits.length > 0;

  return (
    <div style={{ padding: "16px", maxWidth: 860 }}>
      <div style={{ marginBottom: 12, fontSize: 12, color: "#8b949e", lineHeight: 1.6 }}>
        Paste the output from <code style={{ color: "#7ee787" }}>bash scripts/sync-github.sh</code> below.
        The viewer will highlight any commits that were overwritten by a force-push.
      </div>

      <textarea
        value={logText}
        onChange={(e) => setLogText(e.target.value)}
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
          <span>No overwritten commits detected — push was clean.</span>
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
  );
}

interface PushStatus {
  failed: boolean;
  failedAt: string | null;
  failedMessage: string | null;
  lastStatusLine: string | null;
}

function usePushStatus() {
  const [status, setStatus] = useState<PushStatus | null>(null);
  const [dismissed, setDismissed] = useState(false);

  useEffect(() => {
    let lastFailedAt: string | null = null;

    async function poll() {
      try {
        const res = await fetch("/api/push-status");
        if (!res.ok) return;
        const data: PushStatus = await res.json();
        setStatus(data);
        if (data.failed && data.failedAt !== lastFailedAt) {
          lastFailedAt = data.failedAt;
          setDismissed(false);
        }
        if (!data.failed) {
          setDismissed(false);
        }
      } catch {
      }
    }

    poll();
    const id = setInterval(poll, 30_000);
    return () => clearInterval(id);
  }, []);

  return { status, dismissed, dismiss: () => setDismissed(true) };
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
  const { status, dismissed, dismiss } = usePushStatus();
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
    a.download = "DmstPushScript_v1.34.txt";
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
          DmstPushScript v1.34
        </div>
        <div style={{ fontSize: 12, color: "#8b949e" }}>
          {LINE_COUNT.toLocaleString()} lines · {(BYTE_COUNT / 1024).toFixed(1)} KB · production build · probe campaign complete · 2026-05-25
        </div>
        <div style={{ flex: 1 }} />
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
            <strong style={{ color: "#e6edf3" }}>Install:</strong> DMST → Format Data → Scripting tab → Open Script → paste → Save → Write Settings to verifier. Confirm <code style={{ color: "#7ee787" }}>&lt;PushScriptDiag&gt;v1.34 q=r.trucheck m=found&lt;/PushScriptDiag&gt;</code> appears in the next scan output.
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
