import { useState } from "react";
import scriptText from "./v128.txt?raw";

const LINE_COUNT = scriptText.split("\n").length;
const BYTE_COUNT = new Blob([scriptText]).size;

function App() {
  const [copied, setCopied] = useState(false);

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
    a.download = "DmstPushScript_v1.28.txt";
    a.click();
    URL.revokeObjectURL(url);
  };

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
          DmstPushScript v1.28
        </div>
        <div style={{ fontSize: 12, color: "#8b949e" }}>
          {LINE_COUNT.toLocaleString()} lines · {(BYTE_COUNT / 1024).toFixed(1)} KB · wire release (MinReflectance + ErrorCapacityUsed) · 2026-05-18
        </div>
        <div style={{ flex: 1 }} />
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
      </header>

      <div style={{
        padding: "8px 16px",
        background: "#0d1117",
        borderBottom: "1px solid #21262d",
        fontSize: 11,
        color: "#8b949e",
        lineHeight: 1.5,
      }}>
        <strong style={{ color: "#e6edf3" }}>Install:</strong> DMST → Format Data → Scripting tab → Open Script → paste → Save → Write Settings to verifier. Confirm <code style={{ color: "#7ee787" }}>&lt;PushScriptDiag&gt;v1.28 q=r.trucheck m=found&lt;/PushScriptDiag&gt;</code> appears in the next scan output.
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
    </div>
  );
}

export default App;
