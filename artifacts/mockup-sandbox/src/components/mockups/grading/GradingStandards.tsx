import { useState } from "react";

type Mode = "single" | "multi-version" | "multi-standard";
type StandardKey =
  | "15416:2016" | "15416:2025"
  | "15415:2011" | "15415:2024"
  | "29158:2020" | "29158:2025";

const FAMILIES: Record<string, StandardKey[]> = {
  Linear:  ["15416:2016", "15416:2025"],
  "2D":    ["15415:2011", "15415:2024"],
  DPM:     ["29158:2020", "29158:2025"],
};

const LABEL: Record<StandardKey, string> = {
  "15416:2016": "ISO/IEC 15416:2016",
  "15416:2025": "ISO/IEC 15416:2025",
  "15415:2011": "ISO/IEC 15415:2011",
  "15415:2024": "ISO/IEC 15415:2024",
  "29158:2020": "ISO/IEC 29158:2020",
  "29158:2025": "ISO/IEC 29158:2025",
};

const DEFAULT_PRIMARY: StandardKey = "15415:2024";

export function GradingStandards() {
  const [mode, setMode] = useState<Mode>("single");
  const [selected, setSelected] = useState<Set<StandardKey>>(
    new Set([DEFAULT_PRIMARY])
  );
  const [primary, setPrimary] = useState<StandardKey>(DEFAULT_PRIMARY);
  const [pwPrompt, setPwPrompt] = useState(false);
  const [pwInput, setPwInput] = useState("");
  const [pwError, setPwError] = useState(false);

  const familyOf = (k: StandardKey) =>
    Object.entries(FAMILIES).find(([, v]) => v.includes(k))?.[0];

  const toggle = (k: StandardKey) => {
    const next = new Set(selected);
    if (mode === "single") {
      next.clear();
      next.add(k);
      setPrimary(k);
    } else if (mode === "multi-version") {
      const fam = familyOf(k)!;
      if (next.has(k)) {
        if (next.size > 1) next.delete(k);
      } else {
        // remove any from OTHER families
        for (const s of Array.from(next)) {
          if (familyOf(s) !== fam) next.delete(s);
        }
        next.add(k);
      }
      if (!next.has(primary) && next.size > 0) setPrimary(Array.from(next)[0]);
    } else {
      if (next.has(k)) {
        next.delete(k);
      } else {
        next.add(k);
      }
      if (!next.has(primary) && next.size > 0) setPrimary(Array.from(next)[0]);
    }
    setSelected(next);
  };

  const handleModeChange = (m: Mode) => {
    if (m === "multi-standard") { setPwPrompt(true); return; }
    setMode(m);
    if (m === "single") {
      setSelected(new Set([primary]));
    } else {
      // keep only standards in same family as primary
      const fam = familyOf(primary)!;
      const keep = new Set(Array.from(selected).filter(s => familyOf(s) === fam));
      if (keep.size === 0) keep.add(primary);
      setSelected(keep);
    }
  };

  const confirmPassword = () => {
    if (pwInput === "vccs") {
      setMode("multi-standard");
      setPwPrompt(false);
      setPwInput("");
      setPwError(false);
    } else {
      setPwError(true);
    }
  };

  const isDisabled = (k: StandardKey) => {
    if (mode === "single") return false;
    if (mode === "multi-version") {
      const fam = familyOf(k)!;
      const selectedFamilies = new Set(Array.from(selected).map(s => familyOf(s)));
      if (selectedFamilies.size === 0) return false;
      return !selectedFamilies.has(fam) && selected.size > 0
        && !Array.from(selected).some(s => familyOf(s) === fam);
    }
    return false;
  };

  return (
    <div className="min-h-screen bg-[#f0f0f0] flex items-center justify-center p-6 font-sans">
      <div className="bg-white border border-[#adadad] shadow-md w-[640px]">

        {/* Title bar */}
        <div className="bg-gradient-to-r from-[#1a5fa8] to-[#3a8ed4] px-4 py-2 flex items-center">
          <span className="text-white text-sm font-semibold tracking-wide">
            Grading Standards Configuration
          </span>
        </div>

        <div className="p-5 space-y-5">

          {/* Mode selector */}
          <div className="border border-[#c0c0c0] bg-[#f8f8f8]">
            <div className="bg-[#e8e8e8] border-b border-[#c0c0c0] px-3 py-1">
              <span className="text-xs font-semibold text-[#333] uppercase tracking-wider">
                Grading Mode
              </span>
            </div>
            <div className="px-4 py-3 flex gap-8">
              {(["single","multi-version","multi-standard"] as Mode[]).map(m => (
                <label key={m} className="flex items-center gap-2 cursor-pointer select-none">
                  <input
                    type="radio"
                    name="mode"
                    checked={mode === m}
                    onChange={() => handleModeChange(m)}
                    className="accent-[#1a5fa8]"
                  />
                  <span className="text-sm text-[#222] flex items-center gap-1">
                    {m === "single"         && "Single-Standard"}
                    {m === "multi-version"  && "Multi-Version"}
                    {m === "multi-standard" && (
                      <>Multi-Standard <span className="text-[#999] text-xs">🔒</span></>
                    )}
                  </span>
                </label>
              ))}
            </div>
            {/* Mode description */}
            <div className="px-4 pb-3">
              <p className="text-xs text-[#666] italic">
                {mode === "single"         && "One standard and version per scan. Device grades live trigger using the selected standard."}
                {mode === "multi-version"  && "Multiple versions of the same standard analyzed from a single captured image. Select versions within one standard family."}
                {mode === "multi-standard" && "Any combination of standards and versions from a single captured image. Requires operator authorization."}
              </p>
            </div>
          </div>

          {/* Standard selectors */}
          {Object.entries(FAMILIES).map(([family, keys]) => (
            <div key={family} className="border border-[#c0c0c0]">
              <div className="bg-[#e8e8e8] border-b border-[#c0c0c0] px-3 py-1">
                <span className="text-xs font-semibold text-[#333] uppercase tracking-wider">
                  {family === "Linear" && "Linear  ·  ISO/IEC 15416"}
                  {family === "2D"     && "2D Matrix  ·  ISO/IEC 15415"}
                  {family === "DPM"    && "DPM  ·  ISO/IEC 29158"}
                </span>
              </div>
              <div className="px-4 py-3 flex gap-10">
                {keys.map(k => {
                  const dis = isDisabled(k) && !selected.has(k);
                  const isPrim = primary === k && selected.size > 1;
                  return (
                    <label
                      key={k}
                      className={`flex items-center gap-2 cursor-pointer select-none ${dis ? "opacity-40 cursor-not-allowed" : ""}`}
                    >
                      <input
                        type="checkbox"
                        checked={selected.has(k)}
                        disabled={dis}
                        onChange={() => !dis && toggle(k)}
                        className="accent-[#1a5fa8] w-4 h-4"
                      />
                      <span className="text-sm text-[#222]">{LABEL[k]}</span>
                      {isPrim && (
                        <span className="text-[10px] bg-[#1a5fa8] text-white px-1.5 py-0.5 rounded font-medium">
                          PRIMARY
                        </span>
                      )}
                    </label>
                  );
                })}
              </div>
              {/* Primary selector row — only shown in multi modes */}
              {mode !== "single" && keys.some(k => selected.has(k)) && selected.size > 1 && (
                <div className="px-4 pb-3 flex gap-10">
                  {keys.filter(k => selected.has(k)).map(k => (
                    <label key={k} className="flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="radio"
                        name="primary"
                        checked={primary === k}
                        onChange={() => setPrimary(k)}
                        className="accent-[#888]"
                      />
                      <span className="text-[11px] text-[#666]">Set as primary</span>
                    </label>
                  ))}
                </div>
              )}
            </div>
          ))}

          {/* Status bar */}
          <div className="bg-[#f0f4fa] border border-[#c8d8ee] px-3 py-2 text-xs text-[#1a5fa8]">
            {mode === "single" && (
              <span>Live scan standard: <strong>{LABEL[primary]}</strong></span>
            )}
            {mode !== "single" && (
              <span>
                Primary (live scan): <strong>{LABEL[primary]}</strong>
                {" · "}
                Additional re-analysis: <strong>{Array.from(selected).filter(s => s !== primary).map(s => LABEL[s]).join(", ") || "none selected"}</strong>
              </span>
            )}
          </div>

          {/* Buttons */}
          <div className="flex justify-end gap-2 pt-1">
            <button className="px-5 py-1.5 text-sm bg-[#e8e8e8] border border-[#adadad] hover:bg-[#d8d8d8] text-[#222]">
              Cancel
            </button>
            <button className="px-5 py-1.5 text-sm bg-[#1a5fa8] hover:bg-[#155090] text-white border border-[#1050a0]">
              Apply to Session
            </button>
          </div>
        </div>
      </div>

      {/* Password modal */}
      {pwPrompt && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50">
          <div className="bg-white border border-[#adadad] shadow-lg w-72">
            <div className="bg-gradient-to-r from-[#1a5fa8] to-[#3a8ed4] px-4 py-2">
              <span className="text-white text-sm font-semibold">Authorization Required</span>
            </div>
            <div className="p-4 space-y-3">
              <p className="text-sm text-[#333]">
                Multi-Standard mode requires supervisor authorization.
              </p>
              <input
                type="password"
                placeholder="Enter password"
                value={pwInput}
                onChange={e => { setPwInput(e.target.value); setPwError(false); }}
                onKeyDown={e => e.key === "Enter" && confirmPassword()}
                className="w-full border border-[#adadad] px-2 py-1.5 text-sm focus:outline-none focus:border-[#1a5fa8]"
                autoFocus
              />
              {pwError && <p className="text-xs text-red-600">Incorrect password.</p>}
              <div className="flex justify-end gap-2">
                <button
                  onClick={() => { setPwPrompt(false); setPwInput(""); setPwError(false); }}
                  className="px-4 py-1.5 text-sm bg-[#e8e8e8] border border-[#adadad] hover:bg-[#d8d8d8]"
                >
                  Cancel
                </button>
                <button
                  onClick={confirmPassword}
                  className="px-4 py-1.5 text-sm bg-[#1a5fa8] text-white border border-[#1050a0] hover:bg-[#155090]"
                >
                  Unlock
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
