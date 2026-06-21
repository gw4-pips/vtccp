import { useState } from "react";

// ── Schema blocks mirroring TruCheckCompatibleSchema.cs ─────────────────────

type ColDef = { id: string; label: string };

type Block = {
  key: string;
  letter: string;
  title: string;
  cols: ColDef[];
};

const BLOCKS: Block[] = [
  {
    key: "A", letter: "A", title: "Session / Universal",
    cols: [
      { id: "Date",             label: "Date" },
      { id: "Time",             label: "Time" },
      { id: "OperatorId",       label: "Operator Number" },
      { id: "RollNumber",       label: "Roll Number" },
      { id: "JobName",          label: "Job Name" },
      { id: "BatchNumber",      label: "Batch" },
      { id: "CompanyName",      label: "Company" },
      { id: "ProductName",      label: "Product" },
      { id: "Symbology",        label: "Symbology" },
      { id: "SymbologyId",      label: "AIM ID (Symbology ID)" },
      { id: "DecodedData",      label: "Data" },
      { id: "FormalGrade",      label: "Formal Grade" },
      { id: "OverallLetter",    label: "ANSI Letter Grade" },
      { id: "OverallNumeric",   label: "ANSI Numeric Grade" },
      { id: "CustomPassFail",   label: "Custom" },
      { id: "User1",            label: "User 1" },
      { id: "User2",            label: "User 2" },
      { id: "DeviceSerial",     label: "Unit Serial" },
      { id: "DeviceName",       label: "Device Name" },
      { id: "DeviceModel",      label: "Reader Model" },
      { id: "FirmwareVersion",  label: "Firmware" },
      { id: "ConnectionAddress",label: "Connection" },
      { id: "ConnectionMedium", label: "Link Medium" },
      { id: "SensorWidthPx",    label: "Sensor W (px)" },
      { id: "SensorHeightPx",   label: "Sensor H (px)" },
      { id: "SensorPixelPitchUm",label: "Pixel (µm)" },
      { id: "ImageSizeSetting", label: "Image Size" },
      { id: "CalibrationDate",  label: "Last Calibrated" },
      { id: "Aperture",         label: "Aperture (mil)" },
      { id: "Wavelength",       label: "Wavelength (nm)" },
      { id: "Lighting",         label: "Lighting" },
      { id: "Standard",         label: "Standard" },
    ]
  },
  {
    key: "B", letter: "B", title: "1D — ISO 15416",
    cols: [
      { id: "SymbolAnsiGrade_Numeric",   label: "Symbol ANSI Grade" },
      { id: "StartStopGrade_Numeric",    label: "Start/Stop Grade" },
      { id: "StartStopSrpGrade_Numeric", label: "Start/Stop SRP Grade" },
      { id: "Avg_Edge",   label: "Edge" },
      { id: "Avg_RlRd",   label: "Rl/Rd" },
      { id: "Avg_SC",     label: "SC/CC" },
      { id: "Avg_MinEC",  label: "MinEC" },
      { id: "Avg_MOD",    label: "Mod/CMOD" },
      { id: "Avg_Defect", label: "Defect" },
      { id: "Avg_DCOD",   label: "DCD" },
      { id: "Avg_DEC",    label: "DEC" },
      { id: "Avg_LQZ",    label: "QZ-L" },
      { id: "Avg_RQZ",    label: "QZ-R" },
      { id: "Avg_HQZ",    label: "QZ-H" },
      { id: "Avg_MinQZ",  label: "QZ" },
      { id: "BWG_Percent",label: "BWG %" },
      { id: "BWG_Mil",    label: "BWG (mil)" },
      { id: "Magnification",        label: "Magnification" },
      { id: "NominalXDim_1D",       label: "X Dim/Mag" },
      { id: "InspectionZoneHeight", label: "Inspection Zone Height" },
      { id: "DecodableSymbolHeight",label: "Decodable Symbol Height" },
      { id: "Ratio",      label: "Ratio" },
    ]
  },
  {
    key: "C", letter: "C", title: "2D Common — ISO 15415",
    cols: [
      { id: "UEC_Percent",        label: "UEC %" },
      { id: "UEC_Grade_Numeric",  label: "UEC Grade" },
      { id: "SC_Percent",         label: "SC %" },
      { id: "SC_RlRd",            label: "Rl/Rd" },
      { id: "SC_Grade_Numeric",   label: "SC/CC" },
      { id: "MOD_Grade_Numeric",  label: "Mod/CMOD" },
      { id: "RM_Grade_Numeric",   label: "RM Grade" },
      { id: "ANU_Percent",        label: "ANU %" },
      { id: "ANU_Grade_Numeric",  label: "ANU Grade" },
      { id: "GNU_Percent",        label: "GNU %" },
      { id: "GNU_Grade_Numeric",  label: "GNU Grade" },
      { id: "FPD_Grade_Numeric",  label: "FPD" },
      { id: "DECODE_Grade_Numeric",label: "DECODE Grade" },
      { id: "AG_Value",           label: "AG Value" },
      { id: "AG_Grade_Numeric",   label: "AG/DDG" },
      { id: "AG_Grade_Letter",    label: "AG Grade Letter" },
      { id: "OverallPassFail2D",  label: "Overall Grade" },
      { id: "MatrixSize",         label: "Matrix Size" },
      { id: "HorizontalBWG",      label: "Horizontal BWG" },
      { id: "VerticalBWG",        label: "Vertical BWG" },
      { id: "EncodedCharacters",  label: "Encoded Characters" },
      { id: "TotalCodewords",     label: "Total Codewords" },
      { id: "DataCodewords",      label: "Data Codewords" },
      { id: "ErrorCorrectionBudget",label: "Budget" },
      { id: "ErrorsCorrected",    label: "Errors Corrected" },
      { id: "ErrorCapacityUsed",  label: "Error Capacity Used" },
      { id: "ErrorCorrectionType",label: "EC Type" },
      { id: "ImagePolarity",      label: "Image Polarity" },
      { id: "NominalXDim_2D",     label: "Nominal X Dim" },
      { id: "PixelsPerModule",    label: "Pixels/Module" },
      { id: "ContrastUniformity", label: "CU" },
      { id: "MRD",                label: "MRD" },
    ]
  },
  {
    key: "D", letter: "D", title: "Data Matrix — Standard (≤26×26)",
    cols: [
      { id: "LLS_Grade_Numeric", label: "LLS" }, { id: "LLS_Grade_Letter", label: "LLS Ltr" },
      { id: "BLS_Grade_Numeric", label: "BLS" }, { id: "BLS_Grade_Letter", label: "BLS Ltr" },
      { id: "LQZ_Grade_Numeric", label: "LQZ" }, { id: "LQZ_Grade_Letter", label: "LQZ Ltr" },
      { id: "BQZ_Grade_Numeric", label: "BQZ" }, { id: "BQZ_Grade_Letter", label: "BQZ Ltr" },
      { id: "TQZ_Grade_Numeric", label: "TQZ" }, { id: "TQZ_Grade_Letter", label: "TQZ Ltr" },
      { id: "RQZ_Grade_Numeric", label: "RQZ" }, { id: "RQZ_Grade_Letter", label: "RQZ Ltr" },
      { id: "TTR_Percent",       label: "TTR %" }, { id: "TTR_Grade_Numeric", label: "TTR" },
      { id: "RTR_Percent",       label: "RTR %" }, { id: "RTR_Grade_Numeric", label: "RTR" },
      { id: "TCT_Grade_Numeric", label: "TCT" }, { id: "TCT_Grade_Letter", label: "TCT Ltr" },
      { id: "RCT_Grade_Numeric", label: "RCT" }, { id: "RCT_Grade_Letter", label: "RCT Ltr" },
    ]
  },
  {
    key: "E", letter: "E", title: "Data Matrix — Quadrant Expanded (≥32×32)",
    cols: [
      { id: "ULQZ_Grade_Numeric", label: "ULQZ" }, { id: "URQZ_Grade_Numeric", label: "URQZ" },
      { id: "RUQZ_Grade_Numeric", label: "RUQZ" }, { id: "RLQZ_Grade_Numeric", label: "RLQZ" },
      { id: "ULQTTR_Percent", label: "ULQTTR %" }, { id: "ULQTTR_Grade_Numeric", label: "ULQTTR" },
      { id: "URQTTR_Percent", label: "URQTTR %" }, { id: "URQTTR_Grade_Numeric", label: "URQTTR" },
      { id: "LLQTTR_Percent", label: "LLQTTR %" }, { id: "LLQTTR_Grade_Numeric", label: "LLQTTR" },
      { id: "LRQTTR_Percent", label: "LRQTTR %" }, { id: "LRQTTR_Grade_Numeric", label: "LRQTTR" },
      { id: "ULQRTR_Percent", label: "ULQRTR %" }, { id: "ULQRTR_Grade_Numeric", label: "ULQRTR" },
      { id: "URQRTR_Percent", label: "URQRTR %" }, { id: "URQRTR_Grade_Numeric", label: "URQRTR" },
      { id: "LLQRTR_Percent", label: "LLQRTR %" }, { id: "LLQRTR_Grade_Numeric", label: "LLQRTR" },
      { id: "LRQRTR_Percent", label: "LRQRTR %" }, { id: "LRQRTR_Grade_Numeric", label: "LRQRTR" },
      { id: "ULQTCT_Grade_Numeric", label: "ULQTCT" }, { id: "URQTCT_Grade_Numeric", label: "URQTCT" },
      { id: "LLQTCT_Grade_Numeric", label: "LLQTCT" }, { id: "LRQTCT_Grade_Numeric", label: "LRQTCT" },
      { id: "ULQRCT_Grade_Numeric", label: "ULQRCT" }, { id: "URQRCT_Grade_Numeric", label: "URQRCT" },
      { id: "LLQRCT_Grade_Numeric", label: "LLQRCT" }, { id: "LRQRCT_Grade_Numeric", label: "LRQRCT" },
    ]
  },
  {
    key: "F", letter: "F", title: "Military / Standards-Specific",
    cols: [
      { id: "UIDFormat",              label: "UID Format" },
      { id: "MilStd130VersionLetter", label: "MIL-130 Version Letter" },
      { id: "AS9132_Grade_Numeric",   label: "AS9132 Grade" },
      { id: "Rmax",          label: "Rmax" },
      { id: "TargetRmax",    label: "Target Rmax" },
      { id: "RmaxDeviation", label: "Rmax Deviation" },
      { id: "Rmin",          label: "Rmin" },
      { id: "TargetRmin",    label: "Target Rmin" },
      { id: "RminDeviation", label: "Rmin Deviation" },
    ]
  },
  {
    key: "G", letter: "G", title: "Vendor / Part Tracking",
    cols: [
      { id: "VendorName",   label: "Vendor" },
      { id: "PartNumber",   label: "Part Number" },
      { id: "SerialNumber", label: "Serial Number" },
    ]
  },
  {
    key: "H", letter: "H", title: "VTCCP Extensions — QR Code",
    cols: [
      { id: "QR_Version",     label: "QR Version" },
      { id: "QR_ECLevel",     label: "QR EC Level" },
      { id: "QR_MaskPattern", label: "QR Mask Pattern" },
    ]
  },
  {
    key: "I", letter: "I", title: "GS1 Data Format Check",
    cols: [
      { id: "DFC_Standard", label: "DFC Standard" },
      ...Array.from({ length: 8 }, (_, i) => ([
        { id: `DFC_R${i+1}_Name`,  label: `R${i+1} AI Name` },
        { id: `DFC_R${i+1}_Data`,  label: `R${i+1} Data` },
        { id: `DFC_R${i+1}_Check`, label: `R${i+1} Check` },
      ])).flat(),
    ]
  },
  {
    key: "J", letter: "J", title: "OCR — Label Text Recognition",
    cols: [
      { id: "OcrText",    label: "OCR Text" },
      { id: "OcrTier",    label: "OCR Tier" },
      { id: "OcrWinText", label: "OCR Windows Engine" },
      { id: "OcrTessText",label: "OCR Tesseract" },
      { id: "OcrMatch",   label: "OCR Match" },
    ]
  },
];

const ALL_IDS = new Set(BLOCKS.flatMap(b => b.cols.map(c => c.id)));

const PRESETS: Record<string, Set<string>> = {
  "All Columns": ALL_IDS,
  "Summary Only": new Set([
    "Date","Time","OperatorId","JobName","Symbology","SymbologyId","DecodedData",
    "FormalGrade","OverallLetter","OverallNumeric","DeviceName","Standard","CustomNote",
  ]),
  "1D Focus": new Set([
    "Date","Time","JobName","OperatorId","Symbology","SymbologyId","DecodedData",
    "FormalGrade","OverallLetter","OverallNumeric","Standard",
    "SymbolAnsiGrade_Numeric","Avg_SC","Avg_MOD","Avg_Defect","BWG_Percent","NominalXDim_1D",
  ]),
  "2D Focus": new Set([
    "Date","Time","JobName","OperatorId","Symbology","SymbologyId","DecodedData",
    "FormalGrade","OverallLetter","OverallNumeric","Standard","MatrixSize",
    "UEC_Grade_Numeric","SC_Grade_Numeric","MOD_Grade_Numeric","RM_Grade_Numeric",
    "ANU_Grade_Numeric","GNU_Grade_Numeric","FPD_Grade_Numeric","DECODE_Grade_Numeric",
    "DataCodewords","ErrorCorrectionBudget","ErrorsCorrected",
  ]),
  "GS1 / Aerospace": new Set([
    "Date","Time","JobName","OperatorId","Symbology","SymbologyId","DecodedData",
    "FormalGrade","OverallLetter","OverallNumeric","Standard",
    "UIDFormat","MilStd130VersionLetter","AS9132_Grade_Numeric",
    "VendorName","PartNumber","SerialNumber",
    "DFC_Standard","DFC_R1_Name","DFC_R1_Data","DFC_R1_Check",
    "DFC_R2_Name","DFC_R2_Data","DFC_R2_Check",
  ]),
};

function blockState(block: Block, enabled: Set<string>): "all" | "none" | "partial" {
  const on = block.cols.filter(c => enabled.has(c.id)).length;
  if (on === 0) return "none";
  if (on === block.cols.length) return "all";
  return "partial";
}

export function ExcelColumnOptions() {
  const [enabled, setEnabled] = useState<Set<string>>(new Set(ALL_IDS));
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [preset, setPreset] = useState("All Columns");
  const [fileBehavior, setFileBehavior] = useState("session");
  const [freezeHeader, setFreezeHeader] = useState(true);
  const [compactMode, setCompactMode] = useState(false);
  const [saveAsName, setSaveAsName] = useState("");
  const [showSaveAs, setShowSaveAs] = useState(false);
  const [userPresets, setUserPresets] = useState<Record<string, Set<string>>>({});

  const allPresets = { ...PRESETS, ...userPresets };

  const toggleBlock = (block: Block) => {
    const st = blockState(block, enabled);
    setEnabled(prev => {
      const next = new Set(prev);
      if (st === "all") block.cols.forEach(c => next.delete(c.id));
      else block.cols.forEach(c => next.add(c.id));
      return next;
    });
    setPreset("Custom");
  };

  const toggleCol = (id: string) => {
    setEnabled(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
    setPreset("Custom");
  };

  const applyPreset = (name: string) => {
    setPreset(name);
    setEnabled(new Set(allPresets[name]));
  };

  const savePreset = () => {
    const name = saveAsName.trim();
    if (!name) return;
    setUserPresets(p => ({ ...p, [name]: new Set(enabled) }));
    setPreset(name);
    setSaveAsName(""); setShowSaveAs(false);
  };

  const toggleExpand = (key: string) =>
    setExpanded(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });

  const totalSelected = enabled.size;
  const totalCols = ALL_IDS.size;

  const blockColor: Record<string, string> = {
    A: "#e8f0fe", B: "#fff3e0", C: "#e8f5e9", D: "#fce4ec",
    E: "#f3e5f5", F: "#e0f7fa", G: "#fff9c4", H: "#ede7f6",
    I: "#e3f2fd", J: "#f1f8e9",
  };

  return (
    <div className="min-h-screen bg-[#f0f0f0] flex items-start justify-center p-6 font-sans text-sm">
      <div className="bg-white border border-[#999] shadow-lg w-[800px] select-none">

        {/* Title bar */}
        <div className="bg-[#0054a6] text-white px-3 py-1.5 flex items-center justify-between">
          <span className="font-medium text-[13px]">Excel / CSV Log — Column Options</span>
          <button className="text-white hover:bg-white/20 w-5 h-5 flex items-center justify-center text-xs">✕</button>
        </div>

        <div className="p-4 space-y-3">

          {/* ── File Behavior ─────────────────────────────────────────────────── */}
          <fieldset className="border border-[#c0c0c0] px-3 pb-2 pt-0">
            <legend className="text-[12px] text-[#333] px-1">File Behavior</legend>
            <div className="flex flex-wrap gap-x-6 gap-y-1 mt-1">
              {[
                ["session", "New file each session"],
                ["append", "Append to existing file"],
                ["daily",  "New file each calendar day"],
              ].map(([v, l]) => (
                <label key={v} className="flex items-center gap-1.5 cursor-pointer">
                  <input type="radio" value={v} checked={fileBehavior === v}
                    onChange={() => setFileBehavior(v)} className="accent-[#0054a6]" />
                  <span className="text-[12px]">{l}</span>
                </label>
              ))}
              <label className="flex items-center gap-1.5 cursor-pointer ml-auto">
                <input type="checkbox" checked={freezeHeader}
                  onChange={e => setFreezeHeader(e.target.checked)} className="accent-[#0054a6]" />
                <span className="text-[12px]">Freeze header rows</span>
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer">
                <input type="checkbox" checked={compactMode}
                  onChange={e => setCompactMode(e.target.checked)} className="accent-[#0054a6]" />
                <span className="text-[12px]">Hide all-blank columns after write</span>
              </label>
            </div>
          </fieldset>

          {/* ── Preset row ────────────────────────────────────────────────────── */}
          <div className="flex items-center gap-2">
            <label className="text-[12px] text-[#333] shrink-0">Column Preset:</label>
            <select value={preset} onChange={e => applyPreset(e.target.value)}
              className="border border-[#aaa] bg-white text-[12px] px-2 h-6 flex-1">
              {Object.keys(allPresets).map(n => (
                <option key={n} value={n}>{n}</option>
              ))}
              {preset === "Custom" && <option value="Custom">Custom</option>}
            </select>
            <button onClick={() => setShowSaveAs(true)}
              className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px] whitespace-nowrap">
              Save As…
            </button>
            <button onClick={() => { setEnabled(new Set(ALL_IDS)); setPreset("All Columns"); }}
              className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px]">
              All
            </button>
            <button onClick={() => { setEnabled(new Set()); setPreset("Custom"); }}
              className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px]">
              None
            </button>
          </div>

          {showSaveAs && (
            <div className="flex items-center gap-2 bg-[#fffbe6] border border-[#e0c040] px-3 py-2">
              <span className="text-[12px]">Preset name:</span>
              <input autoFocus value={saveAsName} onChange={e => setSaveAsName(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter") savePreset(); if (e.key === "Escape") setShowSaveAs(false); }}
                className="border border-[#aaa] px-2 h-6 text-[12px] flex-1" placeholder="e.g. GS1 Pharma" />
              <button onClick={savePreset} className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px]">Save</button>
              <button onClick={() => setShowSaveAs(false)} className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px]">Cancel</button>
            </div>
          )}

          {/* ── Column Group Accordion ────────────────────────────────────────── */}
          <div className="border border-[#c0c0c0]">
            {/* Header row */}
            <div className="flex items-center bg-[#e0e8f4] px-3 py-1 border-b border-[#c0c0c0] text-[11px] font-semibold text-[#333]">
              <span className="w-5 shrink-0"></span>
              <span className="flex-1">Column Block</span>
              <span className="w-24 text-center">Columns</span>
              <span className="w-24 text-center">Selected</span>
              <span className="w-20 text-right">Quick</span>
            </div>

            {BLOCKS.map((block, bi) => {
              const st = blockState(block, enabled);
              const onCount = block.cols.filter(c => enabled.has(c.id)).length;
              const isExpanded = expanded.has(block.key);
              const bgHdr = bi % 2 === 0 ? "#fafafa" : "#f4f4f4";

              return (
                <div key={block.key} className="border-b border-[#ddd] last:border-b-0">
                  {/* Block row */}
                  <div className="flex items-center px-2 py-1 cursor-pointer hover:bg-[#eef4ff]"
                    style={{ background: bgHdr }}
                    onClick={() => toggleExpand(block.key)}>
                    {/* Expand chevron */}
                    <span className="w-5 text-[11px] text-[#666] shrink-0">{isExpanded ? "▾" : "▸"}</span>

                    {/* Block badge + checkbox */}
                    <label className="flex items-center gap-2 flex-1 cursor-pointer"
                      onClick={e => e.stopPropagation()}>
                      <input type="checkbox"
                        checked={st === "all"}
                        ref={el => { if (el) el.indeterminate = st === "partial"; }}
                        onChange={() => toggleBlock(block)}
                        className="accent-[#0054a6] w-3.5 h-3.5" />
                      <span className="inline-flex items-center justify-center w-6 h-5 text-[11px] font-bold rounded text-white shrink-0"
                        style={{ background: blockColor[block.key] ? "#0054a6" : "#888", opacity: st === "none" ? 0.4 : 1 }}>
                        {block.letter}
                      </span>
                      <span className="text-[12px] text-[#222]"
                        style={{ fontWeight: st !== "none" ? 500 : 400, opacity: st === "none" ? 0.5 : 1 }}>
                        {block.title}
                      </span>
                    </label>

                    <span className="w-24 text-center text-[11px] text-[#666]">{block.cols.length}</span>
                    <span className="w-24 text-center text-[11px]"
                      style={{ color: st === "none" ? "#bbb" : st === "all" ? "#1a7a1a" : "#a06000", fontWeight: 500 }}>
                      {onCount} / {block.cols.length}
                    </span>
                    <span className="w-20 text-right flex items-center justify-end gap-1">
                      <button onClick={e => { e.stopPropagation(); setEnabled(p => { const n = new Set(p); block.cols.forEach(c => n.add(c.id)); return n; }); setPreset("Custom"); }}
                        className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-1.5 h-5 text-[10px]">All</button>
                      <button onClick={e => { e.stopPropagation(); setEnabled(p => { const n = new Set(p); block.cols.forEach(c => n.delete(c.id)); return n; }); setPreset("Custom"); }}
                        className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-1.5 h-5 text-[10px]">None</button>
                    </span>
                  </div>

                  {/* Expanded individual columns */}
                  {isExpanded && (
                    <div className="px-8 py-2 grid grid-cols-3 gap-x-4 gap-y-0.5 border-t border-[#eee]"
                      style={{ background: blockColor[block.key] + "44" }}>
                      {block.cols.map(col => (
                        <label key={col.id} className="flex items-center gap-1.5 cursor-pointer py-0.5">
                          <input type="checkbox"
                            checked={enabled.has(col.id)}
                            onChange={() => toggleCol(col.id)}
                            className="accent-[#0054a6] w-3 h-3 shrink-0" />
                          <span className="text-[11px] text-[#333] truncate"
                            title={col.label}>{col.label}</span>
                        </label>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>

          {/* ── Status bar ────────────────────────────────────────────────────── */}
          <div className="flex items-center justify-between bg-[#f0f0f0] border border-[#ccc] px-3 py-1.5">
            <span className="text-[12px] text-[#444]">
              <span className="font-semibold" style={{ color: totalSelected === 0 ? "#c00" : "#0054a6" }}>
                {totalSelected}
              </span>
              {" / "}{totalCols}{" columns selected"}
            </span>
            <span className="text-[11px] text-[#888]">
              {BLOCKS.filter(b => blockState(b, enabled) !== "none").length} of {BLOCKS.length} blocks active
            </span>
          </div>

          <div className="bg-[#f0f6ff] border border-[#c8d8f0] px-3 py-2 text-[11px] text-[#555]">
            <span className="font-semibold">Notes:</span>
            {" "}Block columns are always written as a contiguous group — individual column suppression
            within a block hides the column but preserves position compatibility with existing log files.
            {" "}Blocks D and E apply only when the graded symbol is ≥32×32 Data Matrix.
            {" "}OCR columns (Block J) require OCR to be enabled for the session.
          </div>

          <div className="flex justify-end gap-2 pt-1 border-t border-[#ddd]">
            <button className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-6 h-7 text-[12px]">OK</button>
            <button className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-6 h-7 text-[12px]">Cancel</button>
          </div>
        </div>
      </div>
    </div>
  );
}
