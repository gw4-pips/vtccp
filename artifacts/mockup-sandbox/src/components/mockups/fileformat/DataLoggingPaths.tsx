import { useState } from "react";

type PathSectionProps = {
  title: string;
  path: string;
  onPath: (v: string) => void;
};

function PathRow({ label, value, onChange }: { label: string; value: string; onChange: (v: string) => void }) {
  return (
    <div className="flex items-center gap-2">
      <span className="text-[12px] text-[#333] w-28 shrink-0">{label}</span>
      <input
        value={value}
        onChange={e => onChange(e.target.value)}
        className="border border-[#aaa] text-[12px] px-2 h-6 flex-1 bg-white"
      />
      <button className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] w-8 h-6 text-[11px] shrink-0">…</button>
    </div>
  );
}

function SectionBox({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <fieldset className="border border-[#c0c0c0] px-3 pb-3 pt-0">
      <legend className="text-[12px] text-[#333] px-1">{title}</legend>
      <div className="space-y-2 mt-1">{children}</div>
    </fieldset>
  );
}

const TEMPLATES = ["Default", "Audit Trail", "Part Traceability"];

export function DataLoggingPaths() {
  const [reportPath,  setReportPath]  = useState("C:\\Users\\Administrator\\Documents\\VTCCP\\Reports");
  const [imagePath,   setImagePath]   = useState("C:\\Users\\Administrator\\Documents\\VTCCP\\Images\\Decoded");
  const [nreadPath,   setNreadPath]   = useState("C:\\Users\\Administrator\\Documents\\VTCCP\\Images\\NoRead");
  const [csvPath,     setCsvPath]     = useState("C:\\Users\\Administrator\\Documents\\VTCCP\\ExcelLog");

  const [reportTemplate, setReportTemplate] = useState("Default");
  const [reportExt,      setReportExt]      = useState(".html");

  const [imageUseReport, setImageUseReport]  = useState(true);
  const [imageCustom,    setImageCustom]     = useState("");

  const [nreadUseReport, setNreadUseReport]  = useState(true);
  const [nreadSuffix,    setNreadSuffix]     = useState("ND");
  const [nreadCustom,    setNreadCustom]     = useState("");

  const [csvUseReport,   setCsvUseReport]    = useState(true);
  const [csvCustom,      setCsvCustom]       = useState("");

  const imagePreview = imageUseReport
    ? "20260620-143022_Lot-A123_A.jpg  (from report name)"
    : (imageCustom || "(enter filename prefix)") + ".jpg";

  const nreadPreview = nreadUseReport
    ? "20260620-143022_Lot-A123_A_ND.jpg  (from report name + suffix)"
    : (nreadCustom || "(enter filename prefix)") + ".jpg";

  return (
    <div className="min-h-screen bg-[#f0f0f0] flex items-center justify-center p-6 font-sans text-sm">
      <div className="bg-white border border-[#999] shadow-lg w-[720px] select-none">

        {/* Title bar */}
        <div className="bg-[#0054a6] text-white px-3 py-1.5 flex items-center justify-between">
          <span className="font-medium text-[13px]">Data Logging — Output Paths &amp; File Naming</span>
          <button className="text-white hover:bg-white/20 w-5 h-5 flex items-center justify-center text-xs">✕</button>
        </div>

        <div className="p-4 space-y-3">

          {/* ── Verification Reports ───────────────────────────────────────── */}
          <SectionBox title="Verification Reports">
            <PathRow label="Path" value={reportPath} onChange={setReportPath} />

            <div className="flex items-center gap-2">
              <span className="text-[12px] text-[#333] w-28 shrink-0">File Name Format</span>
              <select
                value={reportTemplate}
                onChange={e => setReportTemplate(e.target.value)}
                className="border border-[#aaa] bg-white text-[12px] px-2 h-6 flex-1"
              >
                {TEMPLATES.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
              <button className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px] shrink-0 whitespace-nowrap">
                Define…
              </button>
            </div>

            <div className="flex items-center gap-2">
              <span className="text-[12px] text-[#333] w-28 shrink-0">File Extension</span>
              <select
                value={reportExt}
                onChange={e => setReportExt(e.target.value)}
                className="border border-[#aaa] bg-white text-[12px] px-2 h-6 w-28"
              >
                <option>.html</option>
                <option>.pdf</option>
              </select>
            </div>
          </SectionBox>

          {/* ── Decoded Images ─────────────────────────────────────────────── */}
          <SectionBox title="Decoded Images (JPEG)">
            <PathRow label="Path" value={imagePath} onChange={setImagePath} />

            <div className="flex items-start gap-2 mt-1">
              <span className="text-[12px] text-[#333] w-28 shrink-0 mt-0.5">Filename</span>
              <div className="flex flex-col gap-1.5 flex-1">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    checked={imageUseReport}
                    onChange={() => setImageUseReport(true)}
                    className="accent-[#0054a6]"
                  />
                  <span className="text-[12px]">Use report file name</span>
                  <span className="text-[11px] text-[#888] italic">(default)</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    checked={!imageUseReport}
                    onChange={() => setImageUseReport(false)}
                    className="accent-[#0054a6]"
                  />
                  <span className="text-[12px]">Custom prefix:</span>
                  <input
                    value={imageCustom}
                    onChange={e => setImageCustom(e.target.value)}
                    disabled={imageUseReport}
                    placeholder="e.g. SCAN"
                    className="border border-[#aaa] px-2 h-5 text-[12px] w-40 disabled:bg-[#f0f0f0] disabled:text-[#999]"
                  />
                </label>
                <div className="text-[11px] text-[#666] font-mono bg-[#f8f8f8] border border-[#ddd] px-2 py-0.5">
                  {imagePreview}
                </div>
              </div>
            </div>
          </SectionBox>

          {/* ── No-Read / Undecoded Images ──────────────────────────────────── */}
          <SectionBox title="No-Read / Undecoded Images (JPEG)">
            <PathRow label="Path" value={nreadPath} onChange={setNreadPath} />

            <div className="flex items-start gap-2 mt-1">
              <span className="text-[12px] text-[#333] w-28 shrink-0 mt-0.5">Filename</span>
              <div className="flex flex-col gap-1.5 flex-1">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    checked={nreadUseReport}
                    onChange={() => setNreadUseReport(true)}
                    className="accent-[#0054a6]"
                  />
                  <span className="text-[12px]">Use report file name + suffix:</span>
                  <input
                    value={nreadSuffix}
                    onChange={e => setNreadSuffix(e.target.value)}
                    className="border border-[#aaa] px-2 h-5 text-[12px] w-16 text-center"
                  />
                  <span className="text-[11px] text-[#888] italic">(default: ND)</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    checked={!nreadUseReport}
                    onChange={() => setNreadUseReport(false)}
                    className="accent-[#0054a6]"
                  />
                  <span className="text-[12px]">Custom prefix:</span>
                  <input
                    value={nreadCustom}
                    onChange={e => setNreadCustom(e.target.value)}
                    disabled={nreadUseReport}
                    placeholder="e.g. NOREAD"
                    className="border border-[#aaa] px-2 h-5 text-[12px] w-40 disabled:bg-[#f0f0f0] disabled:text-[#999]"
                  />
                </label>
                <div className="text-[11px] text-[#666] font-mono bg-[#f8f8f8] border border-[#ddd] px-2 py-0.5">
                  {nreadPreview}
                </div>
              </div>
            </div>
          </SectionBox>

          {/* ── Excel / CSV Log ─────────────────────────────────────────────── */}
          <SectionBox title="Excel / CSV Verification Log">
            <PathRow label="Path" value={csvPath} onChange={setCsvPath} />
            <div className="flex items-start gap-2 mt-1">
              <span className="text-[12px] text-[#333] w-28 shrink-0 mt-0.5">Filename</span>
              <div className="flex flex-col gap-1.5 flex-1">
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    checked={csvUseReport}
                    onChange={() => setCsvUseReport(true)}
                    className="accent-[#0054a6]"
                  />
                  <span className="text-[12px]">Use job name + date</span>
                  <span className="text-[11px] text-[#888] italic">(default)</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer">
                  <input
                    type="radio"
                    checked={!csvUseReport}
                    onChange={() => setCsvUseReport(false)}
                    className="accent-[#0054a6]"
                  />
                  <span className="text-[12px]">Custom prefix:</span>
                  <input
                    value={csvCustom}
                    onChange={e => setCsvCustom(e.target.value)}
                    disabled={csvUseReport}
                    placeholder="e.g. VerifLog"
                    className="border border-[#aaa] px-2 h-5 text-[12px] w-40 disabled:bg-[#f0f0f0] disabled:text-[#999]"
                  />
                </label>
              </div>
            </div>
          </SectionBox>

          {/* Footer note */}
          <div className="bg-[#f0f6ff] border border-[#c8d8f0] px-3 py-2 text-[11px] text-[#555]">
            <span className="font-semibold">Note:</span>
            {" "}All paths are created automatically if they do not exist.
            {" "}JPEG images are L1 barcode-crop from push XML; PNG files saved by DMST are managed separately.
            {" "}File Name Format templates are defined in the Report File Name Format dialog (
            <span className="text-[#0054a6] cursor-pointer underline">Define…</span> above).
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
