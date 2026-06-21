import { useState } from "react";

const ALL_TOKENS = [
  { id: "Timestamp",     label: "Time Stamp",        sample: "20260620-143022" },
  { id: "Date",          label: "Date",               sample: "20260620" },
  { id: "JobName",       label: "Job Name",           sample: "Lot-A123" },
  { id: "Company",       label: "Company Name",       sample: "VCCS" },
  { id: "Operator",      label: "Operator",           sample: "JSmith" },
  { id: "BatchNum",      label: "Batch #",            sample: "B5092" },
  { id: "ProductName",   label: "Product Name",       sample: "LabelStock-X" },
  { id: "Symbology",     label: "Symbology",          sample: "DataMatrix" },
  { id: "SymbologyId",   label: "Symbology ID",       sample: "d1" },
  { id: "Grade",         label: "Grade (Letter)",     sample: "A" },
  { id: "GradeNum",      label: "Grade (Numeric)",    sample: "4.0" },
  { id: "FormalGrade",   label: "Formal Grade",       sample: "1.0-16-660-45Q" },
  { id: "RollNum",       label: "Roll Number",        sample: "R007" },
  { id: "ScanNum",       label: "Scan #",             sample: "042" },
  { id: "Data",          label: "Data (truncated)",   sample: "010377086500016..." },
];

const SEPARATORS = [
  { value: "_",  label: "Underscore  _" },
  { value: "-",  label: "Hyphen  -" },
  { value: ".",  label: "Period  ." },
  { value: ",",  label: "Comma  ," },
  { value: " ",  label: "Space" },
  { value: "",   label: "None" },
];

const FACTORY_TEMPLATES: Record<string, string[]> = {
  "Default":           ["Timestamp", "JobName", "Grade"],
  "Audit Trail":       ["Timestamp", "Operator", "Symbology", "Grade", "FormalGrade"],
  "Part Traceability": ["Date", "Company", "ProductName", "BatchNum", "Grade"],
};

type Token = typeof ALL_TOKENS[number];

export function FileNameFormatBuilder() {
  const [templates, setTemplates] = useState<Record<string, string[]>>(FACTORY_TEMPLATES);
  const [activeTemplate, setActiveTemplate] = useState("Default");
  const [selected, setSelected] = useState<string[]>([...FACTORY_TEMPLATES["Default"]]);
  const [separator, setSeparator] = useState("_");
  const [availSel, setAvailSel] = useState<string | null>(null);
  const [chosenSel, setChosenSel] = useState<string | null>(null);
  const [saveAsName, setSaveAsName] = useState("");
  const [showSaveAs, setShowSaveAs] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);

  const available = ALL_TOKENS.filter(t => !selected.includes(t.id));

  const add = () => {
    if (!availSel) return;
    setSelected(s => [...s, availSel]);
    setChosenSel(availSel);
    setAvailSel(null);
  };

  const remove = () => {
    if (!chosenSel) return;
    setSelected(s => s.filter(id => id !== chosenSel));
    setAvailSel(chosenSel);
    setChosenSel(null);
  };

  const removeAll = () => { setSelected([]); setChosenSel(null); };

  const moveUp = () => {
    if (!chosenSel) return;
    setSelected(s => {
      const i = s.indexOf(chosenSel);
      if (i <= 0) return s;
      const n = [...s]; [n[i-1], n[i]] = [n[i], n[i-1]]; return n;
    });
  };

  const moveDown = () => {
    if (!chosenSel) return;
    setSelected(s => {
      const i = s.indexOf(chosenSel);
      if (i < 0 || i >= s.length - 1) return s;
      const n = [...s]; [n[i], n[i+1]] = [n[i+1], n[i]]; return n;
    });
  };

  const loadTemplate = (name: string) => {
    setActiveTemplate(name);
    setSelected([...(templates[name] ?? [])]);
    setAvailSel(null); setChosenSel(null);
  };

  const saveTemplate = () => {
    const name = saveAsName.trim();
    if (!name) return;
    setTemplates(t => ({ ...t, [name]: [...selected] }));
    setActiveTemplate(name);
    setSaveAsName(""); setShowSaveAs(false);
  };

  const deleteTemplate = () => {
    const { [activeTemplate]: _, ...rest } = templates;
    setTemplates(rest);
    const first = Object.keys(rest)[0] ?? "";
    setActiveTemplate(first);
    setSelected([...(rest[first] ?? [])]);
    setConfirmDelete(false);
  };

  const tokenById = (id: string): Token | undefined => ALL_TOKENS.find(t => t.id === id);

  const preview = selected.length === 0
    ? "(no fields selected)"
    : selected.map(id => tokenById(id)?.sample ?? id).join(separator) + ".html";

  const isFactory = Object.keys(FACTORY_TEMPLATES).includes(activeTemplate);

  return (
    <div className="min-h-screen bg-[#f0f0f0] flex items-center justify-center p-6 font-sans text-sm">
      <div className="bg-white border border-[#999] shadow-lg w-[760px] select-none">

        {/* Title bar */}
        <div className="bg-[#0054a6] text-white px-3 py-1.5 flex items-center justify-between">
          <span className="font-medium text-[13px]">Define Report File Name Format</span>
          <button className="text-white hover:bg-white/20 w-5 h-5 flex items-center justify-center text-xs">✕</button>
        </div>

        <div className="p-4 space-y-4">

          {/* Template row */}
          <div className="flex items-center gap-2">
            <label className="text-[12px] text-[#333] w-20 shrink-0">Template:</label>
            <select
              value={activeTemplate}
              onChange={e => loadTemplate(e.target.value)}
              className="border border-[#aaa] bg-white text-[12px] px-2 py-0.5 h-6 flex-1"
            >
              {Object.keys(templates).map(n => (
                <option key={n} value={n}>{n}</option>
              ))}
            </select>
            <button
              onClick={() => setShowSaveAs(true)}
              className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px]"
            >Save As…</button>
            <button
              onClick={() => setConfirmDelete(true)}
              disabled={isFactory}
              className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px] disabled:opacity-40 disabled:cursor-not-allowed"
            >Delete</button>
          </div>

          {showSaveAs && (
            <div className="flex items-center gap-2 bg-[#fffbe6] border border-[#e0c040] px-3 py-2">
              <span className="text-[12px]">Template name:</span>
              <input
                autoFocus
                value={saveAsName}
                onChange={e => setSaveAsName(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter") saveTemplate(); if (e.key === "Escape") setShowSaveAs(false); }}
                className="border border-[#aaa] px-2 h-6 text-[12px] flex-1"
                placeholder="e.g. My Template"
              />
              <button onClick={saveTemplate} className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px]">Save</button>
              <button onClick={() => setShowSaveAs(false)} className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px]">Cancel</button>
            </div>
          )}

          {confirmDelete && (
            <div className="flex items-center gap-2 bg-[#fff0f0] border border-[#e04040] px-3 py-2">
              <span className="text-[12px]">Delete template "{activeTemplate}"?</span>
              <button onClick={deleteTemplate} className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px] text-red-700">Delete</button>
              <button onClick={() => setConfirmDelete(false)} className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] px-3 h-6 text-[12px]">Cancel</button>
            </div>
          )}

          {/* Shuttle */}
          <div className="flex gap-3 items-start">

            <div className="flex-1">
              <div className="text-[12px] font-semibold text-[#333] mb-1">Available Fields</div>
              <div className="border border-[#aaa] h-52 overflow-y-auto bg-white">
                {available.length === 0
                  ? <div className="text-[11px] text-[#999] px-2 py-2 italic">All fields selected</div>
                  : available.map(t => (
                    <div key={t.id}
                      onDoubleClick={() => { setAvailSel(t.id); setTimeout(add, 0); }}
                      onClick={() => setAvailSel(t.id)}
                      className={`px-3 py-1 cursor-pointer text-[12px] hover:bg-[#cce4ff] ${availSel === t.id ? "bg-[#0054a6] text-white" : ""}`}
                    >{t.label}</div>
                  ))}
              </div>
              <div className="text-[10px] text-[#888] mt-1 italic">Double-click to add</div>
            </div>

            <div className="flex flex-col gap-1.5 mt-7 pt-1">
              <button onClick={add} disabled={!availSel}
                className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] disabled:opacity-40 w-12 h-7 text-[12px]">→</button>
              <button onClick={remove} disabled={!chosenSel}
                className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] disabled:opacity-40 w-12 h-7 text-[12px]">←</button>
              <button onClick={removeAll} disabled={selected.length === 0}
                className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] disabled:opacity-40 w-12 h-7 text-[11px] mt-1">Clear</button>
            </div>

            <div className="flex-1">
              <div className="text-[12px] font-semibold text-[#333] mb-1">
                Selected Fields <span className="text-[#888] font-normal">(in order)</span>
              </div>
              <div className="flex gap-1.5">
                <div className="border border-[#aaa] h-52 overflow-y-auto bg-white flex-1">
                  {selected.length === 0
                    ? <div className="text-[11px] text-[#999] px-2 py-2 italic">No fields selected</div>
                    : selected.map(id => {
                      const t = tokenById(id);
                      return (
                        <div key={id}
                          onDoubleClick={() => { setChosenSel(id); setTimeout(remove, 0); }}
                          onClick={() => setChosenSel(id)}
                          className={`px-3 py-1 cursor-pointer text-[12px] flex items-center gap-2 hover:bg-[#cce4ff] ${chosenSel === id ? "bg-[#0054a6] text-white" : ""}`}
                        >
                          <span className="text-[10px] opacity-50">≡</span>
                          {t?.label ?? id}
                        </div>
                      );
                    })}
                </div>
                <div className="flex flex-col gap-1">
                  <button onClick={moveUp}
                    disabled={!chosenSel || selected.indexOf(chosenSel) === 0}
                    className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] disabled:opacity-40 w-8 h-8 text-[13px]" title="Move Up">▲</button>
                  <button onClick={moveDown}
                    disabled={!chosenSel || selected.indexOf(chosenSel) === selected.length - 1}
                    className="border border-[#aaa] bg-[#e9e9e9] hover:bg-[#d9d9d9] disabled:opacity-40 w-8 h-8 text-[13px]" title="Move Down">▼</button>
                </div>
              </div>
              <div className="text-[10px] text-[#888] mt-1 italic">Double-click to remove · ▲▼ to reorder</div>
            </div>
          </div>

          {/* Separator + extension */}
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2">
              <label className="text-[12px] text-[#333] shrink-0">Token separator:</label>
              <select value={separator} onChange={e => setSeparator(e.target.value)}
                className="border border-[#aaa] bg-white text-[12px] px-2 h-6 w-40">
                {SEPARATORS.map(s => (
                  <option key={s.value} value={s.value}>{s.label}</option>
                ))}
              </select>
            </div>
            <div className="flex items-center gap-2">
              <label className="text-[12px] text-[#333] shrink-0">Extension:</label>
              <select className="border border-[#aaa] bg-white text-[12px] px-2 h-6 w-20">
                <option>.html</option>
                <option>.pdf</option>
              </select>
            </div>
          </div>

          {/* Preview */}
          <div>
            <div className="text-[12px] font-semibold text-[#333] mb-1">
              Preview <span className="text-[#888] font-normal text-[11px]">(sample values)</span>
            </div>
            <div className="border border-[#aaa] bg-[#f8f8f8] px-3 py-2 text-[12px] font-mono text-[#333] min-h-[32px] break-all">
              {preview}
            </div>
          </div>

          <div className="bg-[#f0f6ff] border border-[#c8d8f0] px-3 py-2 text-[11px] text-[#555]">
            <span className="font-semibold">Notes:</span>
            {" "}Symbology ID omits the leading ] bracket (e.g. d1, Q1, C0).
            {" "}Formal Grade uses hyphen separators (e.g. 1.0-16-660-45Q).
            {" "}Data truncated at 24 characters.
            {" "}Templates are global across all jobs.
            {" "}Total path + filename ≤ 256 characters.
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
