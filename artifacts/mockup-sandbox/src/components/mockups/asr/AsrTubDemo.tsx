import { useMemo, useState } from "react";
import {
  Activity,
  AlertTriangle,
  Antenna,
  ArrowRight,
  Clock3,
  Cpu,
  FlaskConical,
  Gauge,
  LockKeyhole,
  Network,
  Play,
  Radio,
  RotateCcw,
  ShieldAlert,
  SlidersHorizontal,
  Thermometer,
} from "lucide-react";

type PortState = "enabled" | "disabled" | "unassigned" | "reading" | "no-read" | "error" | "disconnected";
type ProfileId = "verified" | "arena" | "shared";
type DemoMode = "console" | "diagnostics";

type Port = {
  id: number;
  state: PortState;
  antenna?: string;
  tub?: string;
  dwell: string;
  lastEvent: string;
  confidence?: string;
};

type Profile = {
  id: ProfileId;
  name: string;
  subtitle: string;
  badge: "VERIFIED" | "UNVERIFIED" | "CONDITIONAL";
  note: string;
  activeTub: string;
  ports: Record<number, Partial<Port>>;
};

const STATUS_META: Record<PortState, { label: string; className: string; dot: string }> = {
  enabled: { label: "Enabled", className: "state-enabled", dot: "bg-[#33d6b2]" },
  disabled: { label: "Disabled", className: "state-disabled", dot: "bg-[#67778b]" },
  unassigned: { label: "Unassigned", className: "state-unassigned", dot: "bg-[#98a6b5]" },
  reading: { label: "Reading", className: "state-reading", dot: "bg-[#ffca58]" },
  "no-read": { label: "No-read", className: "state-noread", dot: "bg-[#e89061]" },
  error: { label: "Error", className: "state-error", dot: "bg-[#f26d72]" },
  disconnected: { label: "Disconnected", className: "state-disconnected", dot: "bg-[#b45c83]" },
};

const PROFILES: Profile[] = [
  {
    id: "verified",
    name: "Loan tub / current",
    subtitle: "3 antennas · verified hardware map",
    badge: "VERIFIED",
    note: "Current loan tub uses ports 1–3. Port assignments reflect the verified demonstration harness.",
    activeTub: "TUB-LOAN-03",
    ports: {
      1: { state: "reading", antenna: "ANT-A1", tub: "TUB-LOAN-03", dwell: "2.4 s", lastEvent: "Tag present", confidence: "98.6%" },
      2: { state: "enabled", antenna: "ANT-A2", tub: "TUB-LOAN-03", dwell: "1.7 s", lastEvent: "Idle / armed", confidence: "—" },
      3: { state: "no-read", antenna: "ANT-A3", tub: "TUB-LOAN-03", dwell: "4.1 s", lastEvent: "No response", confidence: "0.0%" },
    },
  },
  {
    id: "arena",
    name: "Arena preview / 7 antenna",
    subtitle: "7 antennas · hardware validation pending",
    badge: "UNVERIFIED",
    note: "Topology preview only. Physical switching, isolation and reader behavior have not been validated on this seven-antenna arrangement.",
    activeTub: "TUB-ARENA-07",
    ports: {
      1: { state: "enabled", antenna: "ANT-B1", tub: "TUB-ARENA-07", dwell: "1.2 s", lastEvent: "Idle / armed" },
      2: { state: "reading", antenna: "ANT-B2", tub: "TUB-ARENA-07", dwell: "2.9 s", lastEvent: "Tag present", confidence: "91.4%" },
      3: { state: "enabled", antenna: "ANT-B3", tub: "TUB-ARENA-07", dwell: "1.0 s", lastEvent: "Idle / armed" },
      4: { state: "enabled", antenna: "ANT-B4", tub: "TUB-ARENA-07", dwell: "1.0 s", lastEvent: "Idle / armed" },
      5: { state: "error", antenna: "ANT-B5", tub: "TUB-ARENA-07", dwell: "0.3 s", lastEvent: "Switch timeout" },
      6: { state: "enabled", antenna: "ANT-B6", tub: "TUB-ARENA-07", dwell: "1.0 s", lastEvent: "Idle / armed" },
      7: { state: "disconnected", antenna: "ANT-B7", tub: "TUB-ARENA-07", dwell: "—", lastEvent: "No continuity" },
    },
  },
  {
    id: "shared",
    name: "Shared reader / dual tub",
    subtitle: "two 3-antenna tubs · switching conditional",
    badge: "CONDITIONAL",
    note: "Both tubs are mapped to one reader in the model. Actual switching sequence and reader arbitration remain conditional on hardware behavior.",
    activeTub: "TUB-COND-A",
    ports: {
      1: { state: "reading", antenna: "ANT-C1", tub: "TUB-COND-A", dwell: "2.1 s", lastEvent: "Tag present", confidence: "96.8%" },
      2: { state: "enabled", antenna: "ANT-C2", tub: "TUB-COND-A", dwell: "1.0 s", lastEvent: "Idle / armed" },
      3: { state: "enabled", antenna: "ANT-C3", tub: "TUB-COND-A", dwell: "1.0 s", lastEvent: "Idle / armed" },
      4: { state: "disabled", antenna: "ANT-D1", tub: "TUB-COND-B", dwell: "—", lastEvent: "Standby / not selected" },
      5: { state: "disabled", antenna: "ANT-D2", tub: "TUB-COND-B", dwell: "—", lastEvent: "Standby / not selected" },
      6: { state: "disabled", antenna: "ANT-D3", tub: "TUB-COND-B", dwell: "—", lastEvent: "Standby / not selected" },
    },
  },
];

const DEFAULT_PORTS: Port[] = Array.from({ length: 8 }, (_, i) => ({
  id: i + 1,
  state: "unassigned",
  dwell: "—",
  lastEvent: "No assignment",
}));

function badgeClass(badge: Profile["badge"]) {
  return badge === "VERIFIED" ? "badge-verified" : badge === "UNVERIFIED" ? "badge-unverified" : "badge-conditional";
}

export function AsrTubDemo() {
  const [profileId, setProfileId] = useState<ProfileId>("verified");
  const [mode, setMode] = useState<DemoMode>("console");
  const [ports, setPorts] = useState<Port[]>(DEFAULT_PORTS);
  const [running, setRunning] = useState(false);
  const [selectedPort, setSelectedPort] = useState(1);
  const [event, setEvent] = useState("SIM-2048 · operator reset · 14:32:08");
  const profile = PROFILES.find((item) => item.id === profileId)!;

  const configuredPorts = useMemo(() => {
    return DEFAULT_PORTS.map((base) => ({ ...base, ...profile.ports[base.id] }));
  }, [profile]);

  const visiblePorts = ports === DEFAULT_PORTS ? configuredPorts : ports;
  const activeCount = visiblePorts.filter((port) => ["enabled", "reading"].includes(port.state)).length;
  const faultCount = visiblePorts.filter((port) => ["error", "disconnected", "no-read"].includes(port.state)).length;
  const diagnosticRows = visiblePorts
    .filter((port) => ["error", "disconnected", "no-read"].includes(port.state))
    .map((port, index) => {
      const condition = port.state === "error" ? "Switch timeout" : port.state === "disconnected" ? "No continuity" : "No-read";
      const evidence = port.state === "error"
        ? "No settle acknowledgement after 3.0 s simulated dwell."
        : port.state === "disconnected"
          ? "Port model reports no continuity signal in the simulated event."
          : "Reader completed the simulated dwell without a tag response.";
      const attention = port.state === "error"
        ? "Check switch path, cable seating, then repeat port exercise."
        : port.state === "disconnected"
          ? "Verify antenna lead and connector before enabling this path."
          : "Check tag placement and antenna path, then repeat the read.";
      return {
        ...port,
        condition,
        evidence,
        attention,
        eventId: `SIM-${7712 - index * 32}`,
        eventTime: new Date(Date.now() - index * 46000).toLocaleTimeString(),
      };
    });

  const applyProfile = (id: ProfileId) => {
    setProfileId(id);
    setPorts(DEFAULT_PORTS);
    setSelectedPort(1);
    setRunning(false);
    setEvent(`SIM-${Math.floor(1000 + Math.random() * 8999)} · profile loaded · ${new Date().toLocaleTimeString()}`);
  };

  const reset = () => {
    setPorts(DEFAULT_PORTS);
    setRunning(false);
    setSelectedPort(1);
    setEvent(`SIM-${Math.floor(1000 + Math.random() * 8999)} · operator reset · ${new Date().toLocaleTimeString()}`);
  };

  const exercise = () => {
    const next = visiblePorts.map((port, index) => {
      if (index === 0) return { ...port, state: "reading" as PortState, lastEvent: "Simulated tag present", confidence: "97.2%" };
      if (port.antenna && index === 2) return { ...port, state: "no-read" as PortState, lastEvent: "Simulated no response" };
      return port;
    });
    setPorts(next);
    setRunning(true);
    setEvent(`SIM-${Math.floor(1000 + Math.random() * 8999)} · activity exercise · ${new Date().toLocaleTimeString()}`);
  };

  const cyclePortState = (id: number) => {
    const order: PortState[] = ["enabled", "reading", "no-read", "error", "disconnected", "disabled", "unassigned"];
    setPorts(visiblePorts.map((port) => port.id === id ? { ...port, state: order[(order.indexOf(port.state) + 1) % order.length], lastEvent: "Simulated operator change" } : port));
    setSelectedPort(id);
    setEvent(`SIM-${Math.floor(1000 + Math.random() * 8999)} · port ${id} state cycled · ${new Date().toLocaleTimeString()}`);
  };

  return (
    <div className="asr-shell">
      <style>{`
        .asr-shell{--ink:#dbe7f2;--muted:#8395a8;--line:#2a4054;--panel:#102334;--panel2:#0d1c2a;--cyan:#43dfd0;--amber:#ffca58;min-height:100vh;background:#081521;color:var(--ink);font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background-image:linear-gradient(rgba(67,223,208,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(67,223,208,.025) 1px,transparent 1px);background-size:28px 28px}
        .asr-shell button{font:inherit}.asr-panel{background:linear-gradient(145deg,#12283a,#0d1d2c);border:1px solid var(--line);box-shadow:0 12px 32px rgba(0,0,0,.2)}.asr-panel-head{border-bottom:1px solid var(--line);background:#142c3f}.state-enabled{color:#65e3c0;border-color:#286d68;background:#123a3c}.state-disabled{color:#97a7b7;border-color:#334657;background:#1c2a37}.state-unassigned{color:#a9b6c1;border-color:#3c4c5b;background:#1a2732}.state-reading{color:#ffda77;border-color:#8e6a2c;background:#3c2e18}.state-noread{color:#ffad7c;border-color:#84503b;background:#3c261d}.state-error{color:#ff8588;border-color:#82434d;background:#3c2029}.state-disconnected{color:#d58aaf;border-color:#6b3f5c;background:#321e30}.badge-verified{color:#63e4c2;background:#123b3a;border-color:#2d8074}.badge-unverified{color:#ffce66;background:#49371a;border-color:#966b25}.badge-conditional{color:#d7a9ff;background:#382449;border-color:#70489a}.asr-scan{animation:scan 1.4s ease-in-out infinite}.asr-pulse{animation:pulse 1.3s ease-in-out infinite}@keyframes scan{0%,100%{opacity:.4;transform:scaleX(.6)}50%{opacity:1;transform:scaleX(1)}}@keyframes pulse{50%{opacity:.45}} 
      `}</style>
      <header className="border-b border-[#2a4054] bg-[#091a29]/95 px-5 py-4 backdrop-blur">
        <div className="mx-auto flex max-w-[1480px] flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-4"><div className="flex h-11 w-11 items-center justify-center border border-[#3b6478] bg-[#102f42] text-[#43dfd0]"><Radio size={22}/></div><div><div className="text-[10px] tracking-[.28em] text-[#43dfd0]">ASR / FIELD CONSOLE</div><h1 className="text-lg font-semibold tracking-tight">8-port reader · tub topology lab</h1></div></div>
          <div className="flex items-center gap-3 text-[11px]"><span className="flex items-center gap-2 text-[#72e0c1]"><span className="h-2 w-2 rounded-full bg-[#43dfd0] asr-pulse"/>SIMULATION ONLY</span><span className="border-l border-[#2a4054] pl-3 text-[#8395a8]">Reader ASR-08 / virtual link</span></div>
        </div>
      </header>
      <main className="mx-auto max-w-[1480px] space-y-5 p-5">
        <section className="asr-panel overflow-hidden">
          <div className="asr-panel-head flex flex-wrap items-center justify-between gap-3 px-4 py-3"><div className="flex items-center gap-3"><SlidersHorizontal size={16} className="text-[#43dfd0]"/><span className="text-xs font-bold tracking-[.16em]">DEMO PROFILE</span><span className="text-xs text-[#8395a8]">model-driven topology / reset-safe</span></div><div className="flex gap-2 rounded border border-[#2a4054] bg-[#0b1b29] p-1"><button onClick={()=>setMode("console")} className={`px-3 py-1.5 text-[11px] ${mode==="console"?"bg-[#235565] text-[#8ff3e6]":"text-[#8395a8]"}`}>LIVE CONSOLE</button><button onClick={()=>setMode("diagnostics")} className={`px-3 py-1.5 text-[11px] ${mode==="diagnostics"?"bg-[#563148] text-[#f4b9cf]":"text-[#8395a8]"}`}>BUG / STATUS</button></div></div>
          <div className="grid gap-2 p-3 md:grid-cols-3">{PROFILES.map((item)=><button key={item.id} onClick={()=>applyProfile(item.id)} className={`text-left border p-3 transition hover:-translate-y-0.5 ${profileId===item.id?"border-[#43dfd0] bg-[#163648]":"border-[#2a4054] bg-[#0c1d2b]"}`}><div className="mb-2 flex items-center justify-between gap-2"><span className="text-xs font-bold text-[#dce9f2]">{item.name}</span><span className={`border px-1.5 py-0.5 text-[9px] font-bold ${badgeClass(item.badge)}`}>{item.badge}</span></div><p className="text-[10px] leading-relaxed text-[#8395a8]">{item.subtitle}</p></button>)}</div>
        </section>
        <div className="grid gap-5 xl:grid-cols-[1fr_340px]">
          <section className="asr-panel">
            <div className="asr-panel-head flex flex-wrap items-center justify-between gap-3 px-4 py-3"><div className="flex items-center gap-3"><Cpu size={16} className="text-[#43dfd0]"/><div><div className="text-xs font-bold tracking-[.16em]">READER PORT MAP</div><div className="text-[10px] text-[#8395a8]">All eight physical ports are always exposed</div></div></div><div className="flex items-center gap-4 text-[10px] text-[#91a8b8]"><span><b className="text-[#64e4c0]">{activeCount}</b> armed</span><span><b className="text-[#ff8588]">{faultCount}</b> attention</span><span className="flex items-center gap-1.5">{running&&<span className="h-1.5 w-1.5 rounded-full bg-[#ffca58] asr-pulse"/>}{running?"activity running":"idle"}</span></div></div>
            <div className="grid gap-2 p-3 sm:grid-cols-2 xl:grid-cols-4">{visiblePorts.map((port)=><button key={port.id} onClick={()=>cyclePortState(port.id)} className={`group border p-3 text-left transition hover:-translate-y-0.5 ${selectedPort===port.id?"ring-1 ring-[#43dfd0]":""} ${STATUS_META[port.state].className}`}><div className="mb-3 flex items-center justify-between"><span className="text-[10px] font-bold tracking-[.18em]">PORT {String(port.id).padStart(2,"0")}</span><span className={`h-2 w-2 rounded-full ${STATUS_META[port.state].dot} ${port.state==="reading"?"asr-pulse":""}`}/></div><div className="mb-2 text-sm font-bold">{port.antenna || "—"} <span className="text-[10px] font-normal opacity-70">{STATUS_META[port.state].label}</span></div><div className="space-y-1 text-[10px] opacity-80"><div className="flex justify-between"><span>TUB</span><span>{port.tub || "UNASSIGNED"}</span></div><div className="flex justify-between"><span>DWELL</span><span>{port.dwell}</span></div><div className="truncate border-t border-current/20 pt-1">{port.lastEvent}</div></div><div className="mt-3 h-px bg-current/20"><div className={`h-px bg-current ${port.state==="reading"?"asr-scan w-full":"w-1/3 opacity-50"}`}/></div></button>)}</div>
            <div className="grid gap-3 border-t border-[#2a4054] p-4 md:grid-cols-[1fr_auto]"><div><div className="mb-2 flex items-center gap-2 text-[10px] tracking-[.15em] text-[#8395a8]"><Activity size={13}/> SIMULATED EVENT BUS</div><div className="border border-[#2a4054] bg-[#091a27] px-3 py-2 text-[11px] text-[#a6bac8]">{event}</div></div><div className="flex items-end gap-2"><button onClick={exercise} className="flex items-center gap-2 border border-[#3b847d] bg-[#153c45] px-3 py-2 text-[11px] text-[#91f2e4] transition hover:bg-[#1b5360]"><Play size={13}/> Exercise activity</button><button onClick={reset} className="flex items-center gap-2 border border-[#40566a] bg-[#162736] px-3 py-2 text-[11px] text-[#c0ced8] transition hover:bg-[#1e3447]"><RotateCcw size={13}/> Reset</button></div></div>
          </section>
          <aside className="space-y-5">
            <section className="asr-panel"><div className="asr-panel-head flex items-center gap-2 px-4 py-3 text-xs font-bold tracking-[.16em]"><Antenna size={15} className="text-[#43dfd0]"/> ACTIVE TOPOLOGY</div><div className="p-4"><div className="mb-4 flex items-center justify-between"><div><div className="text-[10px] text-[#8395a8]">DEVICE</div><div className="text-sm font-bold">ASR-08 / reader</div></div><div className="flex h-9 w-9 items-center justify-center border border-[#376271] bg-[#123545] text-[#43dfd0]"><Network size={17}/></div></div><div className="space-y-2">{[1,2,3].map((n)=><div key={n} className="flex items-center gap-2 text-[10px]"><span className="w-12 text-[#8395a8]">P{n}</span><ArrowRight size={12} className="text-[#43dfd0]"/><span className="flex-1 border border-[#2a4054] bg-[#0b1b29] px-2 py-1.5">{profile.ports[n]?.antenna || "—"}</span><span className="text-[#8395a8]">{profile.ports[n]?.tub || "—"}</span></div>)}{profileId!=="verified"&&<div className="pt-1 text-[10px] text-[#8395a8]">+ {Object.keys(profile.ports).length-3} additional mapped paths</div>}</div><div className={`mt-4 border-l-2 px-3 py-2 text-[10px] leading-relaxed ${profile.badge==="VERIFIED"?"border-[#43dfd0] text-[#9fe9dc]":"border-[#ffca58] text-[#f1cf83]"}`}>{profile.note}</div></div></section>
            <section className="asr-panel"><div className="asr-panel-head px-4 py-3 text-xs font-bold tracking-[.16em]">SIGNAL SNAPSHOT</div><div className="grid grid-cols-2 gap-px bg-[#2a4054]"><div className="bg-[#102334] p-3"><Gauge size={15} className="mb-2 text-[#43dfd0]"/><div className="text-lg font-bold">−48 <span className="text-[10px] font-normal text-[#8395a8]">dBm</span></div><div className="text-[10px] text-[#8395a8]">simulated RSSI</div></div><div className="bg-[#102334] p-3"><Thermometer size={15} className="mb-2 text-[#ffca58]"/><div className="text-lg font-bold">31.4 <span className="text-[10px] font-normal text-[#8395a8]">°C</span></div><div className="text-[10px] text-[#8395a8]">reader chassis</div></div></div></section>
          </aside>
        </div>
        {mode==="diagnostics"&&<section className="asr-panel border-[#75405c]"><div className="flex flex-wrap items-center justify-between gap-3 border-b border-[#75405c] bg-[#321e30] px-4 py-3"><div className="flex items-center gap-3"><ShieldAlert size={17} className="text-[#f39aba]"/><div><div className="text-xs font-bold tracking-[.16em] text-[#f4c2d2]">ACTIONABLE BUG / STATUS EVIDENCE</div><div className="text-[10px] text-[#c696aa]">Separate diagnostic view · simulated conditions only</div></div></div><span className="border border-[#8d536a] px-2 py-1 text-[10px] text-[#f2b8ca]">{diagnosticRows.length} OPEN CONDITION{diagnosticRows.length === 1 ? "" : "S"}</span></div><div className="overflow-x-auto"><table className="w-full min-w-[900px] text-left text-[10px]"><thead className="border-b border-[#2a4054] bg-[#0c1b28] text-[#8395a8]"><tr>{["DEVICE / TUB","PHYSICAL PORT","MAPPED ANTENNA","EVENT / TIME","CONDITION","OBSERVED EVIDENCE","SUGGESTED ATTENTION"].map(h=><th key={h} className="px-4 py-3 font-normal tracking-[.12em]">{h}</th>)}</tr></thead><tbody>{diagnosticRows.length > 0 ? diagnosticRows.map((row) => <tr key={row.id} className="border-b border-[#2a4054]"><td className="px-4 py-3 font-bold">ASR-08 / {row.tub || profile.activeTub}</td><td className="px-4 py-3 text-[#ff8588]">P{String(row.id).padStart(2, "0")}</td><td className="px-4 py-3">{row.antenna || "UNASSIGNED"}</td><td className="px-4 py-3 text-[#a7bac8]">{row.eventTime}<br/>{row.eventId}</td><td className="px-4 py-3"><span className={`border px-2 py-1 ${STATUS_META[row.state].className}`}>{row.state.toUpperCase()}</span><div className="mt-1 text-[#8395a8]">{row.condition}</div></td><td className="px-4 py-3 text-[#d7a9b7]">{row.evidence}</td><td className="px-4 py-3 text-[#f0c1d0]">{row.attention}</td></tr>) : <tr><td colSpan={7} className="px-4 py-10 text-center text-[#8ea2b2]"><div className="mb-2 text-lg text-[#43dfd0]">—</div><div>No current fault states in the visible port map.</div><div className="mt-1 text-[10px] text-[#687f92]">Cycle a port or exercise activity to create synthetic evidence.</div></td></tr>}</tbody></table></div><div className="border-t border-[#75405c] px-4 py-3 text-[10px] text-[#b991a4]"><AlertTriangle size={13} className="mr-2 inline text-[#ffca58]"/>Diagnostic evidence is synthetic. It does not assert actual ASR switching, RF, reader, or tub hardware behavior.</div></section>}
        <footer className="flex flex-wrap items-center justify-between gap-3 border-t border-[#2a4054] pt-3 text-[10px] text-[#687f92]"><span className="flex items-center gap-2"><FlaskConical size={13}/> Engineering demonstration · no production control path</span><span className="flex items-center gap-2"><LockKeyhole size={12}/> FlexWedge / VeriWedge isolated</span><span className="flex items-center gap-2"><Clock3 size={12}/> state is local to this mockup</span></footer>
      </main>
    </div>
  );
}

export default AsrTubDemo;