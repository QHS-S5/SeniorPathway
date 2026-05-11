import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  Users,
  AlertTriangle,
  CheckCircle2,
  Clock,
  FileDown,
  Settings as SettingsIcon,
  X,
  Search,
  Filter,
  ChevronRight,
  RotateCcw,
} from "lucide-react";

/* ============================================================
   STORAGE — IndexedDB-backed key-value layer.
   Same async get/set/delete surface area as the original
   window.storage wrapper, so the rest of the app is unchanged.
   ============================================================ */
const DB_NAME = "coursing";
const DB_STORE = "kv";

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      req.result.createObjectStore(DB_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbOp(mode, fn) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, mode);
    const store = tx.objectStore(DB_STORE);
    const result = fn(store);
    tx.oncomplete = () => resolve(result?.result ?? result);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

const store = {
  async get(key) {
    try {
      const req = await idbOp("readonly", (s) => s.get(key));
      return req == null ? null : req;
    } catch {
      return null;
    }
  },
  async set(key, value) {
    try {
      await idbOp("readwrite", (s) => s.put(value, key));
      return true;
    } catch (e) {
      console.error("storage.set failed", e);
      return false;
    }
  },
  async delete(key) {
    try {
      await idbOp("readwrite", (s) => s.delete(key));
      return true;
    } catch {
      return false;
    }
  },
};

/* ============================================================
   COURSE CODE PARSING
   Cell formats observed in QHS S5 export:
   "C833 HGeo HIGH"          -> SQA Higher Geography
   "C847 N5Mat NAT5"         -> SQA N5 Maths
   "C806 AHA&D ADVH"         -> SQA AH Art & Design
   "STUD Study School"       -> Study period (general)
   "STAH STUAH School"       -> Study period (AH carry)
   "L6DE L6DEC School"       -> Wider achievement / college
   "G9GF SPDvt 46"           -> SCQF L6 wider achievement
   ============================================================ */
function parseSlot(cellValue, column) {
  const empty = { column, kind: "empty", raw: "" };
  if (cellValue === null || cellValue === undefined) return empty;
  const raw = String(cellValue).trim();
  if (!raw) return empty;

  const parts = raw.split(/\s+/);
  const code = parts[0];
  const last = (parts[parts.length - 1] || "").toUpperCase();
  const mid = parts.slice(1, -1).join(" ");

  if (code === "STUD") {
    return { column, kind: "study", level: null, raw, code, label: "Study" };
  }
  if (code === "STAH") {
    return {
      column,
      kind: "study",
      level: "AH",
      raw,
      code,
      label: "Study (AH carry)",
    };
  }
  if (code.startsWith("L6") || code.startsWith("G9")) {
    return {
      column,
      kind: "wider",
      level: "SCQF6",
      raw,
      code,
      label: mid || code,
    };
  }

  let level = null;
  if (last === "ADVH") level = "AH";
  else if (last === "HIGH") level = "H";
  else if (last === "NAT5") level = "N5";
  else if (last === "NAT4") level = "N4";
  else if (last === "46") level = "SCQF6";

  let abbrev = mid;
  let subject = abbrev;
  if (level === "H" && abbrev.startsWith("H")) subject = abbrev.slice(1);
  else if (level === "AH" && abbrev.startsWith("AH")) subject = abbrev.slice(2);
  else if (level === "N5" && abbrev.startsWith("N5")) subject = abbrev.slice(2);

  return {
    column,
    kind: level ? "sqa" : "other",
    level,
    raw,
    code,
    abbrev,
    subject,
    label: `${subject || abbrev} ${level || ""}`.trim(),
  };
}

/* ============================================================
   NAME NORMALISATION — for fuzzy SCN join when SEEMiS export
   doesn't carry SCN. Strips middle names, hyphens, punctuation.
   ============================================================ */
function normName(forename, surname) {
  const f = String(forename || "")
    .toLowerCase()
    .split(/\s+/)[0]
    .replace(/[^a-z]/g, "");
  const s = String(surname || "")
    .toLowerCase()
    .replace(/[-'\s]/g, "")
    .replace(/[^a-z]/g, "");
  return `${f}|${s}`;
}

/* ============================================================
   FLAG ENGINE
   Pure: (pupil, attainmentBySCN, config) -> flags[]
   Each flag: { id, severity, label, detail }
   Severity: red | amber | info
   ============================================================ */
const SEVERITY_ORDER = { red: 0, amber: 1, info: 2 };

function runFlags(pupil, attainmentBySCN, config) {
  const flags = [];
  const slots = pupil.slots;
  const real = slots.filter((s) => s.kind !== "empty");
  const studies = slots.filter((s) => s.kind === "study");
  const sqa = slots.filter((s) => s.kind === "sqa");
  const ahCount = slots.filter((s) => s.level === "AH").length;

  if (real.length === 0) {
    flags.push({
      id: "leaver",
      severity: "info",
      label: "Leaver",
      detail: "No course allocation on SEEMiS",
    });
    return flags;
  }

  if (real.length < (config.minLoad ?? 5)) {
    flags.push({
      id: "underLoaded",
      severity: "amber",
      label: "Under-loaded",
      detail: `${real.length} of 6 slots have a real allocation`,
    });
  }

  if (studies.length >= 2 && ahCount < (config.ahCarryThreshold ?? 3)) {
    flags.push({
      id: "excessStudy",
      severity: "amber",
      label: "Excess study periods",
      detail: `${studies.length} studies with ${ahCount} AH${
        ahCount === 1 ? "" : "s"
      } (rule: 2+ studies allowed only with ${
        config.ahCarryThreshold ?? 3
      }+ AHs)`,
    });
  }

  // Workflow flags from SEEMiS columns
  if (pupil.workflow?.conversationRequired) {
    flags.push({
      id: "conversationRequired",
      severity: "amber",
      label: "Manual: conversation required",
      detail: String(pupil.workflow.conversationRequired),
    });
  }
  if (pupil.workflow?.awaitingConsortia) {
    flags.push({
      id: "awaitingConsortia",
      severity: "amber",
      label: "Awaiting consortia confirmation",
      detail: String(pupil.workflow.awaitingConsortia),
    });
  }
  if (pupil.workflow?.changes) {
    flags.push({
      id: "changesPending",
      severity: "info",
      label: "Changes pending in SEEMiS",
      detail: String(pupil.workflow.changes),
    });
  }

  // Prior-attainment-dependent rules — only fire if SCN-joined
  const att = pupil.scn ? attainmentBySCN[pupil.scn] : null;
  if (att && att.length) {
    const bySubject = {};
    for (const r of att) {
      const subj = (r.title || "").toLowerCase().trim();
      if (!subj) continue;
      bySubject[subj] = bySubject[subj] || [];
      bySubject[subj].push(r);
    }

    for (const slot of sqa) {
      if (!slot.subject) continue;
      const matches = findSubjectAttainment(slot, bySubject);

      if (slot.level === "AH") {
        const higher = matches.find(
          (m) => m.level === 6 && m.gradeLetter && ["A", "B"].includes(m.gradeLetter),
        );
        if (!higher) {
          flags.push({
            id: "AHWithoutHigher",
            severity: "red",
            label: "AH without Higher A/B",
            detail: `${slot.subject} chosen at AH — no Higher pass at A or B on file`,
          });
        }
      }
      if (slot.level === "H") {
        const n5 = matches.find((m) => m.level === 5);
        if (
          n5 &&
          n5.gradeLetter &&
          ["D", "NS", "F"].includes(n5.gradeLetter)
        ) {
          flags.push({
            id: "HigherWithLowN5",
            severity: "red",
            label: "Higher with weak N5",
            detail: `${slot.subject} chosen at Higher — N5 graded ${n5.gradeLetter}`,
          });
        }
      }
    }
  }

  return flags;
}

function findSubjectAttainment(slot, bySubject) {
  // slot.abbrev is e.g. HGeo / AHMat / N5Phy — try to match subject loosely
  if (!slot.subject) return [];
  const subj = slot.subject.toLowerCase();
  // Find any subject key that starts with or contains our abbrev fragment
  const keys = Object.keys(bySubject);
  let matches = [];
  for (const k of keys) {
    if (k.startsWith(subj) || k.includes(subj)) {
      matches = matches.concat(bySubject[k]);
    }
  }
  return matches;
}

/* ============================================================
   XLSX PARSING — finds the right sheet by name pattern
   ============================================================ */
function findSheet(wb, patterns) {
  const names = wb.SheetNames;
  for (const p of patterns) {
    const hit = names.find((n) =>
      n.toLowerCase().replace(/[^a-z]/g, "").includes(p),
    );
    if (hit) return hit;
  }
  return null;
}

function parseAllocationSheet(wb) {
  const sheetName =
    findSheet(wb, ["currentseemis", "seemisalloc", "allocation"]) ||
    wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows.map((r) => {
    const slots = ["A", "B", "C", "D", "E", "F"].map((c) => parseSlot(r[c], c));
    return {
      id: `${r.Forename}-${r.Surname}-${r.Year || ""}`
        .toLowerCase()
        .replace(/\s+/g, "-"),
      scn: r.SCN ? String(r.SCN).trim() : null,
      forename: String(r.Forename || "").trim(),
      surname: String(r.Surname || "").trim(),
      fullName: `${String(r.Forename || "").trim()} ${String(r.Surname || "").trim()}`.trim(),
      yearGroup: String(r.Year || "").trim(),
      slots,
      workflow: {
        changes: r.Changes && String(r.Changes).trim() ? r.Changes : null,
        awaitingConsortia:
          r["Awaiting Consortia Conf."] && String(r["Awaiting Consortia Conf."]).trim()
            ? r["Awaiting Consortia Conf."]
            : null,
        conversationRequired:
          r.Conversation_Required && String(r.Conversation_Required).trim()
            ? r.Conversation_Required
            : null,
        spaces: r.Spaces ? Number(r.Spaces) : null,
      },
      _normName: normName(r.Forename, r.Surname),
    };
  });
}

// Grade letter from numeric result (SQA convention)
function resultToGrade(result) {
  const n = Number(result);
  if (isNaN(n)) return null;
  if (n <= 2) return "A";
  if (n <= 4) return "B";
  if (n <= 6) return "C";
  if (n === 7) return "D";
  if (n === 8) return "NS";
  return "F";
}

function parseAttainmentSheet(wb) {
  const sheetName = findSheet(wb, ["totalattainment", "attainment", "insight"]);
  if (!sheetName) return { bySCN: {}, byName: {} };
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const bySCN = {};
  const byName = {};
  for (const r of rows) {
    const scn = r.SCN ? String(r.SCN).trim() : null;
    const fullName = String(
      r["Full Name"] || `${r.Forename || ""} ${r.Surname || ""}`,
    ).trim();
    const rec = {
      title: r.Title || r["Subject AKA"] || "",
      level: Number(r.Level) || null,
      gradeLetter: r.Grade ? String(r.Grade).trim() : resultToGrade(r.Result),
      result: Number(r.Result) || null,
      tariff: Number(r.Tariff) || 0,
      code: r.Code,
      stage: r.Stage,
    };
    if (scn) {
      bySCN[scn] = bySCN[scn] || [];
      bySCN[scn].push(rec);
    }
    const nKey = normName(r.Forename, r.Surname);
    if (nKey && nKey !== "|") {
      byName[nKey] = byName[nKey] || { scn, records: [] };
      byName[nKey].records.push(rec);
      if (!byName[nKey].scn && scn) byName[nKey].scn = scn;
    }
  }
  return { bySCN, byName };
}

/* ============================================================
   ROOT APP
   ============================================================ */
export default function App() {
  const [tab, setTab] = useState("import");
  const [cohort, setCohort] = useState(null);
  const [attainment, setAttainment] = useState({ bySCN: {}, byName: {} });
  const [decisions, setDecisions] = useState({});
  const [config, setConfig] = useState({
    minLoad: 5,
    ahCarryThreshold: 3,
    cohortYear: "2026-27 S5",
    reviewer: "JS",
  });
  const [selectedPupil, setSelectedPupil] = useState(null);
  const [loaded, setLoaded] = useState(false);

  // Hydrate from storage
  useEffect(() => {
    (async () => {
      const c = await store.get("coursing-cohort");
      const a = await store.get("coursing-attainment");
      const d = await store.get("coursing-decisions");
      const cfg = await store.get("coursing-config");
      if (c) setCohort(c);
      if (a) setAttainment(a);
      if (d) setDecisions(d);
      if (cfg) setConfig((prev) => ({ ...prev, ...cfg }));
      if (c) setTab("triage");
      setLoaded(true);
    })();
  }, []);

  // Persist on change (after hydration)
  useEffect(() => {
    if (loaded && cohort) store.set("coursing-cohort", cohort);
  }, [cohort, loaded]);
  useEffect(() => {
    if (loaded) store.set("coursing-attainment", attainment);
  }, [attainment, loaded]);
  useEffect(() => {
    if (loaded) store.set("coursing-decisions", decisions);
  }, [decisions, loaded]);
  useEffect(() => {
    if (loaded) store.set("coursing-config", config);
  }, [config, loaded]);

  // Run flag engine over cohort
  const pupilsWithFlags = useMemo(() => {
    if (!cohort) return [];
    return cohort.pupils.map((p) => {
      // Try to resolve SCN if missing
      let scn = p.scn;
      if (!scn) {
        const hit = attainment.byName?.[p._normName];
        if (hit?.scn) scn = hit.scn;
      }
      const pupilForEngine = { ...p, scn };
      const flags = runFlags(pupilForEngine, attainment.bySCN || {}, config);
      flags.sort((a, b) => SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity]);
      return { ...pupilForEngine, flags };
    });
  }, [cohort, attainment, config]);

  const handleImport = async (file, kind) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    if (kind === "allocation") {
      const pupils = parseAllocationSheet(wb);
      setCohort({
        importedAt: new Date().toISOString(),
        sourceFile: file.name,
        pupils,
      });
      setTab("triage");
    } else if (kind === "attainment") {
      const parsed = parseAttainmentSheet(wb);
      setAttainment(parsed);
    }
  };

  const recordDecision = (pupilId, flagId, decision) => {
    setDecisions((prev) => {
      const next = { ...prev };
      const key = pupilId;
      next[key] = next[key] || {};
      next[key][flagId] = {
        ...decision,
        at: new Date().toISOString(),
        by: config.reviewer,
      };
      return next;
    });
  };

  const clearCohort = async () => {
    if (!confirm("Clear cohort and all decisions? This cannot be undone.")) return;
    setCohort(null);
    setDecisions({});
    await store.delete("coursing-cohort");
    await store.delete("coursing-decisions");
    setTab("import");
  };

  return (
    <div
      className="min-h-screen text-stone-900"
      style={{
        backgroundColor: "#F5F1EA",
        fontFamily:
          "'IBM Plex Sans', system-ui, -apple-system, sans-serif",
      }}
    >
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght,SOFT@9..144,400;9..144,500;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap'); .display { font-family: 'Fraunces', Georgia, serif; font-optical-sizing: auto; letter-spacing: -0.01em; } .mono { font-family: 'IBM Plex Mono', ui-monospace, monospace; } .hairline { border-color: rgba(0,0,0,0.08); }`}</style>

      <Header
        tab={tab}
        setTab={setTab}
        cohort={cohort}
        config={config}
        onReset={clearCohort}
      />

      <main className="max-w-7xl mx-auto px-6 py-8">
        {tab === "import" && (
          <ImportPanel
            cohort={cohort}
            attainment={attainment}
            onImport={handleImport}
          />
        )}
        {tab === "triage" && cohort && (
          <TriagePanel
            pupils={pupilsWithFlags}
            decisions={decisions}
            onSelect={setSelectedPupil}
          />
        )}
        {tab === "decisions" && (
          <DecisionsPanel
            decisions={decisions}
            pupils={pupilsWithFlags}
            onSelect={setSelectedPupil}
          />
        )}
        {tab === "changelist" && (
          <ChangeListPanel decisions={decisions} pupils={pupilsWithFlags} config={config} />
        )}
        {tab === "settings" && (
          <SettingsPanel config={config} setConfig={setConfig} />
        )}
      </main>

      {selectedPupil && (
        <PupilDrawer
          pupil={pupilsWithFlags.find((p) => p.id === selectedPupil)}
          decisions={decisions[selectedPupil] || {}}
          attainment={attainment}
          onClose={() => setSelectedPupil(null)}
          onDecide={(flagId, dec) => recordDecision(selectedPupil, flagId, dec)}
        />
      )}
    </div>
  );
}

/* ============================================================
   HEADER
   ============================================================ */
function Header({ tab, setTab, cohort, config, onReset }) {
  const tabs = [
    { id: "import", label: "Import", icon: Upload },
    { id: "triage", label: "Triage", icon: AlertTriangle, disabled: !cohort },
    { id: "decisions", label: "Decisions", icon: CheckCircle2, disabled: !cohort },
    { id: "changelist", label: "Change list", icon: FileDown, disabled: !cohort },
    { id: "settings", label: "Settings", icon: SettingsIcon },
  ];
  return (
    <header className="border-b hairline" style={{ backgroundColor: "#FBF9F4" }}>
      <div className="max-w-7xl mx-auto px-6 py-5 flex items-baseline justify-between">
        <div>
          <h1 className="display text-2xl font-semibold tracking-tight">
            Coursing
          </h1>
          <div className="text-xs text-stone-600 mt-1 flex items-center gap-3">
            <span className="mono">{config.cohortYear}</span>
            {cohort && (
              <>
                <span>·</span>
                <span>{cohort.pupils.length} pupils imported</span>
                <span>·</span>
                <span>reviewer {config.reviewer}</span>
              </>
            )}
          </div>
        </div>
        <div className="flex items-center gap-1">
          {tabs.map((t) => {
            const Icon = t.icon;
            const active = tab === t.id;
            return (
              <button
                key={t.id}
                disabled={t.disabled}
                onClick={() => setTab(t.id)}
                className={`px-3 py-2 text-sm flex items-center gap-2 transition-colors ${
                  active
                    ? "text-stone-900 border-b-2 border-stone-900 -mb-px"
                    : "text-stone-500 hover:text-stone-800"
                } ${t.disabled ? "opacity-30 cursor-not-allowed" : ""}`}
              >
                <Icon size={14} strokeWidth={1.5} />
                {t.label}
              </button>
            );
          })}
          {cohort && (
            <button
              onClick={onReset}
              title="Clear cohort"
              className="ml-2 p-2 text-stone-400 hover:text-stone-700"
            >
              <RotateCcw size={14} strokeWidth={1.5} />
            </button>
          )}
        </div>
      </div>
    </header>
  );
}

/* ============================================================
   IMPORT PANEL
   ============================================================ */
function ImportPanel({ cohort, attainment, onImport }) {
  const allocRef = useRef(null);
  const attRef = useRef(null);

  const attCount = Object.keys(attainment.bySCN || {}).length;
  const attNameCount = Object.keys(attainment.byName || {}).length;

  return (
    <div className="max-w-3xl">
      <h2 className="display text-3xl font-medium mb-2">Import data</h2>
      <p className="text-stone-600 text-sm mb-8 leading-relaxed">
        Upload the SEEMiS allocation export first — that's the canonical roll
        and load. The Total_Attainment export is optional but unlocks
        prior-attainment-dependent flags (AH-without-Higher, weak-N5 Higher
        choices). SCN is the preferred join key; the app will fall back to
        normalised name matching if SCN isn't present.
      </p>

      <div className="grid gap-4">
        <ImportCard
          title="SEEMiS allocation"
          subtitle="Sheet name containing 'Current_SEEMiS_Allocation' or 'Allocation' will be used. Required columns: Forename, Surname, Year, A, B, C, D, E, F. Optional: SCN, Spaces, Changes, Awaiting Consortia Conf., Conversation_Required."
          onUpload={(f) => onImport(f, "allocation")}
          inputRef={allocRef}
          status={
            cohort
              ? `Loaded: ${cohort.pupils.length} pupils from ${cohort.sourceFile}`
              : null
          }
        />
        <ImportCard
          title="Total attainment (Insight export)"
          subtitle="Optional. Sheet 'Total_Attainment' or similar. Required columns: SCN, Full Name, Title, Level, Grade, Result, Tariff."
          onUpload={(f) => onImport(f, "attainment")}
          inputRef={attRef}
          status={
            attCount
              ? `Loaded: ${attCount} pupils by SCN, ${attNameCount} by name`
              : null
          }
        />
      </div>

      {cohort && (
        <div className="mt-8 p-5 border hairline" style={{ backgroundColor: "#FBF9F4" }}>
          <div className="text-xs uppercase tracking-wider text-stone-500 mb-2">
            Next step
          </div>
          <div className="text-sm">
            Switch to the <span className="display italic">Triage</span> tab to
            see the flag list. Each pupil-flag pair can be reviewed and
            recorded; finalised changes export from the{" "}
            <span className="display italic">Change list</span> tab as CSV.
          </div>
        </div>
      )}
    </div>
  );
}

function ImportCard({ title, subtitle, onUpload, inputRef, status }) {
  return (
    <div className="p-5 border hairline bg-white">
      <div className="flex items-start justify-between gap-4">
        <div className="flex-1">
          <h3 className="display text-lg font-medium mb-1">{title}</h3>
          <p className="text-xs text-stone-600 leading-relaxed">{subtitle}</p>
          {status && (
            <div className="mt-3 text-xs mono text-emerald-700">{status}</div>
          )}
        </div>
        <button
          onClick={() => inputRef.current?.click()}
          className="shrink-0 px-4 py-2 text-sm bg-stone-900 text-white hover:bg-stone-800"
        >
          Choose file
        </button>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) onUpload(f);
            e.target.value = "";
          }}
        />
      </div>
    </div>
  );
}

/* ============================================================
   TRIAGE PANEL
   ============================================================ */
function TriagePanel({ pupils, decisions, onSelect }) {
  const [filter, setFilter] = useState("active");
  const [flagFilter, setFlagFilter] = useState(null);
  const [search, setSearch] = useState("");
  const [year, setYear] = useState("all");

  const yearGroups = useMemo(
    () => Array.from(new Set(pupils.map((p) => p.yearGroup).filter(Boolean))).sort(),
    [pupils],
  );

  const allFlags = useMemo(() => {
    const counts = {};
    pupils.forEach((p) =>
      p.flags.forEach((f) => {
        counts[f.id] = counts[f.id] || { ...f, count: 0 };
        counts[f.id].count++;
      }),
    );
    return Object.values(counts).sort(
      (a, b) => SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity],
    );
  }, [pupils]);

  const filtered = useMemo(() => {
    return pupils.filter((p) => {
      if (year !== "all" && p.yearGroup !== year) return false;
      if (search) {
        const q = search.toLowerCase();
        if (!p.fullName.toLowerCase().includes(q)) return false;
      }
      if (filter === "leavers") return p.flags.some((f) => f.id === "leaver");
      if (filter === "clean")
        return p.flags.length === 0 || p.flags.every((f) => f.id === "leaver");
      // active = has at least one non-leaver flag
      const activeFlags = p.flags.filter((f) => f.id !== "leaver");
      if (activeFlags.length === 0) return false;
      if (flagFilter && !activeFlags.some((f) => f.id === flagFilter)) return false;
      const resolved = (decisions[p.id] || {});
      const unresolved = activeFlags.some((f) => !resolved[f.id]);
      return unresolved || filter === "active-all";
    });
  }, [pupils, filter, flagFilter, search, year, decisions]);

  const counts = useMemo(() => {
    const leavers = pupils.filter((p) => p.flags.some((f) => f.id === "leaver")).length;
    const active = pupils.filter((p) =>
      p.flags.some((f) => f.id !== "leaver"),
    ).length;
    const clean = pupils.length - leavers - active;
    return { total: pupils.length, leavers, active, clean };
  }, [pupils]);

  return (
    <div>
      <div className="flex items-baseline justify-between mb-6">
        <h2 className="display text-3xl font-medium">Triage</h2>
        <div className="flex gap-6 text-xs">
          <Stat label="Cohort" value={counts.total} />
          <Stat label="Leavers" value={counts.leavers} muted />
          <Stat label="Clean" value={counts.clean} accent="emerald" />
          <Stat label="Flagged" value={counts.active} accent="amber" />
        </div>
      </div>

      <div className="grid grid-cols-[280px_1fr] gap-8">
        {/* Sidebar */}
        <aside className="space-y-6">
          <div>
            <div className="text-[10px] uppercase tracking-widest text-stone-500 mb-2">
              View
            </div>
            <div className="space-y-1">
              {[
                ["active", `Active flags (${counts.active})`],
                ["active-all", "Active flags (incl. resolved)"],
                ["clean", `Clean (${counts.clean})`],
                ["leavers", `Leavers (${counts.leavers})`],
              ].map(([k, label]) => (
                <button
                  key={k}
                  onClick={() => setFilter(k)}
                  className={`block w-full text-left px-2 py-1.5 text-sm ${
                    filter === k
                      ? "bg-stone-900 text-white"
                      : "hover:bg-stone-200/50"
                  }`}
                >
                  {label}
                </button>
              ))}
            </div>
          </div>

          <div>
            <div className="text-[10px] uppercase tracking-widest text-stone-500 mb-2">
              Flag type
            </div>
            <div className="space-y-1">
              <button
                onClick={() => setFlagFilter(null)}
                className={`block w-full text-left px-2 py-1.5 text-xs ${
                  !flagFilter ? "bg-stone-900 text-white" : "hover:bg-stone-200/50"
                }`}
              >
                All flag types
              </button>
              {allFlags
                .filter((f) => f.id !== "leaver")
                .map((f) => (
                  <button
                    key={f.id}
                    onClick={() => setFlagFilter(f.id)}
                    className={`flex w-full items-center justify-between px-2 py-1.5 text-xs ${
                      flagFilter === f.id
                        ? "bg-stone-900 text-white"
                        : "hover:bg-stone-200/50"
                    }`}
                  >
                    <span className="flex items-center gap-2">
                      <SeverityDot severity={f.severity} />
                      {f.label}
                    </span>
                    <span className="mono text-stone-400">{f.count}</span>
                  </button>
                ))}
            </div>
          </div>

          <div>
            <div className="text-[10px] uppercase tracking-widest text-stone-500 mb-2">
              Year group
            </div>
            <select
              value={year}
              onChange={(e) => setYear(e.target.value)}
              className="w-full px-2 py-1.5 text-sm bg-white border hairline"
            >
              <option value="all">All</option>
              {yearGroups.map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
          </div>

          <div>
            <div className="text-[10px] uppercase tracking-widest text-stone-500 mb-2">
              Search
            </div>
            <div className="relative">
              <Search
                size={12}
                className="absolute left-2 top-2.5 text-stone-400"
              />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Name..."
                className="w-full pl-7 pr-2 py-1.5 text-sm bg-white border hairline"
              />
            </div>
          </div>
        </aside>

        {/* List */}
        <div>
          {filtered.length === 0 ? (
            <div className="text-sm text-stone-500 italic py-12 text-center">
              No pupils match this filter.
            </div>
          ) : (
            <div className="space-y-2">
              {filtered.map((p) => (
                <PupilRow
                  key={p.id}
                  pupil={p}
                  resolved={decisions[p.id] || {}}
                  onClick={() => onSelect(p.id)}
                />
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function Stat({ label, value, accent, muted }) {
  const color =
    accent === "emerald"
      ? "text-emerald-700"
      : accent === "amber"
        ? "text-amber-700"
        : muted
          ? "text-stone-400"
          : "text-stone-900";
  return (
    <div>
      <div className="text-[10px] uppercase tracking-widest text-stone-500">
        {label}
      </div>
      <div className={`display text-2xl font-medium ${color}`}>{value}</div>
    </div>
  );
}

function SeverityDot({ severity }) {
  const c =
    severity === "red"
      ? "bg-red-600"
      : severity === "amber"
        ? "bg-amber-500"
        : "bg-stone-400";
  return <span className={`inline-block w-1.5 h-1.5 rounded-full ${c}`} />;
}

function PupilRow({ pupil, resolved, onClick }) {
  const activeFlags = pupil.flags.filter((f) => f.id !== "leaver");
  const allResolved =
    activeFlags.length > 0 && activeFlags.every((f) => resolved[f.id]);
  return (
    <button
      onClick={onClick}
      className="w-full text-left bg-white border hairline px-4 py-3 hover:border-stone-400 transition-colors group"
    >
      <div className="flex items-baseline justify-between gap-4">
        <div className="flex items-baseline gap-3 min-w-0">
          <span className="display text-base font-medium truncate">
            {pupil.fullName}
          </span>
          <span className="mono text-xs text-stone-500">{pupil.yearGroup}</span>
          {!pupil.scn && (
            <span className="text-[10px] uppercase tracking-wider text-amber-700">
              no SCN
            </span>
          )}
          {allResolved && (
            <span className="text-[10px] uppercase tracking-wider text-emerald-700">
              resolved
            </span>
          )}
        </div>
        <div className="flex items-center gap-1.5 shrink-0">
          {activeFlags.slice(0, 4).map((f) => (
            <FlagPill key={f.id} flag={f} resolved={!!resolved[f.id]} />
          ))}
          {activeFlags.length > 4 && (
            <span className="text-xs text-stone-500">
              +{activeFlags.length - 4}
            </span>
          )}
          <ChevronRight size={14} className="text-stone-300 group-hover:text-stone-600" />
        </div>
      </div>
      <div className="mt-2 grid grid-cols-6 gap-2 mono text-[11px]">
        {pupil.slots.map((s) => (
          <SlotChip key={s.column} slot={s} />
        ))}
      </div>
    </button>
  );
}

function FlagPill({ flag, resolved }) {
  const cls = resolved
    ? "bg-stone-100 text-stone-500 border-stone-200"
    : flag.severity === "red"
      ? "bg-red-50 text-red-800 border-red-200"
      : flag.severity === "amber"
        ? "bg-amber-50 text-amber-800 border-amber-200"
        : "bg-stone-50 text-stone-700 border-stone-200";
  return (
    <span className={`text-[10px] px-1.5 py-0.5 border ${cls}`}>
      {flag.label}
    </span>
  );
}

function SlotChip({ slot }) {
  if (slot.kind === "empty") {
    return (
      <div className="px-2 py-1 bg-stone-50 text-stone-300 border hairline">
        <span className="text-[9px] uppercase">{slot.column}</span>
        <div>—</div>
      </div>
    );
  }
  const color =
    slot.kind === "study"
      ? "text-stone-500"
      : slot.kind === "wider"
        ? "text-violet-700"
        : slot.level === "AH"
          ? "text-stone-900 font-medium"
          : "text-stone-700";
  return (
    <div className="px-2 py-1 bg-stone-50/50 border hairline">
      <span className="text-[9px] uppercase text-stone-400">{slot.column}</span>
      <div className={color}>
        {slot.label || slot.raw}
      </div>
    </div>
  );
}

/* ============================================================
   DECISIONS PANEL
   ============================================================ */
function DecisionsPanel({ decisions, pupils, onSelect }) {
  const rows = useMemo(() => {
    const out = [];
    for (const pupilId of Object.keys(decisions)) {
      const pupil = pupils.find((p) => p.id === pupilId);
      if (!pupil) continue;
      for (const flagId of Object.keys(decisions[pupilId])) {
        const flag = pupil.flags.find((f) => f.id === flagId);
        out.push({
          pupilId,
          pupil,
          flag,
          flagId,
          decision: decisions[pupilId][flagId],
        });
      }
    }
    return out.sort((a, b) => (b.decision.at || "").localeCompare(a.decision.at || ""));
  }, [decisions, pupils]);

  return (
    <div>
      <h2 className="display text-3xl font-medium mb-6">Decisions</h2>
      {rows.length === 0 ? (
        <p className="text-sm text-stone-500 italic">
          No decisions recorded yet.
        </p>
      ) : (
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr className="text-left text-[10px] uppercase tracking-wider text-stone-500 border-b hairline">
              <th className="py-2">Pupil</th>
              <th className="py-2">Flag</th>
              <th className="py-2">Action</th>
              <th className="py-2">Note</th>
              <th className="py-2">By</th>
              <th className="py-2">At</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr
                key={i}
                onClick={() => onSelect(r.pupilId)}
                className="border-b hairline hover:bg-stone-100/50 cursor-pointer"
              >
                <td className="py-2 display">{r.pupil.fullName}</td>
                <td className="py-2">{r.flag?.label || r.flagId}</td>
                <td className="py-2">
                  <ActionBadge action={r.decision.action} />
                </td>
                <td className="py-2 text-stone-600 max-w-md truncate">
                  {r.decision.note || "—"}
                </td>
                <td className="py-2 mono text-xs">{r.decision.by}</td>
                <td className="py-2 mono text-xs text-stone-500">
                  {new Date(r.decision.at).toLocaleString("en-GB", {
                    day: "2-digit",
                    month: "short",
                    hour: "2-digit",
                    minute: "2-digit",
                  })}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

function ActionBadge({ action }) {
  const map = {
    change: ["Change", "bg-amber-50 text-amber-800 border-amber-200"],
    "no-change": ["No change", "bg-emerald-50 text-emerald-800 border-emerald-200"],
    pending: ["Pending", "bg-stone-100 text-stone-700 border-stone-300"],
    dismissed: ["Dismissed", "bg-stone-50 text-stone-500 border-stone-200"],
  };
  const [label, cls] = map[action] || [action, "bg-stone-100 border-stone-200"];
  return (
    <span className={`text-[10px] uppercase tracking-wider px-2 py-0.5 border ${cls}`}>
      {label}
    </span>
  );
}

/* ============================================================
   CHANGE LIST
   ============================================================ */
function ChangeListPanel({ decisions, pupils, config }) {
  const changes = useMemo(() => {
    const out = [];
    for (const pupilId of Object.keys(decisions)) {
      const pupil = pupils.find((p) => p.id === pupilId);
      if (!pupil) continue;
      for (const flagId of Object.keys(decisions[pupilId])) {
        const d = decisions[pupilId][flagId];
        if (d.action !== "change") continue;
        out.push({
          scn: pupil.scn || "",
          fullName: pupil.fullName,
          yearGroup: pupil.yearGroup,
          flag: pupil.flags.find((f) => f.id === flagId)?.label || flagId,
          column: d.changeColumn || "",
          fromCode: d.fromCode || "",
          toCode: d.toCode || "",
          note: d.note || "",
          decidedAt: d.at,
          by: d.by,
        });
      }
    }
    return out;
  }, [decisions, pupils]);

  const exportCSV = () => {
    const headers = [
      "SCN",
      "Full Name",
      "Year",
      "Flag",
      "Column",
      "From",
      "To",
      "Note",
      "Decided At",
      "By",
    ];
    const rows = changes.map((c) => [
      c.scn,
      c.fullName,
      c.yearGroup,
      c.flag,
      c.column,
      c.fromCode,
      c.toCode,
      c.note,
      c.decidedAt,
      c.by,
    ]);
    const csv = [headers, ...rows]
      .map((r) =>
        r.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(","),
      )
      .join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `coursing-changes-${config.cohortYear.replace(/\s+/g, "_")}-${new Date()
      .toISOString()
      .slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div>
      <div className="flex items-baseline justify-between mb-6">
        <h2 className="display text-3xl font-medium">Change list</h2>
        <button
          onClick={exportCSV}
          disabled={changes.length === 0}
          className="px-4 py-2 text-sm bg-stone-900 text-white hover:bg-stone-800 disabled:opacity-30 flex items-center gap-2"
        >
          <FileDown size={14} /> Export CSV
        </button>
      </div>
      <p className="text-sm text-stone-600 mb-6 leading-relaxed max-w-2xl">
        Decisions marked <span className="display italic">Change</span> in the
        triage drawer appear here. Export as CSV for the data manager to apply
        in SEEMiS. Only changes are exported — no-change and dismissed
        decisions stay in the audit trail under Decisions.
      </p>
      {changes.length === 0 ? (
        <p className="text-sm text-stone-500 italic">
          No changes recorded. Open a pupil from Triage and select{" "}
          <span className="display italic">Change</span> as the action.
        </p>
      ) : (
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr className="text-left text-[10px] uppercase tracking-wider text-stone-500 border-b hairline">
              <th className="py-2">SCN</th>
              <th className="py-2">Pupil</th>
              <th className="py-2">Year</th>
              <th className="py-2">Col</th>
              <th className="py-2">From</th>
              <th className="py-2">To</th>
              <th className="py-2">Note</th>
            </tr>
          </thead>
          <tbody>
            {changes.map((c, i) => (
              <tr key={i} className="border-b hairline">
                <td className="py-2 mono text-xs">{c.scn || "—"}</td>
                <td className="py-2 display">{c.fullName}</td>
                <td className="py-2 mono text-xs">{c.yearGroup}</td>
                <td className="py-2 mono text-xs">{c.column}</td>
                <td className="py-2 mono text-xs">{c.fromCode}</td>
                <td className="py-2 mono text-xs">{c.toCode}</td>
                <td className="py-2 text-stone-600">{c.note}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

/* ============================================================
   SETTINGS
   ============================================================ */
function SettingsPanel({ config, setConfig }) {
  return (
    <div className="max-w-xl">
      <h2 className="display text-3xl font-medium mb-6">Settings</h2>
      <div className="space-y-5">
        <Field label="Cohort year">
          <input
            value={config.cohortYear}
            onChange={(e) =>
              setConfig({ ...config, cohortYear: e.target.value })
            }
            className="w-full px-3 py-2 bg-white border hairline"
          />
        </Field>
        <Field label="Reviewer initials">
          <input
            value={config.reviewer}
            onChange={(e) => setConfig({ ...config, reviewer: e.target.value })}
            className="w-full px-3 py-2 bg-white border hairline"
          />
        </Field>
        <Field
          label="Minimum subjects (under-loaded threshold)"
          help="Pupil flagged if fewer than this many slots have a real allocation. Studies count as not-real for this threshold."
        >
          <input
            type="number"
            value={config.minLoad}
            onChange={(e) =>
              setConfig({ ...config, minLoad: Number(e.target.value) })
            }
            className="w-32 px-3 py-2 bg-white border hairline"
          />
        </Field>
        <Field
          label="AH-carry threshold for two studies"
          help="A pupil with 2 studies is acceptable IF they have at least this many AHs. Default 3."
        >
          <input
            type="number"
            value={config.ahCarryThreshold}
            onChange={(e) =>
              setConfig({
                ...config,
                ahCarryThreshold: Number(e.target.value),
              })
            }
            className="w-32 px-3 py-2 bg-white border hairline"
          />
        </Field>
      </div>
    </div>
  );
}

function Field({ label, help, children }) {
  return (
    <div>
      <label className="block text-xs uppercase tracking-wider text-stone-600 mb-1.5">
        {label}
      </label>
      {children}
      {help && (
        <p className="text-xs text-stone-500 mt-1.5 leading-relaxed">{help}</p>
      )}
    </div>
  );
}

/* ============================================================
   PUPIL DRAWER
   ============================================================ */
function PupilDrawer({ pupil, decisions, attainment, onClose, onDecide }) {
  if (!pupil) return null;
  const attRecords =
    (pupil.scn && attainment.bySCN?.[pupil.scn]) ||
    attainment.byName?.[pupil._normName]?.records ||
    [];

  return (
    <div
      className="fixed inset-0 z-50 bg-stone-900/30"
      onClick={onClose}
    >
      <div
        className="absolute right-0 top-0 bottom-0 w-full max-w-2xl bg-white shadow-2xl overflow-y-auto"
        onClick={(e) => e.stopPropagation()}
        style={{ backgroundColor: "#FBF9F4" }}
      >
        <div
          className="sticky top-0 z-10 flex items-baseline justify-between px-6 py-4 border-b hairline"
          style={{ backgroundColor: "#FBF9F4" }}
        >
          <div>
            <h2 className="display text-2xl font-medium">{pupil.fullName}</h2>
            <div className="text-xs text-stone-500 mt-0.5 mono">
              {pupil.yearGroup} · {pupil.scn ? `SCN ${pupil.scn}` : "no SCN"}
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-stone-200/50">
            <X size={18} />
          </button>
        </div>

        <div className="px-6 py-5">
          <h3 className="text-[10px] uppercase tracking-widest text-stone-500 mb-2">
            Allocation
          </h3>
          <div className="grid grid-cols-6 gap-2 mb-6 mono text-xs">
            {pupil.slots.map((s) => (
              <SlotChip key={s.column} slot={s} />
            ))}
          </div>

          <h3 className="text-[10px] uppercase tracking-widest text-stone-500 mb-3">
            Flags ({pupil.flags.length})
          </h3>
          <div className="space-y-3 mb-6">
            {pupil.flags.length === 0 && (
              <p className="text-sm text-stone-500 italic">
                No flags raised.
              </p>
            )}
            {pupil.flags.map((flag) => (
              <FlagDecisionRow
                key={flag.id}
                flag={flag}
                decision={decisions[flag.id]}
                pupil={pupil}
                onDecide={(dec) => onDecide(flag.id, dec)}
              />
            ))}
          </div>

          <h3 className="text-[10px] uppercase tracking-widest text-stone-500 mb-3">
            Prior attainment{" "}
            {attRecords.length === 0 && (
              <span className="text-stone-400 italic normal-case">
                — none on file
              </span>
            )}
          </h3>
          {attRecords.length > 0 && (
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="text-left text-[10px] uppercase tracking-wider text-stone-500 border-b hairline">
                  <th className="py-1.5">Subject</th>
                  <th className="py-1.5">Level</th>
                  <th className="py-1.5">Grade</th>
                  <th className="py-1.5 text-right">Tariff</th>
                </tr>
              </thead>
              <tbody>
                {attRecords
                  .sort((a, b) => (b.level || 0) - (a.level || 0))
                  .map((r, i) => (
                    <tr key={i} className="border-b hairline">
                      <td className="py-1.5">{r.title}</td>
                      <td className="py-1.5 mono">{levelLabel(r.level)}</td>
                      <td className="py-1.5 mono">{r.gradeLetter || "—"}</td>
                      <td className="py-1.5 mono text-right">{r.tariff || "—"}</td>
                    </tr>
                  ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}

function levelLabel(level) {
  if (level === 7) return "AH";
  if (level === 6) return "H";
  if (level === 5) return "N5";
  if (level === 4) return "N4";
  return level ?? "—";
}

function FlagDecisionRow({ flag, decision, pupil, onDecide }) {
  const [showDetail, setShowDetail] = useState(false);
  const [note, setNote] = useState(decision?.note || "");
  const [changeColumn, setChangeColumn] = useState(decision?.changeColumn || "");
  const [toCode, setToCode] = useState(decision?.toCode || "");

  const submit = (action) => {
    const slot = changeColumn ? pupil.slots.find((s) => s.column === changeColumn) : null;
    onDecide({
      action,
      note,
      changeColumn,
      fromCode: slot?.raw || "",
      toCode,
    });
    setShowDetail(false);
  };

  return (
    <div className="border hairline bg-white">
      <div className="flex items-start justify-between gap-4 px-4 py-3">
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-2">
            <SeverityDot severity={flag.severity} />
            <span className="text-sm font-medium">{flag.label}</span>
            {decision && <ActionBadge action={decision.action} />}
          </div>
          <p className="text-xs text-stone-600 mt-1 leading-relaxed">
            {flag.detail}
          </p>
        </div>
        <button
          onClick={() => setShowDetail((v) => !v)}
          className="shrink-0 text-xs text-stone-700 hover:text-stone-900 underline underline-offset-2"
        >
          {decision ? "Edit" : "Decide"}
        </button>
      </div>
      {showDetail && (
        <div className="px-4 pb-4 border-t hairline pt-3 space-y-3">
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className="text-[10px] uppercase tracking-wider text-stone-500">
                Column (if changing)
              </label>
              <select
                value={changeColumn}
                onChange={(e) => setChangeColumn(e.target.value)}
                className="w-full mt-1 px-2 py-1.5 text-sm bg-white border hairline"
              >
                <option value="">—</option>
                {["A", "B", "C", "D", "E", "F"].map((c) => (
                  <option key={c} value={c}>
                    {c}: {pupil.slots.find((s) => s.column === c)?.label || "empty"}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="text-[10px] uppercase tracking-wider text-stone-500">
                To (new code or text)
              </label>
              <input
                value={toCode}
                onChange={(e) => setToCode(e.target.value)}
                placeholder="e.g. C847 N5Mat NAT5"
                className="w-full mt-1 px-2 py-1.5 text-sm bg-white border hairline mono"
              />
            </div>
          </div>
          <div>
            <label className="text-[10px] uppercase tracking-wider text-stone-500">
              Note
            </label>
            <textarea
              value={note}
              onChange={(e) => setNote(e.target.value)}
              rows={2}
              placeholder="Context for the data manager / future reference…"
              className="w-full mt-1 px-2 py-1.5 text-sm bg-white border hairline"
            />
          </div>
          <div className="flex gap-2 pt-1">
            <button
              onClick={() => submit("change")}
              className="px-3 py-1.5 text-xs bg-amber-700 text-white hover:bg-amber-800"
            >
              Confirm change
            </button>
            <button
              onClick={() => submit("no-change")}
              className="px-3 py-1.5 text-xs bg-emerald-700 text-white hover:bg-emerald-800"
            >
              No change needed
            </button>
            <button
              onClick={() => submit("pending")}
              className="px-3 py-1.5 text-xs bg-stone-700 text-white hover:bg-stone-800"
            >
              Pending
            </button>
            <button
              onClick={() => submit("dismissed")}
              className="px-3 py-1.5 text-xs border hairline text-stone-700 hover:bg-stone-100"
            >
              Dismiss
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
