import React, { useState, useRef, useMemo } from "react";
import * as XLSX from "xlsx";

/* ============================================================
   VP Education Members Dashboard
   All data lives in memory + the user's own JSON file.
   No network calls. Export JSON / Excel anytime.
   ============================================================ */

// ---------- Brand tokens (Toastmasters-inspired) ----------
const C = {
  blue: "#004165",      // Loyal Blue — chrome, headers
  blueDeep: "#00314D",
  maroon: "#772432",    // True Maroon — signature accent
  gold: "#F2DF74",      // Happy Yellow — active highlights
  paper: "#F6F5F0",     // page background
  ink: "#1C2A33",
  grayLine: "#E3E1D8",
  green: "#2E7D32",
  greenBg: "#E8F3E9",
  amber: "#B45309",
  amberBg: "#FCF3E3",
  red: "#B3261E",
  redBg: "#FBEAE8",
};

const SERIF = "Georgia, 'Times New Roman', serif";

// ---------- Constants ----------
const PATHS = [
  "Presentation Mastery", "Dynamic Leadership", "Visionary Communication",
  "Engaging Humor", "Leadership Development", "Motivational Strategies",
  "Persuasive Influence", "Strategic Relationships", "Team Collaboration",
  "Innovative Planning", "Effective Coaching",
];

const COMMON_ROLES = [
  "Timekeeper", "Ah Counter", "Grammarian", "Table Topics Master",
  "Evaluator", "General Evaluator", "Toastmaster of the Day",
  "Speaker", "Table Topics Speaker",
];

const ONBOARDING_STAGES = [
  { from: 1, to: 14, label: "Attend & observe" },
  { from: 15, to: 30, label: "Timekeeper role" },
  { from: 31, to: 50, label: "Ah Counter / GE support" },
  { from: 51, to: 70, label: "Table Topics participant" },
  { from: 71, to: 85, label: "Table Topics Master" },
  { from: 86, to: 100, label: "Deliver Icebreaker speech" },
];

const RECOGNITION_TYPES = [
  "Best Table Topics",
  "Best Evaluator",
  "First time in a role",
  "Pathways level completion",
];

const DTM_REQUIREMENTS = [
  "Complete first learning path",
  "Complete second learning path",
  "Serve 12 months as a club officer",
  "Serve 12 months as a district officer",
  "Serve as club mentor or club coach",
  "Serve as club sponsor, or conduct Speechcraft / Youth Leadership",
  "Complete the DTM project",
];

const DEFAULT_WEEK_TASKS = [
  "Ask for role volunteers",
  "Close meeting notes",
  "Post recognition to Game Changers",
];

const CLUB_LINK = "https://toastmasterclub.org";

// Cycle definitions: months are 0-indexed (Jan = 0)
const DEFAULT_CYCLES = [
  { id: "c1", name: "Storytelling", span: "Jul–Aug", months: [6, 7] },
  { id: "c2", name: "Humour", span: "Sep–Oct", months: [8, 9] },
  { id: "c3", name: "Vocal Variety", span: "Nov–Dec", months: [10, 11] },
  { id: "c4", name: "Structure & Clarity", span: "Jan–Feb", months: [0, 1] },
  { id: "c5", name: "Persuasion", span: "Mar–Apr", months: [2, 3] },
  { id: "c6", name: "Member's Choice", span: "May–Jun", months: [4, 5] },
].map((c) => ({ ...c, tips: [], prNotes: "", actions: [] }));

const emptyData = () => ({
  version: 2,
  members: [],
  cycles: DEFAULT_CYCLES.map((c) => ({ ...c, tips: [], actions: [] })),
  recognitions: [],
  weeks: [],
});

// ---------- Helpers ----------
const uid = () => Math.random().toString(36).slice(2, 10);

const daysSince = (iso) => {
  if (!iso) return null;
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d)) return null;
  return Math.floor((Date.now() - d.getTime()) / 86400000);
};

const onboardingDay = (iso) => {
  const ds = daysSince(iso);
  return ds === null ? null : ds + 1; // start date = day 1
};

const fmtDate = (iso) => {
  if (!iso) return "—";
  const d = new Date(iso + "T00:00:00");
  if (isNaN(d)) return "—";
  return d.toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
};

const memberStatus = (m) => {
  const ds = daysSince(m.lastAttended);
  if (ds === null) return { key: "unknown", label: "No attendance recorded", fg: C.amber, bg: C.amberBg };
  if (ds > 60) return { key: "dormant", label: `Dormant — ${ds} days`, fg: C.red, bg: C.redBg };
  if (ds > 30) return { key: "nudge", label: `Needs a nudge — ${ds} days`, fg: C.amber, bg: C.amberBg };
  return { key: "ok", label: `On track — ${ds}d ago`, fg: C.green, bg: C.greenBg };
};

const currentCycleId = (cycles) => {
  const m = new Date().getMonth();
  const found = cycles.find((c) => c.months.includes(m));
  return found ? found.id : null;
};

const stageForDay = (day, plan = ONBOARDING_STAGES) => {
  if (day === null) return null;
  const total = plan.length ? plan[plan.length - 1].to : 100;
  if (day > total) return { done: true };
  return plan.find((s) => day >= s.from && day <= s.to) || plan[0];
};

// ---------- Small UI atoms ----------
function Badge({ fg, bg, children }) {
  return (
    <span
      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold whitespace-nowrap"
      style={{ color: fg, backgroundColor: bg }}
    >
      {children}
    </span>
  );
}

function SectionTitle({ children, sub }) {
  return (
    <div className="mb-5">
      <h2 className="text-2xl" style={{ fontFamily: SERIF, color: C.blue }}>{children}</h2>
      {sub && <p className="text-sm mt-1" style={{ color: "#5B6B73" }}>{sub}</p>}
      <div className="mt-2 h-0.5 w-12" style={{ backgroundColor: C.maroon }} />
    </div>
  );
}

function Card({ children, className = "", accent }) {
  return (
    <div
      className={`bg-white rounded-lg shadow-sm ${className}`}
      style={{
        border: `1px solid ${C.grayLine}`,
        borderLeft: accent ? `4px solid ${accent}` : `1px solid ${C.grayLine}`,
      }}
    >
      {children}
    </div>
  );
}

function Btn({ children, onClick, kind = "primary", className = "", title, disabled }) {
  const styles = {
    primary: { backgroundColor: C.blue, color: "white" },
    maroon: { backgroundColor: C.maroon, color: "white" },
    ghost: { backgroundColor: "white", color: C.blue, border: `1px solid ${C.grayLine}` },
    danger: { backgroundColor: "white", color: C.red, border: `1px solid ${C.red}` },
  };
  return (
    <button
      title={title}
      disabled={disabled}
      onClick={onClick}
      className={`px-3 py-1.5 rounded-md text-sm font-semibold transition-opacity hover:opacity-85 focus:outline-none focus:ring-2 disabled:opacity-40 ${className}`}
      style={styles[kind]}
    >
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label className="block text-sm">
      <span className="block mb-1 font-semibold" style={{ color: C.blueDeep }}>{label}</span>
      {children}
    </label>
  );
}

const inputCls =
  "w-full px-2.5 py-1.5 rounded-md text-sm bg-white focus:outline-none focus:ring-2";
const inputStyle = { border: `1px solid ${C.grayLine}`, color: C.ink };

// ---------- Welcome screen ----------
function Welcome({ onStartFresh, onLoadFile }) {
  const fileRef = useRef(null);
  return (
    <div className="min-h-screen flex items-center justify-center p-6" style={{ backgroundColor: C.blue }}>
      <div className="max-w-md w-full bg-white rounded-xl shadow-lg p-8 text-center"
        style={{ borderTop: `6px solid ${C.maroon}` }}>
        <div className="text-xs tracking-widest font-bold mb-2" style={{ color: C.maroon }}>
          VP EDUCATION
        </div>
        <h1 className="text-3xl mb-3" style={{ fontFamily: SERIF, color: C.blue }}>
          Members Dashboard
        </h1>
        <p className="text-sm mb-6" style={{ color: "#5B6B73" }}>
          Your data stays on this device. Load your saved JSON file to pick up
          where you left off, or start a fresh club record.
        </p>
        <div className="flex flex-col gap-3">
          <Btn onClick={() => fileRef.current?.click()}>Load JSON file</Btn>
          <Btn kind="ghost" onClick={onStartFresh}>Start fresh</Btn>
        </div>
        <input
          ref={fileRef}
          type="file"
          accept=".json,application/json"
          className="hidden"
          onChange={(e) => onLoadFile(e.target.files?.[0])}
        />
        <p className="text-xs mt-6" style={{ color: "#8A958F" }}>
          Remember to Export JSON before closing — nothing is saved automatically.
        </p>
        <a href={CLUB_LINK} target="_blank" rel="noopener noreferrer"
          className="inline-block mt-3 text-xs font-semibold underline" style={{ color: C.blue }}>
          toastmasterclub.org ↗
        </a>
      </div>
    </div>
  );
}

// ---------- Member form modal ----------
function MemberModal({ initial, onSave, onClose }) {
  const [m, setM] = useState(
    initial || {
      id: uid(), name: "", path: "", level: 1, currentProject: "",
      lastAttended: "", totalMeetings: 0, roles: [], isNew: false,
      onboardingStart: "", notes: "",
    }
  );
  const [customRole, setCustomRole] = useState("");
  const set = (k, v) => setM((p) => ({ ...p, [k]: v }));
  const toggleRole = (r) =>
    set("roles", m.roles.includes(r) ? m.roles.filter((x) => x !== r) : [...m.roles, r]);

  return (
    <div className="fixed inset-0 z-50 flex items-start sm:items-center justify-center p-3 sm:p-6 overflow-y-auto"
      style={{ backgroundColor: "rgba(0,49,77,0.55)" }}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-2xl my-6"
        style={{ borderTop: `5px solid ${C.maroon}` }}>
        <div className="px-5 py-4 flex items-center justify-between" style={{ borderBottom: `1px solid ${C.grayLine}` }}>
          <h3 className="text-lg" style={{ fontFamily: SERIF, color: C.blue }}>
            {initial ? "Edit member" : "Add member"}
          </h3>
          <button onClick={onClose} className="text-xl leading-none px-2" style={{ color: "#8A958F" }} aria-label="Close">×</button>
        </div>
        <div className="p-5 grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="Name">
            <input className={inputCls} style={inputStyle} value={m.name}
              onChange={(e) => set("name", e.target.value)} placeholder="Full name" />
          </Field>
          <Field label="Pathways path">
            <select className={inputCls} style={inputStyle} value={m.path}
              onChange={(e) => set("path", e.target.value)}>
              <option value="">— Not selected —</option>
              {PATHS.map((p) => <option key={p}>{p}</option>)}
            </select>
          </Field>
          <Field label="Current level (1–5)">
            <select className={inputCls} style={inputStyle} value={m.level}
              onChange={(e) => set("level", Number(e.target.value))}>
              {[1, 2, 3, 4, 5].map((l) => <option key={l} value={l}>Level {l}</option>)}
            </select>
          </Field>
          <Field label="Current project">
            <input className={inputCls} style={inputStyle} value={m.currentProject}
              onChange={(e) => set("currentProject", e.target.value)} placeholder="e.g. Researching and Presenting" />
          </Field>
          <Field label="Last attended meeting">
            <input type="date" className={inputCls} style={inputStyle} value={m.lastAttended}
              onChange={(e) => set("lastAttended", e.target.value)} />
          </Field>
          <Field label="Total meetings attended">
            <input type="number" min="0" className={inputCls} style={inputStyle} value={m.totalMeetings}
              onChange={(e) => set("totalMeetings", Math.max(0, Number(e.target.value)))} />
          </Field>

          <div className="sm:col-span-2">
            <span className="block mb-1 text-sm font-semibold" style={{ color: C.blueDeep }}>Roles completed</span>
            <div className="flex flex-wrap gap-2">
              {[...new Set([...COMMON_ROLES, ...m.roles])].map((r) => {
                const on = m.roles.includes(r);
                return (
                  <button key={r} onClick={() => toggleRole(r)}
                    className="px-2.5 py-1 rounded-full text-xs font-semibold"
                    style={{
                      backgroundColor: on ? C.blue : "white",
                      color: on ? "white" : C.blue,
                      border: `1px solid ${on ? C.blue : C.grayLine}`,
                    }}>
                    {r}
                  </button>
                );
              })}
            </div>
            <div className="flex gap-2 mt-2">
              <input className={inputCls} style={inputStyle} value={customRole}
                onChange={(e) => setCustomRole(e.target.value)} placeholder="Add another role…" />
              <Btn kind="ghost" onClick={() => {
                const r = customRole.trim();
                if (r && !m.roles.includes(r)) set("roles", [...m.roles, r]);
                setCustomRole("");
              }}>Add</Btn>
            </div>
          </div>

          <div className="sm:col-span-2 flex flex-wrap items-end gap-4 p-3 rounded-md"
            style={{ backgroundColor: m.isNew ? C.amberBg : C.paper, border: `1px dashed ${C.grayLine}` }}>
            <label className="flex items-center gap-2 text-sm font-semibold" style={{ color: C.blueDeep }}>
              <input type="checkbox" checked={m.isNew} onChange={(e) => set("isNew", e.target.checked)} />
              New member (100-day plan)
            </label>
            {m.isNew && (
              <Field label="Onboarding start date">
                <input type="date" className={inputCls} style={inputStyle} value={m.onboardingStart}
                  onChange={(e) => set("onboardingStart", e.target.value)} />
              </Field>
            )}
          </div>

          <div className="sm:col-span-2">
            <Field label="Notes">
              <textarea rows={3} className={inputCls} style={inputStyle} value={m.notes}
                onChange={(e) => set("notes", e.target.value)}
                placeholder="Mentoring pairings, goals, things to remember…" />
            </Field>
          </div>
        </div>
        <div className="px-5 py-4 flex justify-end gap-2" style={{ borderTop: `1px solid ${C.grayLine}` }}>
          <Btn kind="ghost" onClick={onClose}>Cancel</Btn>
          <Btn onClick={() => { if (m.name.trim()) onSave(m); }} disabled={!m.name.trim()}>
            Save member
          </Btn>
        </div>
      </div>
    </div>
  );
}

// ---------- Onboarding progress bar ----------
function OnboardingBar({ startISO, compact, stages }) {
  const plan = stages && stages.length ? stages : ONBOARDING_STAGES;
  const total = plan[plan.length - 1].to;
  const day = onboardingDay(startISO);
  if (day === null) {
    return <p className="text-xs" style={{ color: C.amber }}>Set an onboarding start date to track progress.</p>;
  }
  const pct = Math.min(100, Math.max(0, (day / total) * 100));
  const stage = stageForDay(day, plan);
  return (
    <div>
      <div className="flex justify-between text-xs mb-1">
        <span className="font-semibold" style={{ color: C.blueDeep }}>
          {day > total ? "Plan complete 🎉" : `Day ${day} of ${total} — ${stage.label}`}
        </span>
        <span style={{ color: "#5B6B73" }}>{Math.round(pct)}%</span>
      </div>
      <div className="h-2.5 rounded-full overflow-hidden" style={{ backgroundColor: C.grayLine }}>
        <div className="h-full rounded-full" style={{ width: `${pct}%`, backgroundColor: day > total ? C.green : C.maroon }} />
      </div>
      {!compact && (
        <ol className="mt-3 grid grid-cols-1 sm:grid-cols-2 gap-1.5">
          {plan.map((s) => {
            const done = day > s.to;
            const active = day >= s.from && day <= s.to;
            return (
              <li key={`${s.from}-${s.label}`} className="flex items-center gap-2 text-xs px-2 py-1 rounded"
                style={{
                  backgroundColor: active ? C.gold : done ? C.greenBg : "transparent",
                  color: active ? C.blueDeep : done ? C.green : "#5B6B73",
                  fontWeight: active ? 700 : 500,
                }}>
                <span>{done ? "✓" : active ? "▶" : "○"}</span>
                <span>Days {s.from}–{s.to}: {s.label}</span>
              </li>
            );
          })}
        </ol>
      )}
    </div>
  );
}

// ---------- Custom 100-day plan editor ----------
function PlanModal({ member, onSave, onClose }) {
  const [stages, setStages] = useState(
    (member.customPlan && member.customPlan.length ? member.customPlan : ONBOARDING_STAGES).map((s) => ({ ...s }))
  );
  const set = (i, k, v) =>
    setStages((p) => p.map((s, idx) => (idx === i ? { ...s, [k]: k === "label" ? v : Math.max(1, Number(v) || 1) } : s)));
  const remove = (i) => setStages((p) => p.filter((_, idx) => idx !== i));
  const add = () => {
    const last = stages[stages.length - 1];
    const from = last ? last.to + 1 : 1;
    setStages((p) => [...p, { from, to: from + 13, label: "New stage" }]);
  };
  const valid = stages.length > 0 && stages.every((s, i) =>
    s.label.trim() && s.from <= s.to && (i === 0 || s.from > stages[i - 1].to));

  return (
    <div className="fixed inset-0 z-50 flex items-start sm:items-center justify-center p-3 sm:p-6 overflow-y-auto"
      style={{ backgroundColor: "rgba(0,49,77,0.55)" }}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-xl my-6" style={{ borderTop: `5px solid ${C.maroon}` }}>
        <div className="px-5 py-4 flex items-center justify-between" style={{ borderBottom: `1px solid ${C.grayLine}` }}>
          <h3 className="text-lg" style={{ fontFamily: SERIF, color: C.blue }}>
            Onboarding plan — {member.name}
          </h3>
          <button onClick={onClose} className="text-xl leading-none px-2" style={{ color: "#8A958F" }} aria-label="Close">×</button>
        </div>
        <div className="p-5 space-y-2">
          <div className="grid grid-cols-12 gap-2 text-xs font-bold" style={{ color: C.blueDeep }}>
            <span className="col-span-2">From day</span><span className="col-span-2">To day</span>
            <span className="col-span-7">Stage</span><span />
          </div>
          {stages.map((s, i) => (
            <div key={i} className="grid grid-cols-12 gap-2 items-center">
              <input type="number" min="1" className={`${inputCls} col-span-2`} style={inputStyle}
                value={s.from} onChange={(e) => set(i, "from", e.target.value)} />
              <input type="number" min="1" className={`${inputCls} col-span-2`} style={inputStyle}
                value={s.to} onChange={(e) => set(i, "to", e.target.value)} />
              <input className={`${inputCls} col-span-7`} style={inputStyle}
                value={s.label} onChange={(e) => set(i, "label", e.target.value)} />
              <button className="text-sm" style={{ color: C.red }} onClick={() => remove(i)} aria-label="Remove stage">✕</button>
            </div>
          ))}
          {!valid && (
            <p className="text-xs" style={{ color: C.red }}>
              Each stage needs a name, and day ranges must be in order without overlapping.
            </p>
          )}
          <div className="flex gap-2 pt-2">
            <Btn kind="ghost" onClick={add}>+ Add stage</Btn>
            <Btn kind="ghost" onClick={() => setStages(ONBOARDING_STAGES.map((s) => ({ ...s })))}>
              Reset to standard plan
            </Btn>
          </div>
        </div>
        <div className="px-5 py-4 flex justify-end gap-2" style={{ borderTop: `1px solid ${C.grayLine}` }}>
          <Btn kind="ghost" onClick={onClose}>Cancel</Btn>
          <Btn disabled={!valid} onClick={() => onSave(stages)}>Save plan</Btn>
        </div>
      </div>
    </div>
  );
}

// ---------- Views ----------
function HomeView({ data, go }) {
  const dormant = data.members.filter((m) => memberStatus(m).key === "dormant");
  const newbies = data.members.filter((m) => m.isNew);
  const activeCycle = data.cycles.find((c) => c.id === currentCycleId(data.cycles));
  const unposted = data.recognitions.filter((r) => !r.posted);
  const openActions = (activeCycle?.actions || []).filter((a) => !a.done);
  const latestWeek = (data.weeks || [])[0];
  const openWeekTasks = latestWeek ? latestWeek.tasks.filter((t) => !t.done) : [];

  const stats = [
    { label: "Active members", value: data.members.length, onClick: () => go("members") },
    { label: "Dormant (60+ days)", value: dormant.length, color: dormant.length ? C.red : C.green, onClick: () => go("members", "dormant") },
    { label: "In 100-day plan", value: newbies.length, onClick: () => go("onboarding") },
    { label: "Unposted recognitions", value: unposted.length, color: unposted.length ? C.amber : C.green, onClick: () => go("recognition") },
  ];

  return (
    <div>
      <SectionTitle sub={`Today is ${new Date().toLocaleDateString(undefined, { weekday: "long", day: "numeric", month: "long", year: "numeric" })}`}>
        At a glance
      </SectionTitle>

      {/* North Star */}
      <div className="mb-6 rounded-lg p-5"
        style={{ backgroundColor: C.blue, borderLeft: `5px solid ${C.gold}` }}>
        <div className="text-xs tracking-widest font-bold mb-2" style={{ color: C.gold }}>★ NORTH STAR</div>
        <ol className="space-y-2">
          <li className="flex gap-3 text-sm text-white">
            <span className="font-bold" style={{ color: C.gold, fontFamily: SERIF }}>1</span>
            <span><span className="font-semibold">Every meeting feels worth attending</span> — members get something out of it each time.</span>
          </li>
          <li className="flex gap-3 text-sm text-white">
            <span className="font-bold" style={{ color: C.gold, fontFamily: SERIF }}>2</span>
            <span><span className="font-semibold">Members actively progress through Pathways</span> — not just showing up, but moving forward.</span>
          </li>
        </ol>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-6">
        {stats.map((s) => (
          <button key={s.label} onClick={s.onClick} className="text-left">
            <Card className="p-4 h-full hover:shadow-md transition-shadow">
              <div className="text-3xl font-bold" style={{ fontFamily: SERIF, color: s.color || C.blue }}>
                {s.value}
              </div>
              <div className="text-xs mt-1 font-semibold" style={{ color: "#5B6B73" }}>{s.label}</div>
            </Card>
          </button>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <Card className="p-4" accent={C.gold}>
          <h3 className="font-bold text-sm mb-1" style={{ color: C.blueDeep }}>Current theme cycle</h3>
          {activeCycle ? (
            <>
              <div className="text-xl" style={{ fontFamily: SERIF, color: C.maroon }}>
                {activeCycle.name} <span className="text-sm" style={{ color: "#5B6B73" }}>({activeCycle.span})</span>
              </div>
              <div className="text-xs mt-2" style={{ color: "#5B6B73" }}>
                {activeCycle.tips.filter((t) => !t.posted).length} tip(s) still to post ·{" "}
                {openActions.length} open action(s)
              </div>
              <div className="mt-3"><Btn kind="ghost" onClick={() => go("cycles")}>Open cycle</Btn></div>
            </>
          ) : <p className="text-sm" style={{ color: "#5B6B73" }}>No cycle matches this month.</p>}
        </Card>

        <Card className="p-4" accent={openActions.length + openWeekTasks.length ? C.amber : C.green}>
          <h3 className="font-bold text-sm mb-2" style={{ color: C.blueDeep }}>Upcoming actions & reminders</h3>
          {openActions.length === 0 && openWeekTasks.length === 0 ? (
            <p className="text-sm" style={{ color: C.green }}>All clear — no open actions in the current cycle or week.</p>
          ) : (
            <ul className="space-y-1.5">
              {openWeekTasks.map((t) => (
                <li key={t.id} className="text-sm flex gap-2" style={{ color: C.ink }}>
                  <span style={{ color: C.gold }}>◆</span>{t.text}
                  <button className="text-xs underline ml-1" style={{ color: C.blue }} onClick={() => go("weekly")}>
                    {latestWeek.label}
                  </button>
                </li>
              ))}
              {openActions.slice(0, 5).map((a) => (
                <li key={a.id} className="text-sm flex gap-2" style={{ color: C.ink }}>
                  <span style={{ color: C.maroon }}>▸</span>{a.text}
                </li>
              ))}
              {openActions.length > 5 && (
                <li className="text-xs" style={{ color: "#5B6B73" }}>…and {openActions.length - 5} more in the cycle view.</li>
              )}
            </ul>
          )}
        </Card>

        {dormant.length > 0 && (
          <Card className="p-4 lg:col-span-2" accent={C.red}>
            <h3 className="font-bold text-sm mb-2" style={{ color: C.red }}>Members to reach out to</h3>
            <div className="flex flex-wrap gap-2">
              {dormant.map((m) => (
                <Badge key={m.id} fg={C.red} bg={C.redBg}>
                  {m.name} · {daysSince(m.lastAttended)}d
                </Badge>
              ))}
            </div>
          </Card>
        )}
      </div>
    </div>
  );
}

function MembersView({ data, setData, initialFilter }) {
  const [filter, setFilter] = useState(initialFilter || "all");
  const [search, setSearch] = useState("");
  const [editing, setEditing] = useState(null); // member object or "new"

  const filtered = useMemo(() => {
    return data.members
      .filter((m) => {
        const st = memberStatus(m).key;
        if (filter === "dormant" && st !== "dormant") return false;
        if (filter === "active" && st === "dormant") return false;
        if (filter === "new" && !m.isNew) return false;
        if (search && !m.name.toLowerCase().includes(search.toLowerCase())) return false;
        return true;
      })
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [data.members, filter, search]);

  const saveMember = (m) => {
    setData((d) => {
      const exists = d.members.some((x) => x.id === m.id);
      return { ...d, members: exists ? d.members.map((x) => (x.id === m.id ? m : x)) : [...d.members, m] };
    });
    setEditing(null);
  };

  const deleteMember = (id) => {
    if (window.confirm("Remove this member from the dashboard?")) {
      setData((d) => ({ ...d, members: d.members.filter((m) => m.id !== id) }));
    }
  };

  const filters = [
    ["all", "All"], ["active", "Active"], ["dormant", "Dormant"], ["new", "New members"],
  ];

  return (
    <div>
      <SectionTitle sub="Pathways progress, attendance, and roles for every member.">Members</SectionTitle>

      <div className="flex flex-wrap items-center gap-2 mb-4">
        {filters.map(([k, label]) => (
          <button key={k} onClick={() => setFilter(k)}
            className="px-3 py-1.5 rounded-full text-xs font-semibold"
            style={{
              backgroundColor: filter === k ? C.blue : "white",
              color: filter === k ? "white" : C.blue,
              border: `1px solid ${filter === k ? C.blue : C.grayLine}`,
            }}>
            {label}
          </button>
        ))}
        <input className={`${inputCls} max-w-xs ml-auto`} style={inputStyle}
          placeholder="Search by name…" value={search} onChange={(e) => setSearch(e.target.value)} />
        <Btn kind="maroon" onClick={() => setEditing("new")}>+ Add member</Btn>
      </div>

      {filtered.length === 0 ? (
        <Card className="p-8 text-center">
          <p className="text-sm" style={{ color: "#5B6B73" }}>
            {data.members.length === 0
              ? "No members yet. Add your first member to start tracking."
              : "No members match this filter."}
          </p>
        </Card>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-3">
          {filtered.map((m) => {
            const st = memberStatus(m);
            return (
              <Card key={m.id} className="p-4 flex flex-col gap-2"
                accent={st.key === "dormant" ? C.red : st.key === "ok" ? C.green : C.amber}>
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <div className="font-bold" style={{ color: C.blueDeep }}>{m.name}</div>
                    <div className="text-xs" style={{ color: "#5B6B73" }}>
                      {m.path ? `${m.path} · Level ${m.level}` : (
                        <span className="font-semibold" style={{ color: C.amber }}>⚠ No Pathways path selected</span>
                      )}
                    </div>
                  </div>
                  <div className="flex flex-col items-end gap-1">
                    <Badge fg={st.fg} bg={st.bg}>{st.key === "dormant" ? "● DORMANT" : st.label.split(" — ")[0]}</Badge>
                    {!m.path && <Badge fg={C.amber} bg={C.amberBg}>No path</Badge>}
                  </div>
                </div>

                <div className="text-xs space-y-1" style={{ color: C.ink }}>
                  <div><span className="font-semibold">Project:</span> {m.currentProject || "—"}</div>
                  <div><span className="font-semibold">Last attended:</span> {fmtDate(m.lastAttended)}
                    {st.key === "dormant" && <span style={{ color: C.red }}> ({daysSince(m.lastAttended)} days ago)</span>}
                  </div>
                  <div><span className="font-semibold">Meetings:</span> {m.totalMeetings}</div>
                </div>

                {m.roles.length > 0 && (
                  <div className="flex flex-wrap gap-1">
                    {m.roles.map((r) => (
                      <span key={r} className="px-1.5 py-0.5 rounded text-xs"
                        style={{ backgroundColor: C.paper, color: C.blueDeep, border: `1px solid ${C.grayLine}` }}>
                        {r}
                      </span>
                    ))}
                  </div>
                )}

                {m.isNew && (
                  <div className="mt-1 p-2 rounded" style={{ backgroundColor: C.paper }}>
                    <div className="text-xs font-bold mb-1" style={{ color: C.maroon }}>100-day plan</div>
                    <OnboardingBar startISO={m.onboardingStart} stages={m.customPlan} compact />
                  </div>
                )}

                {m.notes && <p className="text-xs italic" style={{ color: "#5B6B73" }}>“{m.notes}”</p>}

                <div className="flex gap-2 mt-auto pt-2">
                  <Btn kind="ghost" onClick={() => setEditing(m)}>Edit</Btn>
                  <Btn kind="danger" onClick={() => deleteMember(m.id)}>Remove</Btn>
                </div>
              </Card>
            );
          })}
        </div>
      )}

      {editing && (
        <MemberModal
          initial={editing === "new" ? null : editing}
          onSave={saveMember}
          onClose={() => setEditing(null)}
        />
      )}
    </div>
  );
}

function OnboardingView({ data, setData }) {
  const newbies = data.members.filter((m) => m.isNew);
  const [editingPlan, setEditingPlan] = useState(null); // member object

  const savePlan = (stages) => {
    setData((d) => ({
      ...d,
      members: d.members.map((m) => (m.id === editingPlan.id ? { ...m, customPlan: stages } : m)),
    }));
    setEditingPlan(null);
  };

  return (
    <div>
      <SectionTitle sub="Each new member's journey from first visit to Icebreaker. Every plan can be tailored to the member.">
        100-Day Onboarding
      </SectionTitle>
      {newbies.length === 0 ? (
        <Card className="p-8 text-center">
          <p className="text-sm" style={{ color: "#5B6B73" }}>
            No members are marked as new. Toggle “New member” on a member record to start their 100-day plan.
          </p>
        </Card>
      ) : (
        <div className="space-y-4">
          {newbies.map((m) => (
            <Card key={m.id} className="p-4" accent={C.maroon}>
              <div className="flex flex-wrap items-baseline justify-between gap-2 mb-3">
                <div className="font-bold" style={{ color: C.blueDeep, fontFamily: SERIF }}>
                  {m.name}
                  {m.customPlan && m.customPlan.length > 0 && (
                    <span className="ml-2 align-middle"><Badge fg={C.maroon} bg="#F6E8EB">custom plan</Badge></span>
                  )}
                </div>
                <div className="flex items-center gap-3">
                  <span className="text-xs" style={{ color: "#5B6B73" }}>Started {fmtDate(m.onboardingStart)}</span>
                  <Btn kind="ghost" onClick={() => setEditingPlan(m)}>Edit plan</Btn>
                </div>
              </div>
              <OnboardingBar startISO={m.onboardingStart} stages={m.customPlan} />
            </Card>
          ))}
        </div>
      )}
      {editingPlan && (
        <PlanModal member={editingPlan} onSave={savePlan} onClose={() => setEditingPlan(null)} />
      )}
    </div>
  );
}

function CyclesView({ data, setData }) {
  const activeId = currentCycleId(data.cycles);
  const [open, setOpen] = useState(activeId);
  const [drafts, setDrafts] = useState({}); // {cycleId: {tip, action}}

  const setCycle = (id, fn) =>
    setData((d) => ({ ...d, cycles: d.cycles.map((c) => (c.id === id ? fn(c) : c)) }));

  const draft = (id, k) => drafts[id]?.[k] || "";
  const setDraft = (id, k, v) => setDrafts((p) => ({ ...p, [id]: { ...p[id], [k]: v } }));

  return (
    <div>
      <SectionTitle sub="Six-week themes across the Toastmasters year. The active cycle is highlighted in gold.">
        Theme Cycles
      </SectionTitle>
      <div className="space-y-3">
        {data.cycles.map((c) => {
          const isActive = c.id === activeId;
          const isOpen = open === c.id;
          const unpostedTips = c.tips.filter((t) => !t.posted).length;
          const openActions = c.actions.filter((a) => !a.done).length;
          return (
            <Card key={c.id} accent={isActive ? C.gold : undefined}>
              <button onClick={() => setOpen(isOpen ? null : c.id)}
                className="w-full px-4 py-3 flex flex-wrap items-center gap-2 text-left"
                style={{ backgroundColor: isActive ? "#FBF5DA" : "white", borderRadius: "0.5rem" }}>
                <span className="text-lg" style={{ fontFamily: SERIF, color: C.maroon }}>{c.name}</span>
                <span className="text-xs font-semibold" style={{ color: "#5B6B73" }}>{c.span}</span>
                {isActive && <Badge fg={C.blueDeep} bg={C.gold}>ACTIVE NOW</Badge>}
                <span className="ml-auto text-xs" style={{ color: "#5B6B73" }}>
                  {unpostedTips > 0 && <span style={{ color: C.amber }}>{unpostedTips} tip(s) to post · </span>}
                  {openActions > 0 && <span>{openActions} open action(s) · </span>}
                  {isOpen ? "▲" : "▼"}
                </span>
              </button>

              {isOpen && (
                <div className="px-4 pb-4 grid grid-cols-1 lg:grid-cols-3 gap-4"
                  style={{ borderTop: `1px solid ${C.grayLine}` }}>
                  {/* Tips & articles */}
                  <div className="pt-3">
                    <h4 className="text-xs font-bold tracking-wide mb-2" style={{ color: C.blueDeep }}>TIPS & ARTICLES TO POST</h4>
                    <ul className="space-y-1.5 mb-2">
                      {c.tips.map((t) => (
                        <li key={t.id} className="flex items-start gap-2 text-sm">
                          <input type="checkbox" checked={t.posted} className="mt-0.5"
                            onChange={() => setCycle(c.id, (cy) => ({
                              ...cy, tips: cy.tips.map((x) => x.id === t.id ? { ...x, posted: !x.posted } : x),
                            }))} />
                          <span style={{
                            color: t.posted ? "#8A958F" : C.ink,
                            textDecoration: t.posted ? "line-through" : "none",
                          }}>{t.text}</span>
                          {!t.posted && <Badge fg={C.amber} bg={C.amberBg}>to post</Badge>}
                          <button className="ml-auto text-xs" style={{ color: "#8A958F" }}
                            onClick={() => setCycle(c.id, (cy) => ({ ...cy, tips: cy.tips.filter((x) => x.id !== t.id) }))}
                            aria-label="Delete tip">✕</button>
                        </li>
                      ))}
                      {c.tips.length === 0 && <li className="text-xs" style={{ color: "#8A958F" }}>Nothing queued yet.</li>}
                    </ul>
                    <div className="flex gap-2">
                      <input className={inputCls} style={inputStyle} placeholder="Add a tip or article…"
                        value={draft(c.id, "tip")} onChange={(e) => setDraft(c.id, "tip", e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter" && draft(c.id, "tip").trim()) {
                            setCycle(c.id, (cy) => ({ ...cy, tips: [...cy.tips, { id: uid(), text: draft(c.id, "tip").trim(), posted: false }] }));
                            setDraft(c.id, "tip", "");
                          }
                        }} />
                      <Btn kind="ghost" onClick={() => {
                        if (draft(c.id, "tip").trim()) {
                          setCycle(c.id, (cy) => ({ ...cy, tips: [...cy.tips, { id: uid(), text: draft(c.id, "tip").trim(), posted: false }] }));
                          setDraft(c.id, "tip", "");
                        }
                      }}>Add</Btn>
                    </div>
                  </div>

                  {/* VP PR notes */}
                  <div className="pt-3">
                    <h4 className="text-xs font-bold tracking-wide mb-2" style={{ color: C.blueDeep }}>VP PR COLLABORATION</h4>
                    <textarea rows={6} className={inputCls} style={inputStyle}
                      placeholder="Joint posts, social media plans, promo ideas…"
                      value={c.prNotes}
                      onChange={(e) => setCycle(c.id, (cy) => ({ ...cy, prNotes: e.target.value }))} />
                  </div>

                  {/* Actions */}
                  <div className="pt-3">
                    <h4 className="text-xs font-bold tracking-wide mb-2" style={{ color: C.blueDeep }}>MEETING ACTION ITEMS</h4>
                    <ul className="space-y-1.5 mb-2">
                      {c.actions.map((a) => (
                        <li key={a.id} className="flex items-start gap-2 text-sm">
                          <input type="checkbox" checked={a.done} className="mt-0.5"
                            onChange={() => setCycle(c.id, (cy) => ({
                              ...cy, actions: cy.actions.map((x) => x.id === a.id ? { ...x, done: !x.done } : x),
                            }))} />
                          <span style={{
                            color: a.done ? "#8A958F" : C.ink,
                            textDecoration: a.done ? "line-through" : "none",
                          }}>{a.text}</span>
                          <button className="ml-auto text-xs" style={{ color: "#8A958F" }}
                            onClick={() => setCycle(c.id, (cy) => ({ ...cy, actions: cy.actions.filter((x) => x.id !== a.id) }))}
                            aria-label="Delete action">✕</button>
                        </li>
                      ))}
                      {c.actions.length === 0 && <li className="text-xs" style={{ color: "#8A958F" }}>No actions yet.</li>}
                    </ul>
                    <div className="flex gap-2">
                      <input className={inputCls} style={inputStyle} placeholder="Add an action item…"
                        value={draft(c.id, "action")} onChange={(e) => setDraft(c.id, "action", e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter" && draft(c.id, "action").trim()) {
                            setCycle(c.id, (cy) => ({ ...cy, actions: [...cy.actions, { id: uid(), text: draft(c.id, "action").trim(), done: false }] }));
                            setDraft(c.id, "action", "");
                          }
                        }} />
                      <Btn kind="ghost" onClick={() => {
                        if (draft(c.id, "action").trim()) {
                          setCycle(c.id, (cy) => ({ ...cy, actions: [...cy.actions, { id: uid(), text: draft(c.id, "action").trim(), done: false }] }));
                          setDraft(c.id, "action", "");
                        }
                      }}>Add</Btn>
                    </div>
                  </div>
                </div>
              )}
            </Card>
          );
        })}
      </div>
    </div>
  );
}

function RecognitionView({ data, setData }) {
  const [form, setForm] = useState({ type: RECOGNITION_TYPES[0], member: "", detail: "" });
  const unposted = data.recognitions.filter((r) => !r.posted);

  const add = () => {
    if (!form.member.trim()) return;
    setData((d) => ({
      ...d,
      recognitions: [...d.recognitions, { id: uid(), ...form, member: form.member.trim(), posted: false }],
    }));
    setForm({ type: RECOGNITION_TYPES[0], member: "", detail: "" });
  };

  const togglePosted = (id) =>
    setData((d) => ({ ...d, recognitions: d.recognitions.map((r) => r.id === id ? { ...r, posted: !r.posted } : r) }));

  const remove = (id) =>
    setData((d) => ({ ...d, recognitions: d.recognitions.filter((r) => r.id !== id) }));

  const clearAll = () => {
    if (window.confirm("Clear all recognitions for the next meeting? This can't be undone (export JSON first if you want a record).")) {
      setData((d) => ({ ...d, recognitions: [] }));
    }
  };

  const needsDetail = form.type === "First time in a role" || form.type === "Pathways level completion";

  return (
    <div>
      <SectionTitle sub="Capture wins at each meeting, post them to Game Changers, then clear the slate.">
        Meeting Recognition
      </SectionTitle>

      {unposted.length > 0 && (
        <div className="mb-4 px-4 py-3 rounded-lg flex items-center gap-3"
          style={{ backgroundColor: C.amberBg, border: `1px solid ${C.amber}` }}>
          <span className="text-xl">⚠️</span>
          <span className="text-sm font-semibold" style={{ color: C.amber }}>
            {unposted.length} recognition{unposted.length > 1 ? "s" : ""} not yet posted to Game Changers.
          </span>
        </div>
      )}

      <Card className="p-4 mb-5" accent={C.maroon}>
        <h3 className="text-sm font-bold mb-3" style={{ color: C.blueDeep }}>Add recognition</h3>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 items-end">
          <Field label="Type">
            <select className={inputCls} style={inputStyle} value={form.type}
              onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}>
              {RECOGNITION_TYPES.map((t) => <option key={t}>{t}</option>)}
            </select>
          </Field>
          <Field label="Member">
            {data.members.length > 0 ? (
              <select className={inputCls} style={inputStyle} value={form.member}
                onChange={(e) => setForm((f) => ({ ...f, member: e.target.value }))}>
                <option value="">Choose member…</option>
                {data.members.map((m) => <option key={m.id} value={m.name}>{m.name}</option>)}
                <option value="__guest__">Guest / other (type below)</option>
              </select>
            ) : (
              <input className={inputCls} style={inputStyle} value={form.member}
                onChange={(e) => setForm((f) => ({ ...f, member: e.target.value }))} placeholder="Member name" />
            )}
          </Field>
          <Field label={needsDetail ? (form.type === "First time in a role" ? "Which role?" : "Which level / path?") : "Detail (optional)"}>
            <input className={inputCls} style={inputStyle} value={form.detail}
              onChange={(e) => setForm((f) => ({ ...f, detail: e.target.value }))}
              placeholder={form.type === "First time in a role" ? "e.g. Grammarian" : "e.g. Level 2 — Presentation Mastery"} />
          </Field>
        </div>
        {form.member === "__guest__" && (
          <input className={`${inputCls} mt-2 max-w-xs`} style={inputStyle} placeholder="Name"
            onChange={(e) => setForm((f) => ({ ...f, member: e.target.value }))} />
        )}
        <div className="mt-3"><Btn kind="maroon" onClick={add} disabled={!form.member.trim() || form.member === "__guest__"}>Add recognition</Btn></div>
      </Card>

      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-bold" style={{ color: C.blueDeep }}>
          This meeting cycle ({data.recognitions.length})
        </h3>
        <Btn kind="danger" onClick={clearAll} disabled={data.recognitions.length === 0}>
          Clear all for next meeting
        </Btn>
      </div>

      {data.recognitions.length === 0 ? (
        <Card className="p-8 text-center">
          <p className="text-sm" style={{ color: "#5B6B73" }}>A clean slate. Add recognitions as the meeting happens.</p>
        </Card>
      ) : (
        <div className="space-y-2">
          {data.recognitions.map((r) => (
            <Card key={r.id} className="px-4 py-3 flex flex-wrap items-center gap-3"
              accent={r.posted ? C.green : C.amber}>
              <div className="min-w-0">
                <div className="text-sm font-bold" style={{ color: C.blueDeep }}>
                  {r.member} <span className="font-normal" style={{ color: "#5B6B73" }}>— {r.type}</span>
                </div>
                {r.detail && <div className="text-xs" style={{ color: "#5B6B73" }}>{r.detail}</div>}
              </div>
              <div className="ml-auto flex items-center gap-3">
                {r.posted
                  ? <Badge fg={C.green} bg={C.greenBg}>✓ Posted</Badge>
                  : <Badge fg={C.amber} bg={C.amberBg}>Not posted</Badge>}
                <label className="flex items-center gap-1.5 text-xs font-semibold" style={{ color: C.blueDeep }}>
                  <input type="checkbox" checked={r.posted} onChange={() => togglePosted(r.id)} />
                  Game Changers
                </label>
                <button className="text-xs" style={{ color: "#8A958F" }} onClick={() => remove(r.id)} aria-label="Delete">✕</button>
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Weekly meeting checklist ----------
function WeeklyView({ data, setData }) {
  const weeks = data.weeks || [];
  const [drafts, setDrafts] = useState({}); // {weekId: text}

  const addWeek = () => {
    const today = new Date();
    const label = `Week of ${today.toLocaleDateString(undefined, { day: "numeric", month: "short" })}`;
    setData((d) => ({
      ...d,
      weeks: [
        {
          id: uid(),
          label,
          created: today.toISOString().slice(0, 10),
          tasks: DEFAULT_WEEK_TASKS.map((t) => ({ id: uid(), text: t, done: false })),
        },
        ...(d.weeks || []),
      ],
    }));
  };

  const setWeek = (id, fn) =>
    setData((d) => ({ ...d, weeks: d.weeks.map((w) => (w.id === id ? fn(w) : w)) }));

  const removeWeek = (id) => {
    if (window.confirm("Delete this week and its checklist?")) {
      setData((d) => ({ ...d, weeks: d.weeks.filter((w) => w.id !== id) }));
    }
  };

  return (
    <div>
      <SectionTitle sub="One checklist per meeting week: volunteers, notes, and Game Changers — nothing slips.">
        Weekly Checklist
      </SectionTitle>

      <div className="mb-4">
        <Btn kind="maroon" onClick={addWeek}>+ Start a new week</Btn>
      </div>

      {weeks.length === 0 ? (
        <Card className="p-8 text-center">
          <p className="text-sm" style={{ color: "#5B6B73" }}>
            No weeks yet. Start a new week and it comes pre-loaded with your three standing tasks:
            role volunteers, meeting notes, and Game Changers recognition.
          </p>
        </Card>
      ) : (
        <div className="space-y-4">
          {weeks.map((w, i) => {
            const open = w.tasks.filter((t) => !t.done).length;
            const isCurrent = i === 0;
            return (
              <Card key={w.id} className="p-4" accent={open === 0 ? C.green : isCurrent ? C.gold : C.amber}>
                <div className="flex flex-wrap items-center gap-2 mb-3">
                  <input
                    className="font-bold text-sm bg-transparent focus:outline-none focus:ring-2 rounded px-1"
                    style={{ color: C.blueDeep, fontFamily: SERIF, border: "1px solid transparent" }}
                    value={w.label}
                    onChange={(e) => setWeek(w.id, (wk) => ({ ...wk, label: e.target.value }))}
                  />
                  {isCurrent && <Badge fg={C.blueDeep} bg={C.gold}>CURRENT</Badge>}
                  {open === 0
                    ? <Badge fg={C.green} bg={C.greenBg}>✓ All done</Badge>
                    : <Badge fg={C.amber} bg={C.amberBg}>{open} open</Badge>}
                  <button className="ml-auto text-xs" style={{ color: "#8A958F" }}
                    onClick={() => removeWeek(w.id)}>Delete week</button>
                </div>
                <ul className="space-y-1.5 mb-2">
                  {w.tasks.map((t) => (
                    <li key={t.id} className="flex items-start gap-2 text-sm">
                      <input type="checkbox" checked={t.done} className="mt-0.5"
                        onChange={() => setWeek(w.id, (wk) => ({
                          ...wk, tasks: wk.tasks.map((x) => x.id === t.id ? { ...x, done: !x.done } : x),
                        }))} />
                      <span style={{
                        color: t.done ? "#8A958F" : C.ink,
                        textDecoration: t.done ? "line-through" : "none",
                      }}>{t.text}</span>
                      <button className="ml-auto text-xs" style={{ color: "#8A958F" }}
                        onClick={() => setWeek(w.id, (wk) => ({ ...wk, tasks: wk.tasks.filter((x) => x.id !== t.id) }))}
                        aria-label="Delete task">✕</button>
                    </li>
                  ))}
                </ul>
                <div className="flex gap-2">
                  <input className={inputCls} style={inputStyle} placeholder="Add a task for this week…"
                    value={drafts[w.id] || ""}
                    onChange={(e) => setDrafts((p) => ({ ...p, [w.id]: e.target.value }))}
                    onKeyDown={(e) => {
                      const text = (drafts[w.id] || "").trim();
                      if (e.key === "Enter" && text) {
                        setWeek(w.id, (wk) => ({ ...wk, tasks: [...wk.tasks, { id: uid(), text, done: false }] }));
                        setDrafts((p) => ({ ...p, [w.id]: "" }));
                      }
                    }} />
                  <Btn kind="ghost" onClick={() => {
                    const text = (drafts[w.id] || "").trim();
                    if (text) {
                      setWeek(w.id, (wk) => ({ ...wk, tasks: [...wk.tasks, { id: uid(), text, done: false }] }));
                      setDrafts((p) => ({ ...p, [w.id]: "" }));
                    }
                  }}>Add</Btn>
                </div>
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------- DTM tracker ----------
function DTMView({ data, setData }) {
  const tracked = data.members.filter((m) => m.dtm);
  const untracked = data.members.filter((m) => !m.dtm);
  const [pick, setPick] = useState("");

  const startTracking = () => {
    if (!pick) return;
    setData((d) => ({
      ...d,
      members: d.members.map((m) =>
        m.id === pick ? { ...m, dtm: DTM_REQUIREMENTS.map(() => false) } : m),
    }));
    setPick("");
  };

  const toggleReq = (memberId, idx) =>
    setData((d) => ({
      ...d,
      members: d.members.map((m) =>
        m.id === memberId ? { ...m, dtm: m.dtm.map((v, i) => (i === idx ? !v : v)) } : m),
    }));

  const stopTracking = (memberId) => {
    if (window.confirm("Remove this member from the DTM tracker? Their checklist will be lost.")) {
      setData((d) => ({
        ...d,
        members: d.members.map((m) => (m.id === memberId ? { ...m, dtm: undefined } : m)),
      }));
    }
  };

  return (
    <div>
      <SectionTitle sub="Track each member's road to Distinguished Toastmaster — paths, officer service, mentoring, and the DTM project.">
        DTM Tracker
      </SectionTitle>

      <Card className="p-4 mb-5" accent={C.maroon}>
        <h3 className="text-sm font-bold mb-2" style={{ color: C.blueDeep }}>Add a member to the DTM track</h3>
        <div className="flex flex-wrap gap-2">
          <select className={`${inputCls} max-w-xs`} style={inputStyle} value={pick}
            onChange={(e) => setPick(e.target.value)}>
            <option value="">Choose member…</option>
            {untracked.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
          </select>
          <Btn kind="maroon" onClick={startTracking} disabled={!pick}>Start tracking</Btn>
        </div>
        {data.members.length === 0 && (
          <p className="text-xs mt-2" style={{ color: "#5B6B73" }}>Add members first, then track their DTM journey here.</p>
        )}
      </Card>

      {tracked.length === 0 ? (
        <Card className="p-8 text-center">
          <p className="text-sm" style={{ color: "#5B6B73" }}>No one on the DTM track yet.</p>
        </Card>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {tracked.map((m) => {
            const done = m.dtm.filter(Boolean).length;
            const total = DTM_REQUIREMENTS.length;
            const pct = (done / total) * 100;
            return (
              <Card key={m.id} className="p-4" accent={done === total ? C.green : C.blue}>
                <div className="flex items-center justify-between gap-2 mb-2">
                  <div className="font-bold" style={{ color: C.blueDeep, fontFamily: SERIF }}>{m.name}</div>
                  {done === total
                    ? <Badge fg={C.green} bg={C.greenBg}>🏅 DTM complete</Badge>
                    : <Badge fg={C.blue} bg="#E5EEF4">{done}/{total}</Badge>}
                </div>
                <div className="h-2 rounded-full overflow-hidden mb-3" style={{ backgroundColor: C.grayLine }}>
                  <div className="h-full rounded-full"
                    style={{ width: `${pct}%`, backgroundColor: done === total ? C.green : C.blue }} />
                </div>
                <ul className="space-y-1.5">
                  {DTM_REQUIREMENTS.map((req, i) => (
                    <li key={req} className="flex items-start gap-2 text-sm">
                      <input type="checkbox" checked={m.dtm[i]} className="mt-0.5"
                        onChange={() => toggleReq(m.id, i)} />
                      <span style={{
                        color: m.dtm[i] ? "#8A958F" : C.ink,
                        textDecoration: m.dtm[i] ? "line-through" : "none",
                      }}>{req}</span>
                    </li>
                  ))}
                </ul>
                <div className="mt-3 text-right">
                  <button className="text-xs" style={{ color: "#8A958F" }} onClick={() => stopTracking(m.id)}>
                    Stop tracking
                  </button>
                </div>
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------- Main app ----------
const NAV = [
  { key: "home", label: "Home", icon: "⌂" },
  { key: "members", label: "Members", icon: "👥" },
  { key: "onboarding", label: "100-Day", icon: "🗓" },
  { key: "cycles", label: "Cycles", icon: "🔄" },
  { key: "weekly", label: "Weekly", icon: "☑" },
  { key: "recognition", label: "Recognition", short: "Awards", icon: "🏆" },
  { key: "dtm", label: "DTM", icon: "🎖" },
];

export default function VPEDashboard() {
  const [data, setData] = useState(null);
  const [view, setView] = useState("home");
  const [memberFilter, setMemberFilter] = useState("all");
  const [toast, setToast] = useState("");
  const fileRef = useRef(null);

  const notify = (msg) => { setToast(msg); setTimeout(() => setToast(""), 2600); };

  const go = (v, filter) => {
    if (filter) setMemberFilter(filter); else setMemberFilter("all");
    setView(v);
  };

  // --- file IO (all local) ---
  const loadFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        const base = emptyData();
        setData({
          ...base,
          ...parsed,
          members: Array.isArray(parsed.members) ? parsed.members : [],
          cycles: Array.isArray(parsed.cycles) && parsed.cycles.length === 6 ? parsed.cycles : base.cycles,
          recognitions: Array.isArray(parsed.recognitions) ? parsed.recognitions : [],
          weeks: Array.isArray(parsed.weeks) ? parsed.weeks : [],
        });
        notify("Data loaded from file.");
      } catch {
        window.alert("That file couldn't be read as valid JSON. Please check it and try again.");
      }
    };
    reader.readAsText(file);
  };

  const exportJSON = () => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `vpe-dashboard-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    notify("JSON exported — keep it somewhere safe.");
  };

  const exportExcel = () => {
    const rows = data.members.map((m) => ({
      Name: m.name,
      Path: m.path,
      Level: m.level,
      "Current project": m.currentProject,
      "Last attended": m.lastAttended || "",
      "Days since attended": daysSince(m.lastAttended) ?? "",
      Status: memberStatus(m).key,
      "Total meetings": m.totalMeetings,
      "Roles completed": m.roles.join(", "),
      "New member": m.isNew ? "Yes" : "No",
      "Onboarding start": m.onboardingStart || "",
      "Onboarding day": m.isNew ? (onboardingDay(m.onboardingStart) ?? "") : "",
      "DTM progress": m.dtm ? `${m.dtm.filter(Boolean).length}/${DTM_REQUIREMENTS.length}` : "",
      Notes: m.notes,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Members");
    XLSX.writeFile(wb, `club-members-${new Date().toISOString().slice(0, 10)}.xlsx`);
    notify("Excel file downloaded.");
  };

  if (!data) {
    return <Welcome onStartFresh={() => setData(emptyData())} onLoadFile={loadFile} />;
  }

  const unpostedCount = data.recognitions.filter((r) => !r.posted).length;
  const dormantCount = data.members.filter((m) => memberStatus(m).key === "dormant").length;

  return (
    <div className="min-h-screen flex flex-col md:flex-row" style={{ backgroundColor: C.paper, color: C.ink }}>
      {/* Sidebar (desktop) */}
      <aside className="hidden md:flex md:flex-col w-56 shrink-0 p-4 sticky top-0 h-screen"
        style={{ backgroundColor: C.blue }}>
        <div className="mb-6">
          <div className="text-xs tracking-widest font-bold" style={{ color: C.gold }}>VP EDUCATION</div>
          <div className="text-xl text-white" style={{ fontFamily: SERIF }}>Dashboard</div>
        </div>
        <nav className="flex flex-col gap-1">
          {NAV.map((n) => (
            <button key={n.key} onClick={() => go(n.key)}
              className="flex items-center gap-2 px-3 py-2 rounded-md text-sm font-semibold text-left"
              style={{
                backgroundColor: view === n.key ? "rgba(255,255,255,0.14)" : "transparent",
                color: view === n.key ? C.gold : "rgba(255,255,255,0.85)",
                borderLeft: view === n.key ? `3px solid ${C.gold}` : "3px solid transparent",
              }}>
              <span aria-hidden>{n.icon}</span>{n.label}
              {n.key === "recognition" && unpostedCount > 0 && (
                <span className="ml-auto px-1.5 rounded-full text-xs font-bold"
                  style={{ backgroundColor: C.gold, color: C.blueDeep }}>{unpostedCount}</span>
              )}
              {n.key === "members" && dormantCount > 0 && (
                <span className="ml-auto px-1.5 rounded-full text-xs font-bold"
                  style={{ backgroundColor: C.red, color: "white" }}>{dormantCount}</span>
              )}
            </button>
          ))}
        </nav>
        <div className="mt-auto flex flex-col gap-2">
          <a href={CLUB_LINK} target="_blank" rel="noopener noreferrer"
            className="flex items-center gap-2 px-3 py-2 rounded-md text-sm font-semibold"
            style={{ color: C.gold, border: `1px solid rgba(242,223,116,0.4)` }}>
            ↗ toastmasterclub.org
          </a>
          <Btn kind="maroon" onClick={exportJSON}>Export JSON</Btn>
          <Btn kind="ghost" onClick={exportExcel}>Export Excel</Btn>
          <button onClick={() => fileRef.current?.click()}
            className="text-xs underline mt-1" style={{ color: "rgba(255,255,255,0.7)" }}>
            Load a different JSON file
          </button>
          <p className="text-xs mt-2" style={{ color: "rgba(255,255,255,0.45)" }}>
            Data lives only on this device. Export before you close.
          </p>
        </div>
      </aside>

      {/* Mobile top bar */}
      <header className="md:hidden flex items-center justify-between px-4 py-3 sticky top-0 z-40"
        style={{ backgroundColor: C.blue }}>
        <div>
          <div className="text-xs tracking-widest font-bold" style={{ color: C.gold }}>VP EDUCATION</div>
          <div className="text-base text-white" style={{ fontFamily: SERIF }}>Dashboard</div>
        </div>
        <div className="flex gap-2">
          <Btn kind="maroon" onClick={exportJSON}>JSON</Btn>
          <Btn kind="ghost" onClick={exportExcel}>Excel</Btn>
        </div>
      </header>

      {/* Main */}
      <main className="flex-1 p-4 sm:p-6 lg:p-8 pb-24 md:pb-8 max-w-6xl">
        {view === "home" && <HomeView data={data} go={go} />}
        {view === "members" && <MembersView key={memberFilter} data={data} setData={setData} initialFilter={memberFilter} />}
        {view === "onboarding" && <OnboardingView data={data} setData={setData} />}
        {view === "cycles" && <CyclesView data={data} setData={setData} />}
        {view === "weekly" && <WeeklyView data={data} setData={setData} />}
        {view === "recognition" && <RecognitionView data={data} setData={setData} />}
        {view === "dtm" && <DTMView data={data} setData={setData} />}
      </main>

      {/* Mobile bottom nav */}
      <nav className="md:hidden fixed bottom-0 inset-x-0 z-40 flex"
        style={{ backgroundColor: C.blue, borderTop: `2px solid ${C.maroon}` }}>
        {NAV.map((n) => (
          <button key={n.key} onClick={() => go(n.key)}
            className="flex-1 py-2 flex flex-col items-center gap-0.5 text-xs font-semibold relative"
            style={{ color: view === n.key ? C.gold : "rgba(255,255,255,0.8)" }}>
            <span aria-hidden>{n.icon}</span>{n.short || n.label}
            {n.key === "recognition" && unpostedCount > 0 && (
              <span className="absolute top-1 right-1/4 w-2 h-2 rounded-full" style={{ backgroundColor: C.gold }} />
            )}
            {n.key === "members" && dormantCount > 0 && (
              <span className="absolute top-1 right-1/4 w-2 h-2 rounded-full" style={{ backgroundColor: C.red }} />
            )}
          </button>
        ))}
      </nav>

      {/* Hidden file input for re-loading */}
      <input ref={fileRef} type="file" accept=".json,application/json" className="hidden"
        onChange={(e) => { loadFile(e.target.files?.[0]); e.target.value = ""; }} />

      {/* Toast */}
      {toast && (
        <div className="fixed bottom-20 md:bottom-6 left-1/2 -translate-x-1/2 px-4 py-2 rounded-full text-sm font-semibold shadow-lg z-50"
          style={{ backgroundColor: C.blueDeep, color: "white" }}>
          {toast}
        </div>
      )}
    </div>
  );
}
