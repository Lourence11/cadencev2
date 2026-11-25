import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import os
import json
import html as html_lib
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Fill, PatternFill, Border, Alignment, Protection
from copy import copy


# ============================================================================
# UI LAYER (presentation only — no business logic lives here)
# React 18 + Tailwind are loaded from CDN and rendered inside Streamlit
# components. Native Streamlit widgets (uploaders, radio, buttons) are kept
# so the data flow is untouched; they are restyled with CSS.
# ============================================================================

GLOBAL_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, .stApp, [data-testid="stMarkdownContainer"], button, input, label {
    font-family: 'Inter', ui-sans-serif, system-ui, -apple-system, 'Segoe UI', sans-serif;
}
.stApp { background: #f4f6fb; }
.block-container { padding-top: 1.25rem; padding-bottom: 3rem; max-width: 1320px; }
header[data-testid="stHeader"] { background: transparent; }
#MainMenu, footer { visibility: hidden; }
iframe { border: 0; }

/* ---------- Section headers ---------- */
.ui-section { display: flex; align-items: center; gap: .85rem; margin: 1.6rem 0 .8rem; }
.ui-step {
    flex: none; width: 2.2rem; height: 2.2rem; border-radius: .7rem;
    display: grid; place-items: center; font-weight: 700; font-size: .9rem;
    color: #4338ca; background: #eef2ff; box-shadow: inset 0 0 0 1px #c7d2fe;
}
.ui-section h3 { margin: 0 !important; padding: 0 !important; font-size: 1.08rem !important; font-weight: 700 !important; color: #0f172a; }
.ui-section p { margin: .1rem 0 0 !important; color: #64748b; font-size: .85rem; }

/* ---------- File uploaders ---------- */
[data-testid="stFileUploader"] {
    background: #fff; border: 1px solid #e2e8f0; border-radius: 16px;
    padding: .9rem 1rem 1rem; box-shadow: 0 1px 2px rgba(15, 23, 42, .04);
    transition: box-shadow .15s, border-color .15s;
}
[data-testid="stFileUploader"]:hover { border-color: #c7d2fe; box-shadow: 0 6px 18px -10px rgba(79, 70, 229, .35); }
[data-testid="stFileUploader"] label p { font-weight: 600 !important; color: #0f172a; font-size: .9rem; }
[data-testid="stFileUploaderDropzone"] {
    background: #f8fafc; border: 1.5px dashed #cbd5e1; border-radius: 12px; transition: all .15s;
}
[data-testid="stFileUploaderDropzone"]:hover { border-color: #6366f1; background: #eef2ff; }

/* ---------- Tip card ---------- */
.ui-tip {
    height: 100%; background: linear-gradient(180deg, #ffffff, #f8faff); border: 1px dashed #c7d2fe;
    border-radius: 16px; padding: 1rem 1.1rem; color: #475569; font-size: .82rem;
}
.ui-tip b { color: #3730a3; font-size: .85rem; }
.ui-tip ul { margin: .45rem 0 0; padding-left: 1.05rem; }
.ui-tip li { margin: .2rem 0; }

/* ---------- Radio as pills ---------- */
div[role="radiogroup"] { gap: .5rem; flex-wrap: wrap; }
div[role="radiogroup"] > label, [data-testid="stRadioOption"] {
    background: #fff; border: 1px solid #e2e8f0; border-radius: 999px;
    padding: .5rem 1.05rem .5rem .75rem; margin: 0 !important; transition: all .15s;
}
div[role="radiogroup"] > label:hover, [data-testid="stRadioOption"]:hover { border-color: #a5b4fc; }
div[role="radiogroup"] > label:has(input:checked), [data-testid="stRadioOption"][data-selected="true"] {
    background: #eef2ff; border-color: #6366f1; box-shadow: 0 0 0 3px rgba(99, 102, 241, .15);
}
div[role="radiogroup"] > label:has(input:checked) p, [data-testid="stRadioOption"][data-selected="true"] p { color: #3730a3; font-weight: 600; }

/* ---------- Primary buttons (process + download) ---------- */
button[kind="primary"], button[data-testid="stBaseButton-primary"] {
    background: linear-gradient(135deg, #4f46e5, #0284c7) !important; border: 0 !important;
    border-radius: 12px !important; padding: .65rem 1.6rem !important; font-weight: 600 !important;
    box-shadow: 0 10px 22px -10px rgba(79, 70, 229, .7); transition: transform .12s, box-shadow .12s;
}
button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover {
    transform: translateY(-1px); box-shadow: 0 14px 26px -10px rgba(79, 70, 229, .8);
}

/* ---------- Alerts, expanders, bordered containers ---------- */
[data-testid="stAlert"] { border-radius: 12px; }
[data-testid="stExpander"] details { border-radius: 14px; background: #fff; border-color: #e2e8f0; }
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 16px; }
.ui-dl { display: flex; align-items: center; gap: .9rem; }
.ui-dl-icon {
    flex: none; width: 2.8rem; height: 2.8rem; border-radius: .8rem; display: grid; place-items: center;
    background: #ecfdf5; color: #059669; font-size: 1.3rem; box-shadow: inset 0 0 0 1px #a7f3d0;
}
.ui-dl-title { font-weight: 700; color: #0f172a; font-size: .95rem; }
.ui-dl-name { color: #64748b; font-size: .82rem; word-break: break-all; }
.ui-ready { color: #64748b; font-size: .85rem; padding-top: .55rem; }
.ui-ready b { color: #059669; }
.ui-ready i { color: #e11d48; font-style: normal; font-weight: 600; }

/* ---------- Sidebar ---------- */
[data-testid="stSidebar"] { background: #fff; border-right: 1px solid #e2e8f0; }
.ui-side-title { font-weight: 800; font-size: 1.05rem; color: #0f172a; margin: 0 0 .1rem; }
.ui-side-sub { color: #64748b; font-size: .8rem; margin-bottom: .6rem; }
</style>
"""

_REACT_SHELL = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
__TAILWIND__
<script>
  if (window.tailwind) {
    tailwind.config = { theme: { extend: { fontFamily: { sans: ['Inter', 'ui-sans-serif', 'system-ui', 'sans-serif'] } } } };
  }
</script>
__REACT_LIBS__
<style>
  html, body { margin: 0; padding: 0; background: transparent; font-family: Inter, ui-sans-serif, system-ui, sans-serif; -webkit-font-smoothing: antialiased; }
  ::-webkit-scrollbar { width: 6px; height: 6px; }
  ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 9px; }
</style>
</head>
<body>
<div id="root"></div>
<script>
  if (!window.React || !window.ReactDOM || !window.htm) {
    document.getElementById("root").innerHTML =
      '<div style="font:13px system-ui;color:#64748b;padding:12px;border:1px dashed #cbd5e1;border-radius:12px">' +
      'UI libraries could not load (CDN blocked?). The app still works &mdash; only this panel is affected.</div>';
  } else {
    const PROPS = __PROPS__;
    __SHARED__
    __APP__
    ReactDOM.createRoot(document.getElementById("root")).render(React.createElement(App, PROPS));
  }
</script>
</body>
</html>"""

SHARED_JS = r"""
const { useState } = React;
const h = htm.bind(React.createElement);

const ICON_PATHS = {
  check: "M5 13l4 4L19 7",
  file: "M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8zM14 3v5h5",
  calendar: "M8 3v4M16 3v4M3 10h18M5 5h14a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2z",
  clock: "M12 7v5l3 2M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0z",
  users: "M17 20v-2a4 4 0 0 0-4-4H7a4 4 0 0 0-4 4v2M10 10a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM21 20v-2a4 4 0 0 0-3-3.87M16 2.13a4 4 0 0 1 0 7.75",
  shield: "M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10zM9 12l2 2 4-4",
  lock: "M7 11V7a5 5 0 0 1 10 0v4M5 11h14v10H5z",
  logout: "M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9",
  promise: "M8 3v4M16 3v4M3 10h18M5 5h14a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2zM9 15l2 2 4-4",
  cash: "M3 7h18v10H3zM12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z",
  arrow: "M5 12h14M13 6l6 6-6 6",
  inbox: "M22 12h-6l-2 3h-4l-2-3H2M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z",
  merge: "M6 3v12M18 9a3 3 0 1 0 0-6 3 3 0 0 0 0 6zM6 21a3 3 0 1 0 0-6 3 3 0 0 0 0 6zM18 9a9 9 0 0 1-9 9",
};

function Icon({ name, className }) {
  return h`<svg className=${className || "h-4 w-4"} width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
    strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d=${ICON_PATHS[name] || ""} /></svg>`;
}

const TONES = {
  slate:   { soft: "bg-slate-100",  text: "text-slate-700",   strong: "bg-slate-400",   ring: "ring-slate-200" },
  violet:  { soft: "bg-violet-50",  text: "text-violet-700",  strong: "bg-violet-500",  ring: "ring-violet-200" },
  amber:   { soft: "bg-amber-50",   text: "text-amber-700",   strong: "bg-amber-500",   ring: "ring-amber-200" },
  rose:    { soft: "bg-rose-50",    text: "text-rose-700",    strong: "bg-rose-500",    ring: "ring-rose-200" },
  sky:     { soft: "bg-sky-50",     text: "text-sky-700",     strong: "bg-sky-500",     ring: "ring-sky-200" },
  emerald: { soft: "bg-emerald-50", text: "text-emerald-700", strong: "bg-emerald-500", ring: "ring-emerald-200" },
  indigo:  { soft: "bg-indigo-50",  text: "text-indigo-700",  strong: "bg-indigo-500",  ring: "ring-indigo-200" },
};
const tone = (t) => TONES[t] || TONES.slate;
const fmt = (n) => Number(n || 0).toLocaleString();
"""

HERO_JS = r"""
function Pill({ icon, label, value }) {
  return h`
    <div className="flex items-center gap-2.5 rounded-xl bg-white/10 px-3.5 py-2 ring-1 ring-inset ring-white/20 backdrop-blur">
      <${Icon} name=${icon} className="h-4 w-4 text-sky-100" />
      <div className="leading-tight">
        <div className="text-[10px] font-semibold uppercase tracking-wider text-indigo-100/80">${label}</div>
        <div className="text-sm font-semibold">${value}</div>
      </div>
    </div>`;
}

function Step({ index, label, detail, state }) {
  const done = state === "done", current = state === "current";
  const badge = done ? "bg-emerald-400 text-emerald-950"
    : current ? "bg-white text-indigo-700 ring-4 ring-white/25" : "bg-white/15 text-white/70";
  return h`
    <div className="flex min-w-0 items-center gap-3 px-5 py-3.5">
      <div className=${"grid h-8 w-8 flex-none place-items-center rounded-full text-xs font-bold " + badge}>
        ${done ? h`<${Icon} name="check" className="h-4 w-4" />` : index}
      </div>
      <div className="min-w-0 leading-tight">
        <div className=${"truncate text-sm font-semibold " + (done || current ? "text-white" : "text-white/70")}>${label}</div>
        <div className="truncate text-xs text-indigo-100/80">${detail}</div>
      </div>
    </div>`;
}

function App({ title, subtitle, dateLabel, monthLabel, steps }) {
  return h`
  <div className="p-0.5">
    <div className="relative overflow-hidden rounded-2xl bg-gradient-to-br from-indigo-800 via-indigo-600 to-sky-500 text-white shadow-xl shadow-indigo-900/10">
      <div className="pointer-events-none absolute -right-20 -top-24 h-72 w-72 rounded-full bg-white/10"></div>
      <div className="pointer-events-none absolute right-48 -bottom-28 h-56 w-56 rounded-full bg-sky-300/20 blur-2xl"></div>

      <div className="relative flex flex-wrap items-end justify-between gap-4 px-7 pb-5 pt-6">
        <div>
          <div className="inline-flex items-center gap-2 rounded-full bg-white/15 px-3 py-1 text-[11px] font-semibold uppercase tracking-wider">
            <span className="h-1.5 w-1.5 rounded-full bg-emerald-300"></span> PTP Status Updater
          </div>
          <h1 className="mt-3 text-[28px] font-extrabold leading-tight tracking-tight">${title}</h1>
          <p className="mt-1 text-sm text-indigo-100">${subtitle}</p>
          <div className="mt-3 flex flex-wrap gap-2 text-xs">
            <span className="rounded-md bg-white/10 px-2.5 py-1 ring-1 ring-inset ring-white/15">PTP Date must be today or later</span>
            <span className="rounded-md bg-white/10 px-2.5 py-1 ring-1 ring-inset ring-white/15">Claim Paid Date must be within ${monthLabel}</span>
          </div>
        </div>
        <div className="flex flex-wrap gap-2">
          <${Pill} icon="calendar" label="Today's date" value=${dateLabel} />
          <${Pill} icon="clock" label="Current month" value=${monthLabel} />
        </div>
      </div>

      <div className="relative grid grid-cols-4 divide-x divide-white/10 border-t border-white/15 bg-indigo-950/20">
        ${steps.map((s, i) => h`<${Step} key=${i} index=${i + 1} ...${s} />`)}
      </div>
    </div>
  </div>`;
}
"""

FILES_JS = r"""
function App({ files }) {
  return h`
  <div className="grid grid-cols-5 gap-2.5 p-0.5">
    ${files.map((f, i) => {
      const ok = !!f.name;
      const card = ok ? "bg-emerald-50/70 ring-emerald-200" : f.required ? "bg-white ring-rose-200" : "bg-white ring-slate-200";
      const badge = ok ? "bg-emerald-500 text-white" : f.required ? "bg-rose-50 text-rose-500" : "bg-slate-100 text-slate-400";
      const sub = ok ? "text-emerald-700" : f.required ? "text-rose-500" : "text-slate-400";
      return h`
        <div key=${i} className=${"flex min-w-0 items-center gap-3 rounded-xl px-3 py-2.5 ring-1 " + card}>
          <div className=${"grid h-8 w-8 flex-none place-items-center rounded-lg " + badge}>
            <${Icon} name=${ok ? "check" : "file"} />
          </div>
          <div className="min-w-0 leading-tight">
            <div className="flex items-center gap-1.5 text-[12px] font-semibold text-slate-800">
              <span className="truncate">${f.label}</span>
              ${f.required && h`<span className="rounded bg-slate-100 px-1 text-[9px] font-bold uppercase text-slate-500">Req</span>`}
            </div>
            <div className=${"truncate text-[11px] " + sub} title=${f.name || ""}>
              ${ok ? f.name + " · " + f.size : f.required ? "Missing" : "Optional · skipped"}
            </div>
          </div>
        </div>`;
    })}
  </div>`;
}
"""

STATS_JS = r"""
function StatCard({ label, value, icon, t, share }) {
  const c = tone(t);
  const w = share == null ? 100 : Math.min(100, share);
  return h`
    <div className="min-w-[160px] flex-1 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-slate-200">
      <div className="flex items-center justify-between gap-2">
        <span className="text-[10.5px] font-semibold uppercase leading-tight tracking-wide text-slate-500">${label}</span>
        <span className=${"grid h-7 w-7 flex-none place-items-center rounded-lg " + c.soft + " " + c.text}>
          <${Icon} name=${icon} className="h-3.5 w-3.5" />
        </span>
      </div>
      <div className="mt-2 text-[26px] font-bold tabular-nums tracking-tight text-slate-900">${fmt(value)}</div>
      <div className="mt-1 h-1 w-full overflow-hidden rounded-full bg-slate-100">
        <div className=${"h-full rounded-full " + c.strong} style=${{ width: w + "%" }}></div>
      </div>
      <div className="mt-1.5 text-[11px] text-slate-400">${share == null ? "All accounts" : share.toFixed(1) + "% of total"}</div>
    </div>`;
}

function App({ total, cards, dist }) {
  const sum = dist.reduce((a, d) => a + d.value, 0) || 1;
  return h`
  <div className="space-y-3 p-0.5">
    <div className="flex gap-3 overflow-x-auto pb-1">
      ${cards.map((c, i) => h`<${StatCard} key=${i} ...${c}
          share=${c.isTotal ? null : (total ? (c.value / total) * 100 : 0)} />`)}
    </div>
    <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-slate-200">
      <div className="mb-2.5 flex items-center justify-between">
        <span className="text-sm font-semibold text-slate-800">Final status distribution</span>
        <span className="text-xs text-slate-400">${fmt(total)} accounts</span>
      </div>
      <div className="flex h-3 w-full overflow-hidden rounded-full bg-slate-100">
        ${dist.filter((d) => d.value > 0).map((d, i) => h`
          <div key=${i} className=${tone(d.t).strong + " h-full border-r-2 border-white last:border-r-0"}
               style=${{ width: (d.value / sum) * 100 + "%" }} title=${d.label + ": " + fmt(d.value)}></div>`)}
      </div>
      <div className="mt-3 flex flex-wrap gap-x-5 gap-y-1.5">
        ${dist.map((d, i) => h`
          <div key=${i} className="flex items-center gap-1.5 text-xs text-slate-600">
            <span className=${"h-2 w-2 rounded-full " + tone(d.t).strong}></span>${d.label}
            <span className="font-semibold tabular-nums text-slate-900">${fmt(d.value)}</span>
          </div>`)}
      </div>
    </div>
  </div>`;
}
"""

SAMPLES_JS = r"""
function App({ groups }) {
  const avail = groups.filter((g) => g.count > 0);
  const [active, setActive] = useState(avail.length ? avail[0].key : null);

  if (!avail.length) {
    return h`
      <div className="m-0.5 flex flex-col items-center justify-center rounded-2xl bg-white py-10 text-center ring-1 ring-slate-200">
        <div className="grid h-11 w-11 place-items-center rounded-xl bg-slate-100 text-slate-400"><${Icon} name="inbox" className="h-5 w-5" /></div>
        <div className="mt-3 text-sm font-semibold text-slate-700">No accounts moved to ISSA CERT, LOCKED, PAYMENT, ACTIVE PTP or PULLED OUT</div>
        <div className="mt-1 text-xs text-slate-400">All statuses were kept as they were in the template.</div>
      </div>`;
  }

  const g = avail.find((x) => x.key === active) || avail[0];
  const c = tone(g.t);
  return h`
  <div className="m-0.5 overflow-hidden rounded-2xl bg-white shadow-sm ring-1 ring-slate-200">
    <div className="flex flex-wrap items-center gap-1.5 border-b border-slate-100 bg-slate-50/70 px-3 py-2.5">
      ${avail.map((x) => {
        const on = x.key === g.key, tc = tone(x.t);
        return h`
          <button key=${x.key} onClick=${() => setActive(x.key)}
            className=${"inline-flex items-center gap-2 rounded-lg px-3 py-1.5 text-[13px] font-medium transition " +
              (on ? "bg-white text-slate-900 shadow-sm ring-1 ring-slate-200" : "text-slate-500 hover:bg-white/70 hover:text-slate-800")}>
            <span className=${"h-2 w-2 rounded-full " + tc.strong}></span>${x.label}
            <span className=${"rounded-md px-1.5 py-0.5 text-[11px] font-semibold tabular-nums " +
              (on ? tc.soft + " " + tc.text : "bg-slate-100 text-slate-500")}>${fmt(x.count)}</span>
          </button>`;
      })}
    </div>
    <table className="w-full text-left text-sm">
      <thead>
        <tr className="text-[11px] uppercase tracking-wider text-slate-400">
          <th className="w-14 px-4 py-2.5 font-semibold">#</th>
          <th className="px-4 py-2.5 font-semibold">UID</th>
          <th className="px-4 py-2.5 font-semibold">Account Status</th>
        </tr>
      </thead>
      <tbody>
        ${g.rows.map((r, i) => h`
          <tr key=${i} className="border-t border-slate-100 hover:bg-slate-50">
            <td className="px-4 py-2 text-xs tabular-nums text-slate-400">${i + 1}</td>
            <td className="px-4 py-2 font-mono text-[13px] text-slate-800">${r.uid}</td>
            <td className="px-4 py-2">
              <span className=${"inline-flex rounded-full px-2.5 py-0.5 text-xs font-semibold ring-1 ring-inset " + c.soft + " " + c.text + " " + c.ring}>${r.status}</span>
            </td>
          </tr>`)}
      </tbody>
    </table>
    <div className="border-t border-slate-100 px-4 py-2.5 text-xs text-slate-400">
      Showing first ${g.rows.length} of ${fmt(g.count)} ${g.label} accounts
    </div>
  </div>`;
}
"""

SIDEBAR_JS = r"""
function Order({ items }) {
  return h`
    <ol className="space-y-0">
      ${items.map((it, i) => h`
        <li key=${i} className="flex gap-3">
          <div className="flex flex-col items-center">
            <div className=${"grid h-7 w-7 flex-none place-items-center rounded-full text-xs font-bold text-white " + tone(it.t).strong}>${i + 1}</div>
            ${i < items.length - 1 && h`<div className="my-1 w-px flex-1 bg-slate-200"></div>`}
          </div>
          <div className="mb-3 min-w-0 flex-1 rounded-xl bg-white p-3 ring-1 ring-slate-200">
            <div className="flex items-start justify-between gap-2">
              <span className="text-[13px] font-semibold leading-tight text-slate-800">${it.title}</span>
              ${it.result && h`<span className=${"flex-none rounded-md px-1.5 py-0.5 text-[10px] font-bold " + tone(it.t).soft + " " + tone(it.t).text}>${it.result}</span>`}
            </div>
            <ul className="mt-1.5 space-y-1">
              ${it.lines.map((l, j) => h`
                <li key=${j} className=${"flex gap-1.5 text-[11.5px] leading-snug " + (l.em ? "font-semibold text-slate-700" : "text-slate-500")}>
                  <span className="mt-1.5 h-1 w-1 flex-none rounded-full bg-slate-300"></span><span>${l.text}</span>
                </li>`)}
            </ul>
          </div>
        </li>`)}
    </ol>`;
}

function Flag({ label, v }) {
  const cls = v === "Yes" ? "bg-emerald-50 text-emerald-700" : v === "No" ? "bg-rose-50 text-rose-600" : "bg-slate-50 text-slate-400";
  return h`<span className=${"rounded-md px-1.5 py-0.5 font-medium " + cls}>${label}: ${v}</span>`;
}

function Rules({ rows }) {
  return h`
    <div className="space-y-2">
      ${rows.map((r, i) => h`
        <div key=${i} className="rounded-xl bg-white p-3 ring-1 ring-slate-200">
          <div className="flex flex-wrap items-center gap-1.5 text-[11px]">
            <span className="rounded-md bg-slate-100 px-1.5 py-0.5 font-semibold text-slate-700">${r.current}</span>
            <${Flag} label="PTP" v=${r.ptp} />
            <${Flag} label="KEPT" v=${r.kept} />
          </div>
          <div className="mt-2 flex items-center gap-1.5 text-[12px] font-semibold">
            <${Icon} name="arrow" className="h-3.5 w-3.5 text-slate-400" />
            <span className=${tone(r.t).text}>${r.final}</span>
          </div>
        </div>`)}
    </div>`;
}

function Features({ items }) {
  return h`
    <ul className="space-y-1.5">
      ${items.map((f, i) => h`
        <li key=${i} className="flex items-center gap-2.5 rounded-xl bg-white px-3 py-2 text-[12px] text-slate-700 ring-1 ring-slate-200">
          <span className="grid h-5 w-5 flex-none place-items-center rounded-full bg-emerald-50 text-emerald-600">
            <${Icon} name="check" className="h-3 w-3" />
          </span>${f}
        </li>`)}
    </ul>`;
}

function App({ order, rules, features }) {
  const [tab, setTab] = useState("order");
  const tabs = [["order", "Order"], ["rules", "Rules"], ["features", "Features"]];
  return h`
  <div className="m-0.5 flex h-[684px] flex-col overflow-hidden rounded-2xl bg-slate-50 ring-1 ring-slate-200">
    <div className="flex gap-1 border-b border-slate-200 bg-white p-1.5">
      ${tabs.map(([k, l]) => h`
        <button key=${k} onClick=${() => setTab(k)}
          className=${"flex-1 rounded-lg px-2 py-1.5 text-xs font-semibold transition " +
            (tab === k ? "bg-indigo-600 text-white shadow" : "text-slate-500 hover:bg-slate-100")}>${l}</button>`)}
    </div>
    <div className="flex-1 overflow-y-auto p-3">
      ${tab === "order" && h`<${Order} items=${order} />`}
      ${tab === "rules" && h`<${Rules} rows=${rules} />`}
      ${tab === "features" && h`<${Features} items=${features} />`}
    </div>
  </div>`;
}
"""


def inject_global_css():
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)


# UI libraries: loaded from CDN by default. For networks that block CDNs, put
# these files in a "ui_assets" folder next to this script and they are inlined.
UI_ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ui_assets")
UI_LIBS = [
    ("tailwind.js", "https://cdn.tailwindcss.com"),
    ("react.production.min.js", "https://unpkg.com/react@18.3.1/umd/react.production.min.js"),
    ("react-dom.production.min.js", "https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js"),
    ("htm.umd.js", "https://unpkg.com/htm@3.1.1/dist/htm.umd.js"),
]


@st.cache_resource
def _lib_tags():
    tags = []
    for filename, url in UI_LIBS:
        local_path = os.path.join(UI_ASSETS_DIR, filename)
        if os.path.isfile(local_path):
            with open(local_path, "r", encoding="utf-8") as f:
                tags.append("<script>" + f.read().replace("</script", "<\\/script") + "</script>")
        else:
            tags.append(f'<script crossorigin src="{url}"></script>')
    return tags


def render_react(app_js, props, height, scrolling=False):
    """Render a React (htm) + Tailwind component inside a Streamlit iframe."""
    props_json = json.dumps(props, default=str).replace("</", "<\\/")
    libs = _lib_tags()
    page = (
        _REACT_SHELL
        .replace("__TAILWIND__", libs[0])
        .replace("__REACT_LIBS__", "\n".join(libs[1:]))
        .replace("__SHARED__", SHARED_JS)
        .replace("__APP__", app_js)
        .replace("__PROPS__", props_json)
    )
    if hasattr(st, "iframe"):
        # Newer Streamlit (components.html is deprecated there)
        st.iframe(page, height=height)
    else:
        components.html(page, height=height, scrolling=scrolling)


def section_header(step, title, subtitle=""):
    st.markdown(
        f'<div class="ui-section"><span class="ui-step">{html_lib.escape(str(step))}</span>'
        f'<div><h3>{html_lib.escape(title)}</h3><p>{html_lib.escape(subtitle)}</p></div></div>',
        unsafe_allow_html=True,
    )


def _fmt_size(num_bytes):
    if num_bytes is None:
        return ""
    size = float(num_bytes)
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024 or unit == "GB":
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024


def _file_info(label, uploaded, required):
    return {
        "label": label,
        "required": required,
        "name": uploaded.name if uploaded else None,
        "size": _fmt_size(getattr(uploaded, "size", None)) if uploaded else "",
    }


def _sample_rows(df, n=10):
    return [
        {"uid": str(uid), "status": str(status)}
        for uid, status in df[["UID", "Account Status"]].head(n).itertuples(index=False)
    ]


SIDEBAR_ORDER = [
    {"title": "Active Accounts", "result": "PULLED OUT", "t": "rose", "lines": [
        {"text": "Check if UID exists in Active Accounts"},
        {"text": "If NOT: update to \"PULLED OUT\" (stops here)", "em": True},
    ]},
    {"title": "Locked Accounts", "result": "LOCKED", "t": "amber", "lines": [
        {"text": "Check if UID exists in the \"Old IC\" column of Locked Accounts"},
        {"text": "If YES: update to \"LOCKED\" (only if NOT PULLED OUT)", "em": True},
    ]},
    {"title": "ISSA CERT", "result": "ISSA CERT", "t": "violet", "lines": [
        {"text": "Check if UID exists in ISSA CERT List"},
        {"text": "If YES: update to \"ISSA CERT\" (only if NOT PULLED OUT and NOT LOCKED)", "em": True},
    ]},
    {"title": "PTP Lookup", "result": "ACTIVE PTP", "t": "sky", "lines": [
        {"text": "Only if current status is \"ACTIVE\""},
        {"text": "Only if PTP Date is TODAY or onward", "em": True},
        {"text": "If PTP Date is in the past → keep original status"},
        {"text": "Update to: \"ACTIVE PTP - mm/dd/yyyy\""},
    ]},
    {"title": "KEPT Lookup", "result": "PAYMENT", "t": "emerald", "lines": [
        {"text": "Only if current status is \"ACTIVE\" or \"ACTIVE PTP\""},
        {"text": "Only if Claim Paid Date is within current month", "em": True},
        {"text": "If not in current month → keep previous status"},
        {"text": "Update to: \"PAYMENT\""},
    ]},
    {"title": "MAPPED → DRR", "result": "APPEND", "t": "indigo", "lines": [
        {"text": "Read \"MAPPED\" sheet from Cadence Output"},
        {"text": "Read \"DRR\" sheet from Cadence Template"},
        {"text": "Append MAPPED data to DRR"},
    ]},
]

SIDEBAR_RULES = [
    {"current": "ACTIVE", "ptp": "Yes", "kept": "No", "final": "ACTIVE PTP - mm/dd/yyyy", "t": "sky"},
    {"current": "ACTIVE", "ptp": "Yes", "kept": "Yes", "final": "PAYMENT", "t": "emerald"},
    {"current": "ACTIVE", "ptp": "No", "kept": "No", "final": "ACTIVE (unchanged)", "t": "slate"},
    {"current": "ACTIVE", "ptp": "No", "kept": "Yes", "final": "PAYMENT", "t": "emerald"},
    {"current": "ACTIVE PTP", "ptp": "-", "kept": "Yes", "final": "PAYMENT", "t": "emerald"},
    {"current": "Other", "ptp": "-", "kept": "-", "final": "Unchanged", "t": "slate"},
    {"current": "Not in Active", "ptp": "-", "kept": "-", "final": "PULLED OUT", "t": "rose"},
]

SIDEBAR_FEATURES = [
    "ISSA CERT Status",
    "LOCKED Status (Old IC column search)",
    "PTP Lookup (Descending Order)",
    "PTP Date Verification (Today or onward only)",
    "KEPT Lookup (Overrides PTP)",
    "PAYMENT Date Verification (Within current month only)",
    "MAPPED → DRR (Append data)",
    "Active Accounts Filter",
    "PULLED OUT Status",
    "Original Formatting Preserved",
    "No Summary Sheets in Download",
]


# ============================================================================
# APP
# ============================================================================

def main():
    # Page Configuration
    st.set_page_config(page_title="PTP Status Updater", page_icon="📊", layout="wide")
    inject_global_css()

    # Hero is rendered later into this slot (so it can reflect upload state)
    hero_slot = st.container()

    # ============================================
    # FILE UPLOADS SECTION
    # ============================================
    section_header(1, "Upload Files", "Cadence Template and Cadence Output are required. The other lists are optional.")

    col1, col2, col3 = st.columns(3)

    with col1:
        cadence_template = st.file_uploader(
            "Upload Cadence Template *",
            type=['xlsx'],
            help="Upload the BPI PL Collection Effort Cadence TEMPLATE.xlsx"
        )

    with col2:
        cadence_output = st.file_uploader(
            "Upload Cadence Output *",
            type=['xlsx'],
            help="Upload the Cadence_Output TEMPLATE.xlsx"
        )

    with col3:
        active_accounts = st.file_uploader(
            "Upload Active Accounts",
            type=['xlsx', 'csv'],
            help="Upload the active accounts file"
        )

    # Second row for additional files
    col4, col5, col6 = st.columns(3)

    with col4:
        issacert_list = st.file_uploader(
            "Upload ISSA CERT List",
            type=['xlsx', 'csv'],
            help="Upload the ISSA CERT List file (UIDs will be marked as ISSA CERT)"
        )

    with col5:
        locked_accounts = st.file_uploader(
            "Upload Locked Accounts",
            type=['xlsx', 'csv'],
            help="Upload locked accounts file (searches Old IC column)"
        )

    with col6:
        st.markdown(
            '<div class="ui-tip"><b>💡 Tips</b><ul>'
            '<li>Optional lists accept .xlsx or .csv</li>'
            '<li>UID columns are detected automatically</li>'
            '<li>Locked Accounts are matched on the <i>Old I.C.</i> column</li>'
            '<li>Skipping a list skips its check</li>'
            '</ul></div>',
            unsafe_allow_html=True,
        )

    # Upload checklist (React)
    render_react(FILES_JS, {"files": [
        _file_info("Cadence Template", cadence_template, True),
        _file_info("Cadence Output", cadence_output, True),
        _file_info("Active Accounts", active_accounts, False),
        _file_info("ISSA CERT List", issacert_list, False),
        _file_info("Locked Accounts", locked_accounts, False),
    ]}, height=68)

    # ============================================
    # HELPER FUNCTIONS
    # ============================================
    def get_latest_entry_by_uid(df, date_column):
        """Get the most recent entry (descending order) for each UID."""
        df_with_date = df[df[date_column].notna()].copy()

        if df_with_date.empty:
            return {}

        df_with_date['_sort_date'] = pd.to_datetime(df_with_date[date_column], errors='coerce')
        df_with_date = df_with_date.sort_values('_sort_date', ascending=False)
        df_with_date = df_with_date.drop_duplicates(subset='UID', keep='first')

        lookup = df_with_date.set_index('UID')[date_column].to_dict()

        return lookup

    def find_uid_column(df):
        """Auto-detect UID column from dataframe"""
        for col in df.columns:
            if 'uid' in col.lower() or col.lower() in ['account', 'acct', 'uid code', 'old i.c.']:
                return col
        return None

    def copy_cell_format(source_cell, target_cell):
        """Copy all formatting from source cell to target cell"""
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

    if 'play_sound' not in st.session_state:
        st.session_state.play_sound = False

    # ============================================
    # PROCESSING FUNCTION
    # ============================================
    def update_status_new_logic(conso_df, ptp_df, kept_df=None, active_accounts_df=None, issacert_df=None, locked_accounts_df=None):
        """
        Sequential Logic:
        1. Active Accounts Check → PULLED OUT if not found
        2. Locked Accounts Check → LOCKED if UID found in "Old IC" column
        3. ISSA CERT Check → ISSA CERT if found in ISSACERT list
        4. PTP Lookup → Only if status is "ACTIVE" → ACTIVE PTP - mm/dd/yyyy
        5. KEPT Lookup → Only if status is "ACTIVE" or "ACTIVE PTP" → PAYMENT
        """

        # ===== STEP 1: ACTIVE ACCOUNTS LOOKUP =====
        active_uids = set()
        if active_accounts_df is not None:
            uid_col = find_uid_column(active_accounts_df)
            if uid_col:
                active_uids = set(active_accounts_df[uid_col].dropna().astype(str).tolist())

        # ===== STEP 2: LOCKED ACCOUNTS LOOKUP (search in "Old IC" column) =====
        locked_old_ic = set()
        if locked_accounts_df is not None:
            # Look for "Old IC" column specifically
            for col in locked_accounts_df.columns:
                if 'old i.c.' in col.lower():
                    locked_old_ic = set(locked_accounts_df[col].dropna().astype(str).tolist())
                    break

        # ===== STEP 2.5: ISSA CERT LOOKUP =====
        issacert_uids = set()
        if issacert_df is not None:
            uid_col = find_uid_column(issacert_df)
            if uid_col:
                issacert_uids = set(issacert_df[uid_col].dropna().astype(str).tolist())

        # ===== STEP 2: PTP LOOKUP (Descending Order) =====
        ptp_lookup = get_latest_entry_by_uid(ptp_df, 'PTP Date')

        # ===== STEP 3: KEPT LOOKUP (Descending Order) =====
        kept_lookup = {}
        if kept_df is not None:
            kept_lookup = get_latest_entry_by_uid(kept_df, 'Claim Paid Date')

        # ===== APPLY UPDATES SEQUENTIALLY =====
        new_statuses = []

        for idx, row in conso_df.iterrows():
            uid = str(row['UID'])
            current_status = str(row.get('Account Status', row.get('Status', ''))).strip().upper()

            # ==== STEP 1: ACTIVE ACCOUNTS CHECK (First) ====
            if active_uids and uid not in active_uids:
                final_status = "PULLED OUT"
            # ==== STEP 2: LOCKED ACCOUNTS CHECK (Second - searches Old IC column) ====
            elif locked_old_ic and uid in locked_old_ic:
                final_status = "LOCKED"
            # ==== STEP 3: ISSA CERT CHECK (Third - only if NOT PULLED OUT and NOT LOCKED) ====
            elif issacert_uids and uid in issacert_uids:
                final_status = "ISSA CERT"
            else:
                # UID IS in Active Accounts and not in ISSACERT - continue with PTP/KEPT checks

                # ==== STEP 3: PTP LOOKUP (Third) ====
                updated_status = row.get('Account Status', row.get('Status', 'Unknown'))

                if current_status == 'ACTIVE' and uid in ptp_lookup and pd.notna(ptp_lookup[uid]):
                    ptp_date = ptp_lookup[uid]

                    # Parse the PTP date
                    try:
                        ptp_datetime = pd.to_datetime(ptp_date)
                    except:
                        ptp_datetime = None

                    # Only update if PTP Date is today or onward
                    if ptp_datetime is not None:
                        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                        ptp_date_only = ptp_datetime.replace(hour=0, minute=0, second=0, microsecond=0)

                        if ptp_date_only >= today:
                            # Format the date as mm/dd/yyyy
                            formatted_date = ptp_datetime.strftime("%m/%d/%Y")
                            updated_status = f"ACTIVE PTP - {formatted_date}"
                        # If PTP Date is in the past, keep original status (ACTIVE)

                # ==== STEP 4: KEPT LOOKUP (Fourth) ====
                status_check = str(updated_status).strip().upper()

                if (status_check == 'ACTIVE' or 'ACTIVE PTP' in status_check) and uid in kept_lookup and pd.notna(kept_lookup[uid]):
                    claim_paid_date = kept_lookup[uid]

                    # Parse the Claim Paid Date and verify it's within current month
                    try:
                        # Try parsing with dd/mm/yyyy format first (common format)
                        claim_paid_datetime = pd.to_datetime(claim_paid_date, dayfirst=True)
                    except:
                        try:
                            claim_paid_datetime = pd.to_datetime(claim_paid_date)
                        except:
                            claim_paid_datetime = None

                    # Only update to PAYMENT if Claim Paid Date is within current month
                    if claim_paid_datetime is not None:
                        today = datetime.now()
                        if claim_paid_datetime.year == today.year and claim_paid_datetime.month == today.month:
                            updated_status = "PAYMENT"
                        # If not in current month, keep the previous status (ACTIVE or ACTIVE PTP)

                final_status = updated_status

            new_statuses.append(final_status)

        conso_df['Account Status'] = new_statuses

        # Get summary info
        uids_with_ptp = set(ptp_lookup.keys())
        uids_with_kept = set(kept_lookup.keys())
        uids_pulled_out = set()
        uids_with_issacert = set()
        uids_with_locked = set()

        if active_uids:
            all_conso_uids = set(conso_df['UID'].astype(str).tolist())
            uids_pulled_out = all_conso_uids - active_uids

        if issacert_uids:
            all_conso_uids = set(conso_df['UID'].astype(str).tolist())
            uids_with_issacert = all_conso_uids & issacert_uids

        if locked_old_ic:
            all_conso_uids = set(conso_df['UID'].astype(str).tolist())
            uids_with_locked = all_conso_uids & locked_old_ic

        return conso_df, uids_with_ptp, uids_with_kept, uids_pulled_out, uids_with_issacert, uids_with_locked

    # ============================================
    # FILE NAME PREFIX SELECTION
    # ============================================
    section_header(2, "File Name Options", "Choose the prefix used for the downloaded file name.")

    file_prefix = st.radio(
        "Select File Name Prefix:",
        options=["BPI LUZ PL", "BPI VISAYAS PL", "BPI MINDANAO PL", "BPI VISAYAS CC", "BPI MINDANAO CC"],
        horizontal=True,
        help="Choose the prefix for the output file name"
    )

    # Show current date for PTP verification
    today = datetime.now()
    current_month = today.strftime("%B %Y")

    # ============================================
    # HERO (React) — rendered into the top slot
    # ============================================
    required_ready = bool(cadence_template and cadence_output)
    required_count = int(bool(cadence_template)) + int(bool(cadence_output))
    optional_count = int(bool(active_accounts)) + int(bool(issacert_list)) + int(bool(locked_accounts))
    with hero_slot:
        render_react(HERO_JS, {
            "title": "Cadence Automation",
            "subtitle": "Fields marked * are required",
            "dateLabel": today.strftime("%B %d, %Y"),
            "monthLabel": current_month,
            "steps": [
                {"label": "Upload files", "detail": f"{required_count}/2 required · {optional_count}/3 optional",
                 "state": "done" if required_ready else "current"},
                {"label": "File name", "detail": file_prefix, "state": "done" if required_ready else "todo"},
                {"label": "Process", "detail": "Ready to run" if required_ready else "Waiting for files",
                 "state": "current" if required_ready else "todo"},
                {"label": "Download", "detail": "Available after processing", "state": "todo"},
            ],
        }, height=272)

    # ============================================
    # PROCESS BUTTON
    # ============================================
    section_header(3, "Process Files",
                   f"PTP Date must be today or later · Claim Paid Date must be within {current_month}")

    btn_col, hint_col = st.columns([1, 4])
    with btn_col:
        process_clicked = st.button("🚀 Process Files", type="primary")
    with hint_col:
        if required_ready:
            st.markdown(f'<div class="ui-ready"><b>● Ready.</b> Output prefix: <b>{html_lib.escape(file_prefix)}</b></div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div class="ui-ready"><i>● Missing required files:</i> upload the Cadence Template and Cadence Output.</div>',
                        unsafe_allow_html=True)

    if process_clicked:
        # Set flag to play sound after processing completes
        st.session_state.play_sound = True

        if cadence_template and cadence_output:
            status_box = st.status("Processing files…", expanded=True)
            try:
                with status_box:
                    # Load Conso Accts sheet from Cadence Template
                    conso_df = pd.read_excel(cadence_template, sheet_name="Conso Accts")

                    # Load PTP sheet from Cadence Output
                    ptp_df = pd.read_excel(cadence_output, sheet_name="PTP")

                    # Try to load KEPT sheet (optional - may not exist)
                    kept_df = None
                    try:
                        kept_df = pd.read_excel(cadence_output, sheet_name="KEPT")
                        st.success("✅ KEPT sheet found in Cadence Output!")
                    except:
                        st.warning("⚠️ KEPT sheet not found in Cadence Output. Only PTP lookup will be applied.")

                    # Load MAPPED sheet from Cadence Output (for DRR append)
                    mapped_df = None
                    drr_combined_df = None
                    try:
                        mapped_df = pd.read_excel(cadence_output, sheet_name="MAPPED")
                        st.success(f"✅ MAPPED sheet found in Cadence Output! {len(mapped_df)} rows")

                        # Load DRR sheet from Cadence Template
                        try:
                            drr_df = pd.read_excel(cadence_template, sheet_name="DRR")
                            st.success(f"✅ DRR sheet found in Cadence Template! {len(drr_df)} existing rows")

                            # Append MAPPED data to DRR
                            if mapped_df is not None and not mapped_df.empty:
                                drr_combined_df = pd.concat([drr_df, mapped_df], ignore_index=True)
                                st.success(f"✅ DRR updated! {len(drr_df)} original + {len(mapped_df)} new = {len(drr_combined_df)} total rows")
                            else:
                                drr_combined_df = drr_df
                        except Exception as e:
                            st.warning(f"⚠️ DRR sheet not found in Cadence Template: {str(e)}")
                            if mapped_df is not None and not mapped_df.empty:
                                drr_combined_df = mapped_df
                                st.info(f"ℹ️ Using MAPPED data only for DRR: {len(drr_combined_df)} rows")
                    except Exception as e:
                        st.warning(f"⚠️ MAPPED sheet not found in Cadence Output: {str(e)}")

                    # Load Active Accounts (optional)
                    active_accounts_df = None
                    if active_accounts:
                        try:
                            if active_accounts.name.endswith('.csv'):
                                active_accounts_df = pd.read_csv(active_accounts)
                            else:
                                active_accounts_df = pd.read_excel(active_accounts)
                            st.success(f"✅ Active Accounts loaded! {len(active_accounts_df)} rows")
                        except Exception as e:
                            st.warning(f"⚠️ Could not load Active Accounts: {str(e)}")

                    # Load ISSA CERT List (optional)
                    issacert_df = None
                    if issacert_list:
                        try:
                            if issacert_list.name.endswith('.csv'):
                                issacert_df = pd.read_csv(issacert_list)
                            else:
                                issacert_df = pd.read_excel(issacert_list)
                            st.success(f"✅ ISSA CERT List loaded! {len(issacert_df)} rows")
                        except Exception as e:
                            st.warning(f"⚠️ Could not load ISSA CERT List: {str(e)}")

                    # Load Locked Accounts (optional)
                    locked_accounts_df = None
                    if locked_accounts:
                        try:
                            if locked_accounts.name.endswith('.csv'):
                                locked_accounts_df = pd.read_csv(locked_accounts)
                            else:
                                locked_accounts_df = pd.read_excel(locked_accounts)
                            st.success(f"✅ Locked Accounts loaded! {len(locked_accounts_df)} rows")
                        except Exception as e:
                            st.warning(f"⚠️ Could not load Locked Accounts: {str(e)}")

                    # Snapshot for the data preview (display only — taken before
                    # processing because the status update edits conso_df in place)
                    preview_conso = conso_df.head(3).copy()
                    preview_ptp = ptp_df.head(3).copy()
                    preview_kept = kept_df.head(3).copy() if kept_df is not None else None
                    preview_counts = (len(conso_df), len(ptp_df), len(kept_df) if kept_df is not None else None)

                    # Process with new logic
                    updated_df, uids_with_ptp, uids_with_kept, uids_pulled_out, uids_with_issacert, uids_with_locked = update_status_new_logic(
                        conso_df, ptp_df, kept_df, active_accounts_df, issacert_df, locked_accounts_df
                    )

                    # ============================================
                    # BUILD OUTPUT WORKBOOK - WITH FORMATTING PRESERVED
                    # ============================================
                    # Read the original template file to preserve formatting
                    original_template = pd.read_excel(cadence_template, sheet_name="Conso Accts")

                    # Find the Account Status column index
                    status_col_name = None
                    for col in original_template.columns:
                        if 'account' in col.lower() and 'status' in col.lower():
                            status_col_name = col
                            break

                    # Create output buffer
                    output_buffer = io.BytesIO()

                    # Use openpyxl to preserve formatting
                    wb = load_workbook(cadence_template)
                    ws = wb['Conso Accts']

                    # Find the header row and Account Status column
                    header_row = 1
                    status_col_idx = None

                    for col_idx, cell in enumerate(ws[header_row], start=1):
                        if cell.value and 'account' in str(cell.value).lower() and 'status' in str(cell.value).lower():
                            status_col_idx = col_idx
                            break

                    # Update the Account Status column with new values
                    if status_col_idx:
                        for row_idx in range(2, ws.max_row + 1):  # Skip header row
                            df_idx = row_idx - 2  # DataFrame index (0-based)
                            if df_idx < len(updated_df):
                                cell = ws.cell(row=row_idx, column=status_col_idx)
                                cell.value = updated_df.iloc[df_idx]['Account Status']

                    # ============================================
                    # UPDATE DRR SHEET - APPEND MAPPED DATA
                    # ============================================
                    if drr_combined_df is not None and not drr_combined_df.empty:
                        # Check if DRR sheet exists in workbook
                        if 'DRR' in wb.sheetnames:
                            ws_drr = wb['DRR']
                            # Clear existing data (keep header row) and write new combined data
                            # First, find the number of columns
                            drr_cols = len(drr_combined_df.columns)

                            # Write header (row 1)
                            for col_idx, col_name in enumerate(drr_combined_df.columns, start=1):
                                ws_drr.cell(row=1, column=col_idx, value=col_name)

                            # Write data starting from row 2
                            for row_idx, row_data in enumerate(drr_combined_df.itertuples(index=False), start=2):
                                for col_idx, value in enumerate(row_data, start=1):
                                    ws_drr.cell(row=row_idx, column=col_idx, value=value)

                            st.success(f"✅ DRR sheet updated with {len(drr_combined_df)} rows")
                        else:
                            # Create new DRR sheet if it doesn't exist
                            ws_drr = wb.create_sheet("DRR")

                            # Write header
                            for col_idx, col_name in enumerate(drr_combined_df.columns, start=1):
                                ws_drr.cell(row=1, column=col_idx, value=col_name)

                            # Write data
                            for row_idx, row_data in enumerate(drr_combined_df.itertuples(index=False), start=2):
                                for col_idx, value in enumerate(row_data, start=1):
                                    ws_drr.cell(row=row_idx, column=col_idx, value=value)

                            st.success(f"✅ New DRR sheet created with {len(drr_combined_df)} rows")

                    # Save workbook
                    wb.save(output_buffer)
                    output_buffer.seek(0)

                    today = datetime.today().strftime("%B %d %Y")

                status_box.update(label="Processing complete — see results below", state="complete", expanded=False)

                # ============================================
                # RESULTS (React dashboard)
                # ============================================
                section_header("✓", "Processing Results", "Summary of status updates applied to the Conso Accts sheet.")

                active_ptp_count = len(updated_df[updated_df['Account Status'].str.contains('ACTIVE PTP', na=False)])
                payment_count = len(updated_df[updated_df['Account Status'] == 'PAYMENT'])

                issacertAccounts = updated_df[updated_df['Account Status'] == 'ISSA CERT']
                lockedAccounts = updated_df[updated_df['Account Status'] == 'LOCKED']
                paymentAccounts = updated_df[updated_df['Account Status'] == 'PAYMENT']
                ptpAccounts = updated_df[updated_df['Account Status'].str.contains('ACTIVE PTP', na=False)]
                pulledOutAccounts = updated_df[updated_df['Account Status'] == 'PULLED OUT']

                total_accounts = len(updated_df)
                final_counts = {
                    "PULLED OUT": len(pulledOutAccounts),
                    "LOCKED": len(lockedAccounts),
                    "ISSA CERT": len(issacertAccounts),
                    "ACTIVE PTP": len(ptpAccounts),
                    "PAYMENT": len(paymentAccounts),
                }
                other_count = max(0, total_accounts - sum(final_counts.values()))

                render_react(STATS_JS, {
                    "total": total_accounts,
                    "cards": [
                        {"label": "Total Accounts", "value": total_accounts, "icon": "users", "t": "indigo", "isTotal": True},
                        {"label": "ISSA CERT", "value": len(uids_with_issacert), "icon": "shield", "t": "violet"},
                        {"label": "LOCKED", "value": len(uids_with_locked), "icon": "lock", "t": "amber"},
                        {"label": "PULLED OUT", "value": len(uids_pulled_out), "icon": "logout", "t": "rose"},
                        {"label": "ACTIVE PTP", "value": active_ptp_count, "icon": "promise", "t": "sky"},
                        {"label": "PAYMENT", "value": payment_count, "icon": "cash", "t": "emerald"},
                    ],
                    "dist": [
                        {"label": "PULLED OUT", "value": final_counts["PULLED OUT"], "t": "rose"},
                        {"label": "LOCKED", "value": final_counts["LOCKED"], "t": "amber"},
                        {"label": "ISSA CERT", "value": final_counts["ISSA CERT"], "t": "violet"},
                        {"label": "ACTIVE PTP", "value": final_counts["ACTIVE PTP"], "t": "sky"},
                        {"label": "PAYMENT", "value": final_counts["PAYMENT"], "t": "emerald"},
                        {"label": "Other / unchanged", "value": other_count, "t": "slate"},
                    ],
                }, height=270)

                sample_groups = [
                    {"key": "issa", "label": "ISSA CERT", "t": "violet", "count": len(issacertAccounts), "rows": _sample_rows(issacertAccounts)},
                    {"key": "locked", "label": "LOCKED", "t": "amber", "count": len(lockedAccounts), "rows": _sample_rows(lockedAccounts)},
                    {"key": "payment", "label": "PAYMENT", "t": "emerald", "count": len(paymentAccounts), "rows": _sample_rows(paymentAccounts)},
                    {"key": "ptp", "label": "ACTIVE PTP", "t": "sky", "count": len(ptpAccounts), "rows": _sample_rows(ptpAccounts)},
                    {"key": "pulled", "label": "PULLED OUT", "t": "rose", "count": len(pulledOutAccounts), "rows": _sample_rows(pulledOutAccounts)},
                ]
                max_rows = max([len(g["rows"]) for g in sample_groups if g["count"] > 0] or [0])
                render_react(SAMPLES_JS, {"groups": sample_groups},
                             height=(150 + max_rows * 38) if max_rows else 190)

                with st.expander("📋 Data preview (first 3 rows of each input sheet)"):
                    col_a, col_b, col_c = st.columns(3)
                    with col_a:
                        st.caption(f"**Conso Accts Sheet** · {preview_counts[0]} rows")
                        st.dataframe(preview_conso, use_container_width=True)
                    with col_b:
                        st.caption(f"**PTP Sheet** · {preview_counts[1]} rows")
                        st.dataframe(preview_ptp, use_container_width=True)
                    with col_c:
                        if preview_kept is not None:
                            st.caption(f"**KEPT Sheet** · {preview_counts[2]} rows")
                            st.dataframe(preview_kept, use_container_width=True)
                        else:
                            st.caption("**KEPT Sheet** · Not loaded")

                # ============================================
                # DOWNLOAD SECTION
                # ============================================
                section_header(4, "Download Result", "Original template formatting is preserved.")

                output_file_name = f"{file_prefix} Collection Effort Cadence as of {today}.xlsx"

                # Play sound notification when process is complete
                # if st.session_state.play_sound:
                #     with open("/Volumes/ACTIVE/cnocon/BPI WFH FILES/AUTOMATION/CALVIN_AUTOMATIONS/answer-correct.mp3", "rb") as f:
                #         st.audio(f.read(), format="audio/mp3", autoplay=True)
                #     # Reset session state so it doesn't replay on every rerun
                #     st.session_state.play_sound = False

                st.balloons()
                # st.session_state.play_sound = True

                with st.container(border=True):
                    dl_info, dl_btn = st.columns([2, 1])
                    with dl_info:
                        st.markdown(
                            '<div class="ui-dl"><div class="ui-dl-icon">📗</div><div>'
                            '<div class="ui-dl-title">Updated Cadence Template is ready</div>'
                            f'<div class="ui-dl-name">{html_lib.escape(output_file_name)}</div>'
                            '</div></div>',
                            unsafe_allow_html=True,
                        )
                    with dl_btn:
                        st.download_button(
                            label="📥 Download Updated Cadence Template",
                            data=output_buffer,
                            file_name=output_file_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )

            except Exception as e:
                status_box.update(label="Processing failed", state="error", expanded=True)
                st.error(f"❌ Error processing files: {str(e)}")
                st.info("Please ensure:")
                st.info("- Cadence Template has 'Conso Accts' sheet")
                st.info("- Cadence Output has 'PTP' sheet")
                st.info("- Both files have 'UID' column")
        else:
            st.warning("⚠️ Please upload both Cadence Template and Cadence Output files!")

    # ============================================
    # SIDEBAR - Instructions (React)
    # ============================================
    with st.sidebar:
        st.markdown('<div class="ui-side-title">📖 Instructions</div>'
                    '<div class="ui-side-sub">How each account status is decided, in order.</div>',
                    unsafe_allow_html=True)
        render_react(SIDEBAR_JS, {
            "order": SIDEBAR_ORDER,
            "rules": SIDEBAR_RULES,
            "features": SIDEBAR_FEATURES,
        }, height=692)

if __name__ == "__main__":
    main()
