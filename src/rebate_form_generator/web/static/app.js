"use strict";

/* ── helpers ──────────────────────────────────────────────────────── */
const $  = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

const PATH_KEYS = ["nb_kb", "dt_kb", "peripheral"];
const JOB_LABEL = {
  consolidate: "Consolidating…",
  stage5:      "Writing rebate raw…",
  "form-data": "Generating form data…",
  report:      "Generating report…",
};

const state = {
  config: null,
  running: null,
  lastOk: null,
  runAll: false,
  browseKey: null,
  browsePath: "",
  suppliers: [],
};

async function api(path, body) {
  const opts = body === undefined
    ? {}
    : { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body) };
  const res = await fetch(path, opts);
  let data = {};
  try { data = await res.json(); } catch (_) { /* empty body */ }
  if (!res.ok || data.ok === false) {
    const err = new Error(data.error || `Request failed (${res.status})`);
    err.data = data;            // callers may need the structured payload
    throw err;
  }
  return data;
}

function toast(msg, kind = "error", ms = 6000) {
  const el = document.createElement("div");
  el.className = `toast ${kind}`;
  el.textContent = msg;
  $("#toasts").appendChild(el);
  setTimeout(() => el.remove(), ms);
}

/* ── modals ───────────────────────────────────────────────────────── */
const modalStack = [];

function showModal(id) {
  $("#overlay").hidden = false;
  const el = $("#" + id);
  el.hidden = false;
  modalStack.push(id);
  const focusable = el.querySelector("select, input, button");
  if (focusable) setTimeout(() => focusable.focus(), 30);
}

function closeModal(id) {
  const target = id || modalStack[modalStack.length - 1];
  if (!target) return;
  $("#" + target).hidden = true;
  const i = modalStack.lastIndexOf(target);
  if (i >= 0) modalStack.splice(i, 1);
  if (!modalStack.length) {
    $("#overlay").hidden = true;
    state.runAll = false;   // abandoning a dialog cancels the Run All chain
  }
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && modalStack.length) closeModal();
});

/* ── status / buttons ─────────────────────────────────────────────── */
const ACTION_BUTTONS = ["#btn-consolidate", "#btn-formdata", "#btn-report", "#btn-runall"];

function setRunning(name) {
  state.running = name;
  const pill = $("#status");
  const txt  = $("#status-text");
  if (name) {
    pill.dataset.state = "running";
    txt.textContent = JOB_LABEL[name] || "Running…";
  } else if (state.lastOk === false) {
    pill.dataset.state = "error";
    txt.textContent = "Error";
  } else if (state.lastOk === true) {
    pill.dataset.state = "done";
    txt.textContent = "Done";
  } else {
    pill.dataset.state = "idle";
    txt.textContent = "Idle";
  }
  ACTION_BUTTONS.forEach((s) => { $(s).disabled = !!name; });
}

/* ── log ──────────────────────────────────────────────────────────── */
const logEl = () => $("#log");

function clearLog() {
  logEl().innerHTML = '<div class="log-empty">No output yet.</div>';
}

function appendLog(level, msg) {
  const box = logEl();
  const empty = box.querySelector(".log-empty");
  if (empty) empty.remove();

  const atBottom = box.scrollHeight - box.scrollTop - box.clientHeight < 60;
  const isHead = msg.trimStart().startsWith("===");

  const line = document.createElement("div");
  line.className = `line ${level}${isHead ? " head" : ""}`;
  const tag = document.createElement("span");
  tag.className = "tag";
  tag.textContent = `[${level}] `;
  line.appendChild(tag);
  line.appendChild(document.createTextNode(msg));
  box.appendChild(line);

  if (atBottom) box.scrollTop = box.scrollHeight;
}

/* ── SSE ──────────────────────────────────────────────────────────── */
function connect() {
  const es = new EventSource("/api/events");
  es.onmessage = (e) => {
    let ev;
    try { ev = JSON.parse(e.data); } catch (_) { return; }
    switch (ev.type) {
      case "log":   appendLog(ev.level || "INFO", ev.msg || ""); break;
      case "clear": clearLog(); break;
      case "state": setRunning(ev.running); break;
      case "done":  onJobDone(ev); break;
    }
  };
  es.onerror = () => { /* EventSource reconnects on its own */ };
}

function onJobDone(ev) {
  state.lastOk = !ev.error;
  const r = ev.result || {};
  if (ev.error) { state.runAll = false; return; }

  if (ev.job === "consolidate") {
    state.runAll = !!r.run_all;
    if (r.fy_sheets && r.fy_sheets.length) {
      openFyModal(r.fy_sheets, r.coverage || {});
    } else {
      state.runAll = false;
    }
    showIssues(r.issues);           // stacks on top of the FY dialog
  } else if (ev.job === "stage5") {
    // Stages 1-5 run here now, so this is where source-file problems surface.
    if (r.path) toast("rebate raw.xlsx written.", "ok", 4000);
    if (state.runAll) {
      openQuarterModal();
      showIssues(r.issues);
    } else {
      showIssues(r.issues);
    }
  } else if (ev.job === "form-data") {
    if (state.runAll) {
      openReportModal().then(() => showIssues(r.issues));
    } else {
      showIssues(r.issues);
    }
  } else if (ev.job === "report") {
    state.runAll = false;
    if (r.download) {
      downloadBundle(r.download, r.zipName, (r.files || []).length);
    }
    if (r.blanks && Object.keys(r.blanks).length) openBlanksModal(r.blanks);
    showIssues(r.issues);
  }
}

/* Hand the generated contracts to the browser as one zip, which lands in
   the normal Downloads folder. Nothing is kept under data/output. */
function downloadBundle(url, zipName, count) {
  const a = document.createElement("a");
  a.href = url;
  a.download = zipName || "contracts.zip";
  document.body.appendChild(a);
  a.click();
  a.remove();
  toast(`${count} contract${count === 1 ? "" : "s"} downloaded as ${zipName}`, "ok", 9000);
}

/* ── config ───────────────────────────────────────────────────────── */
async function loadConfig() {
  state.config = await api("/api/config");
  PATH_KEYS.forEach((k) => {
    const input = $("#" + k);
    input.value = state.config[k] || "";
    input.title = input.value;        // the full path on hover
  });
  return state.config;
}

let saveTimer = null;
function saveConfigSoon() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    const body = {};
    PATH_KEYS.forEach((k) => { body[k] = $("#" + k).value.trim(); });
    api("/api/config", body).catch(() => {});
    checkPaths();
  }, 400);
}

/* Ask the server whether each folder really holds the expected
   "Master price table_…" sub-folder, and show the verdict inline. */
const MARK = { ok: "✓", bad: "✕", empty: "!" };

async function checkPaths() {
  const qs = PATH_KEYS
    .map((k) => `${k}=${encodeURIComponent($("#" + k).value.trim())}`)
    .join("&");
  let data;
  try { data = await api("/api/check-paths?" + qs); } catch (_) { return; }
  PATH_KEYS.forEach((k) => {
    const el = $("#status-" + k);
    const r = data[k] || {};
    el.dataset.state = r.state || "";
    $(".mark", el).textContent = MARK[r.state] || "·";
    $(".detail", el).textContent = r.detail || "";
  });
}

function currentPaths() {
  const out = {};
  PATH_KEYS.forEach((k) => { out[k] = $("#" + k).value.trim(); });
  return out;
}

/* ── Stage 1-4 ────────────────────────────────────────────────────── */
async function startConsolidate(runAll) {
  state.runAll = runAll;
  state.lastOk = null;
  try {
    await api("/api/consolidate", { ...currentPaths(), run_all: runAll });
  } catch (err) {
    state.runAll = false;
    toast(err.message);
  }
}

/* ── Stage 5: FY modal ────────────────────────────────────────────── */
function fyKey(s) {
  const digits = (s.match(/\d/g) || []).join("");
  return digits ? parseInt(digits, 10) : 0;
}

let fyCoverage = {};

function openFyModal(sheets, coverage) {
  fyCoverage = coverage;
  const sel = $("#fy-select");
  sel.innerHTML = "";
  sheets.forEach((s) => {
    const o = document.createElement("option");
    o.value = s; o.textContent = s;
    sel.appendChild(o);
  });
  const def = sheets.slice().sort((a, b) => fyKey(b) - fyKey(a))[0];
  if (def) sel.value = def;
  updateFyInfo();
  showModal("modal-fy");
}

function updateFyInfo() {
  const fy = ($("#fy-select").value || "").trim().toUpperCase();
  const box = $("#fy-info");
  const missing = [];
  for (const [seg, suppliers] of Object.entries(fyCoverage)) {
    const bad = Object.entries(suppliers)
      .filter(([, fys]) => !(fys || []).some((f) => String(f).trim().toUpperCase() === fy))
      .map(([sup]) => sup);
    if (bad.length) missing.push(`${seg}: ${bad.join(", ")}`);
  }
  box.classList.remove("ok", "missing");
  if (!Object.keys(fyCoverage).length) {
    box.textContent = "(No coverage data)";
  } else if (missing.length) {
    box.classList.add("missing");
    box.textContent = `Missing ${$("#fy-select").value}:\n` +
      missing.map((m) => "  " + m).join("\n");
  } else {
    box.classList.add("ok");
    box.textContent = `\u2713  All segments have ${$("#fy-select").value}`;
  }
}

async function runStage5() {
  const fy_sheet = $("#fy-select").value;
  const runAll = state.runAll;
  closeModal("modal-fy");
  state.runAll = runAll;
  try {
    await api("/api/stage5", { fy_sheet, run_all: runAll });
  } catch (err) { state.runAll = false; toast(err.message); }
}

/* ── Stage 6: quarter modal ───────────────────────────────────────── */
function openQuarterModal() {
  const cfg = state.config;
  const fy = cfg.current_fy;
  const sel = $("#quarter-select");
  sel.innerHTML = "";
  for (let q = 1; q <= 4; q++) {
    const label = `FY${String(fy).padStart(2, "0")} Q${q}`;
    const o = document.createElement("option");
    o.value = label; o.textContent = label;
    sel.appendChild(o);
  }
  sel.value = `FY${String(fy).padStart(2, "0")} Q${cfg.current_quarter}`;

  const grid = $("#column-grid");
  grid.innerHTML = "";
  cfg.optional_columns.forEach((col) => {
    const label = document.createElement("label");
    label.className = "check";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = col;
    cb.checked = cfg.default_checked.includes(col);
    label.appendChild(cb);
    label.appendChild(document.createTextNode(col));
    grid.appendChild(label);
  });
  showModal("modal-quarter");
}

async function runStage6() {
  const label = $("#quarter-select").value;          // e.g. "FY26 Q3"
  const [fyPart, qPart] = label.split(" ");
  const fy = parseInt(fyPart.slice(2), 10);
  const quarter = parseInt(qPart.slice(1), 10);
  const columns = $$("#column-grid input:checked").map((cb) => cb.value);
  const runAll = state.runAll;
  closeModal("modal-quarter");
  state.runAll = runAll;
  try {
    await api("/api/form-data", { fy, quarter, columns, run_all: runAll });
    state.config = await api("/api/config");
  } catch (err) { state.runAll = false; toast(err.message); }
}

/* ── Stage 7: report modal ────────────────────────────────────────── */
async function openReportModal() {
  // Check the signer settings before asking for Form #s, not after.
  try {
    const pre = await api("/api/report/preflight");
    if (!pre.ok) {
      const err = new Error(pre.error || "Settings needed");
      err.data = pre;
      if (showNeedsSettings(err)) { state.runAll = false; return; }
    }
  } catch (err) {
    if (showNeedsSettings(err)) { state.runAll = false; return; }
  }

  const cfg = await loadConfig();
  const { suppliers } = await api("/api/suppliers");
  state.suppliers = suppliers;

  const tbody = $("#sup-rows");
  tbody.innerHTML = "";
  $("#report-error").textContent = "";

  if (!suppliers.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 4;
    td.className = "empty";
    td.textContent = "No contract input files found. Run “Generate Form Data” first.";
    tr.appendChild(td);
    tbody.appendChild(tr);
  } else {
    suppliers.forEach((sup) => {
      const tr = document.createElement("tr");

      const tdChk = document.createElement("td");
      const chk = document.createElement("input");
      chk.type = "checkbox";
      chk.checked = true;
      chk.dataset.supplier = sup;
      chk.className = "sup-check";
      tdChk.appendChild(chk);

      const tdName = document.createElement("td");
      tdName.textContent = sup;

      const tdIc = document.createElement("td");
      const ic = document.createElement("input");
      ic.type = "text";
      ic.placeholder = "iCertis Code";
      ic.value = (cfg.icertis_codes || {})[sup] || "";
      ic.dataset.ic = sup;
      tdIc.appendChild(ic);

      const tdFn = document.createElement("td");
      const fn = document.createElement("input");
      fn.type = "text";
      fn.placeholder = "Form #";
      fn.value = (cfg.form_numbers || {})[sup] || "";
      fn.dataset.fn = sup;
      tdFn.appendChild(fn);

      tr.append(tdChk, tdName, tdIc, tdFn);
      tbody.appendChild(tr);
    });
  }
  $("#sup-all").checked = true;
  showModal("modal-report");
}

async function runStage7() {
  const selected = $$(".sup-check").filter((c) => c.checked).map((c) => c.dataset.supplier);
  const errEl = $("#report-error");
  if (!selected.length) {
    errEl.textContent = "Select at least one supplier.";
    return;
  }
  const form_numbers = {};
  const icertis_codes = {};
  const missing = [];
  selected.forEach((sup) => {
    const fn = $(`[data-fn="${CSS.escape(sup)}"]`).value.trim();
    const ic = $(`[data-ic="${CSS.escape(sup)}"]`).value.trim();
    if (fn) form_numbers[sup] = fn; else missing.push(sup);
    if (ic) icertis_codes[sup] = ic;
  });
  if (missing.length) {
    errEl.textContent = `Form # required for: ${missing.join(", ")}`;
    return;
  }
  try {
    await api("/api/report", { suppliers: selected, form_numbers, icertis_codes });
    closeModal("modal-report");
  } catch (err) {
    // Missing signer settings stop the run; keep the dialog so nothing typed
    // here is lost while the user goes and fills them in.
    if (!showNeedsSettings(err)) {
      closeModal("modal-report");
      toast(err.message);
    }
  }
}

/* Halt-and-explain dialog for settings that must exist before any contract
   can be produced. Returns true when it handled the error. */
function showNeedsSettings(err) {
  const payload = (err && err.data) || {};
  const missing = payload.needsSettings || [];
  if (!missing.length) {
    if (payload.error && /template/i.test(payload.error)) {
      $("#needs-intro").textContent = payload.error;
      $("#needs-list").innerHTML = "";
      $("#needs-why").hidden = true;
      $("#needs-open").textContent = "Open Contract Template";
      $("#needs-open").onclick = () => { closeModal("modal-needs-settings"); openTemplates(); };
      showModal("modal-needs-settings");
      return true;
    }
    return false;
  }

  $("#needs-intro").textContent = missing.length === 1
    ? "This value has to be set before contracts can be generated:"
    : "These values have to be set before contracts can be generated:";
  const list = $("#needs-list");
  list.innerHTML = "";
  missing.forEach((m) => {
    const row = document.createElement("div");
    row.className = "needs-item";
    const label = document.createElement("b");
    label.textContent = m.label;
    const code = document.createElement("code");
    code.textContent = m.key;
    row.append(label, code);
    list.appendChild(row);
  });
  $("#needs-why").hidden = false;
  $("#needs-open").textContent = "Open Contract Info";
  $("#needs-open").onclick = () => {
    closeModal("modal-needs-settings");
    openContractInfo();
  };
  showModal("modal-needs-settings");
  return true;
}

/* ── File Tracker ─────────────────────────────────────────────────── */
function fmtWhen(iso) {
  if (!iso) return { date: "—", rel: "never" };
  const d = new Date(iso);
  if (isNaN(d)) return { date: iso, rel: "" };
  const days = Math.round((Date.now() - d.getTime()) / 86400000);
  const rel = days <= 0 ? "today" : days === 1 ? "yesterday" : `${days} days ago`;
  return {
    date: d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" }),
    rel,
  };
}

async function openTracker() {
  showModal("modal-tracker");
  await refreshTracker();
}

async function refreshTracker() {
  const list = $("#tracker-list");
  const btn = $("#tracker-refresh");
  btn.disabled = true;
  list.innerHTML = '<div class="tracker-error" style="border-color:var(--border-soft);' +
                   'background:transparent;color:var(--text-faint)">Checking…</div>';

  // Use whatever is typed in the form, so the tracker follows an edited path
  // before it has been saved.
  const qs = PATH_KEYS
    .map((k) => `${k}=${encodeURIComponent(($("#" + k) || {}).value || "")}`)
    .join("&");

  let data;
  try {
    data = await api("/api/file-tracker?" + qs);
  } catch (err) {
    list.innerHTML = "";
    const e = document.createElement("div");
    e.className = "tracker-error";
    e.textContent = err.message;
    list.appendChild(e);
    btn.disabled = false;
    return;
  }

  $("#tracker-sub").textContent =
    `Newest workbook in each supplier folder, and whether it has been updated in ${data.month}.`;
  $("#tracker-checked").textContent = `Checked at ${data.checkedAt}`;

  const summary = $("#tracker-summary");
  summary.dataset.state = data.stale ? "stale" : "ok";
  summary.innerHTML = "";
  const icon = document.createElement("span");
  icon.textContent = data.stale ? "⚠" : "✓";
  const text = document.createElement("span");
  if (!data.total) {
    text.textContent = "No supplier folders found.";
  } else if (data.stale) {
    text.innerHTML = `<b>${data.stale}</b> of <b>${data.total}</b> supplier file(s) ` +
                     `have not been updated in ${data.month}.`;
  } else {
    text.innerHTML = `All <b>${data.total}</b> supplier files were updated in ${data.month}.`;
  }
  summary.append(icon, text);

  list.innerHTML = "";
  (data.segments || []).forEach((seg) => {
    const group = document.createElement("div");
    group.className = "tracker-group";

    const h = document.createElement("h4");
    h.textContent = seg.label;
    if (!seg.error) {
      const n = seg.files.length;
      const bad = seg.files.filter((f) => !f.thisMonth).length;
      const note = document.createElement("span");
      note.textContent = bad ? `${n} supplier(s) · ${bad} out of date` : `${n} supplier(s)`;
      h.appendChild(note);
    }
    group.appendChild(h);

    if (seg.error) {
      const e = document.createElement("div");
      e.className = "tracker-error";
      e.textContent = seg.error;
      group.appendChild(e);
      list.appendChild(group);
      return;
    }

    if (!seg.files.length) {
      const e = document.createElement("div");
      e.className = "tracker-error";
      e.style.cssText = "border-color:var(--border-soft);background:transparent;" +
                        "color:var(--text-faint)";
      e.textContent = "No supplier folders matched the naming convention.";
      group.appendChild(e);
      list.appendChild(group);
      return;
    }

    // Out-of-date first: that is what this screen is for.
    const files = seg.files.slice().sort(
      (a, b) => (a.thisMonth === b.thisMonth ? 0 : a.thisMonth ? 1 : -1));

    files.forEach((f) => {
      const row = document.createElement("div");
      row.className = "tracker-row";
      row.dataset.ok = String(!!f.thisMonth);

      const mark = document.createElement("span");
      mark.className = "tracker-mark";
      mark.textContent = f.thisMonth ? "✓" : "⚠";
      mark.title = f.thisMonth ? `Updated in ${data.month}` : `Not updated in ${data.month}`;

      const sup = document.createElement("span");
      sup.className = "tracker-supplier";
      sup.textContent = f.supplier;

      const name = document.createElement("span");
      name.className = "tracker-file" + (f.file ? "" : " missing");
      name.textContent = f.file || f.note || "No .xlsx found";
      if (f.path) name.title = f.path;

      const when = document.createElement("span");
      when.className = "tracker-when";
      const w = fmtWhen(f.modified);
      const strong = document.createElement("b");
      strong.textContent = w.date;
      when.append(strong, document.createTextNode(w.rel));

      row.append(mark, sup, name, when);
      group.appendChild(row);
    });

    list.appendChild(group);
  });

  btn.disabled = false;
}

/* ── Exit and idle shutdown ───────────────────────────────────────── */
const IDLE_LIMIT_MS = 30 * 60 * 1000;      // close after this much inactivity
const IDLE_WARN_MS = 60 * 1000;            // warn this long before closing
const IDLE_TICK_MS = 10 * 1000;

const idle = { last: Date.now(), warned: false, closing: false };

function markActive() {
  idle.last = Date.now();
  if (idle.warned) {
    idle.warned = false;
    const t = $("#idle-toast");
    if (t) t.remove();
  }
}

function idleWarn(secondsLeft) {
  let el = $("#idle-toast");
  if (!el) {
    el = document.createElement("div");
    el.className = "toast warn";
    el.id = "idle-toast";
    $("#toasts").appendChild(el);
  }
  el.textContent =
    `Closing in ${secondsLeft}s — no activity for 30 minutes. Move the mouse to stay open.`;
}

function idleTick() {
  if (idle.closing) return;
  // A running job is work in progress, not idleness.
  if (state.running || modalStack.length) { markActive(); return; }

  const since = Date.now() - idle.last;
  if (since >= IDLE_LIMIT_MS) {
    shutdownApp("Closed after 30 minutes of inactivity.");
  } else if (since >= IDLE_LIMIT_MS - IDLE_WARN_MS) {
    idle.warned = true;
    idleWarn(Math.max(1, Math.ceil((IDLE_LIMIT_MS - since) / 1000)));
  }
}

async function shutdownApp(message) {
  if (idle.closing) return;
  idle.closing = true;
  $("#closed-title").textContent = "Closed";
  $("#closed-text").textContent =
    message || "The local server has stopped. You can close this window.";
  $("#closed-screen").hidden = false;
  try {
    await fetch("/api/shutdown", { method: "POST" });
  } catch (_) { /* the server may drop the connection as it goes */ }
  // The launcher closes the app window itself; this covers the cases where
  // it cannot (a plain browser tab, or Chrome handed the URL to an existing
  // instance), so the page at least stops pretending to be live.
  setTimeout(() => { try { window.close(); } catch (_) {} }, 400);
}

function confirmExit() {
  if (state.running &&
      !window.confirm("A job is still running. Close anyway?")) return;
  shutdownApp("The local server has stopped. You can close this window.");
}

/* ── Settings: Contract Info ──────────────────────────────────────── */
const ci = { fields: [], keyField: "Supplier Name", reserved: [], sharedLabels: {} };

/* The table is rebuilt from scratch whenever the columns change, so the
   current cell values are read back out first and carried over. */
function ciSnapshotRows() {
  return $$("#ci-rows tr").map((tr) => {
    const row = {};
    $$("input[data-field]", tr).forEach((i) => { row[i.dataset.field] = i.value; });
    return row;
  });
}

function ciRenderHead() {
  const head = $("#ci-head");
  head.innerHTML = "";
  ci.fields.forEach((f, idx) => {
    const th = document.createElement("th");
    const wrap = document.createElement("div");
    wrap.className = "col-head";

    const name = document.createElement("input");
    name.type = "text";
    name.className = "col-name";
    name.value = f;
    name.spellcheck = false;
    if (idx === 0) {
      name.disabled = true;
      name.title = "The key column cannot be renamed";
    } else {
      name.title = "This name must match the placeholder in the Word template";
      name.onchange = () => ciRenameColumn(idx, name.value);
    }
    wrap.appendChild(name);

    if (idx > 0) {
      const del = document.createElement("button");
      del.className = "col-del";
      del.textContent = "×";
      del.title = `Remove the ${f} column`;
      del.onclick = () => ciRemoveColumn(idx);
      wrap.appendChild(del);
    }
    th.appendChild(wrap);
    head.appendChild(th);
  });
  head.appendChild(document.createElement("th"));
}

function ciRenderRows(rows) {
  $("#ci-rows").innerHTML = "";
  (rows && rows.length ? rows : [null]).forEach((r) => ciAddRow(r));
  ciUpdateNote();
}

function ciRebuild(rows) {
  ciRenderHead();
  ciRenderRows(rows);
}

function ciRenameColumn(idx, raw) {
  const name = (raw || "").trim();
  const rows = ciSnapshotRows();
  const old = ci.fields[idx];
  if (!name) { ciRebuild(rows); return; }
  if (ci.reserved.includes(name)) {
    $("#ci-error").textContent =
      `${name} is filled in automatically and cannot be a column.`;
    ciRebuild(rows);
    return;
  }
  if (ci.fields.some((f, i) => i !== idx && f.toLowerCase() === name.toLowerCase())) {
    $("#ci-error").textContent = `There is already a ${name} column.`;
    ciRebuild(rows);
    return;
  }
  ci.fields[idx] = name;
  rows.forEach((r) => { r[name] = r[old] || ""; delete r[old]; });
  $("#ci-error").textContent = "";
  ciRebuild(rows);
}

function ciRemoveColumn(idx) {
  const field = ci.fields[idx];
  const rows = ciSnapshotRows();
  const used = rows.filter((r) => (r[field] || "").trim()).length;
  if (used && !window.confirm(
      `Remove the "${field}" column? ${used} supplier(s) have a value in it.`)) {
    return;
  }
  ci.fields.splice(idx, 1);
  rows.forEach((r) => delete r[field]);
  $("#ci-error").textContent = "";
  ciRebuild(rows);
}

function ciAddColumn() {
  const name = (window.prompt(
    "New column name.\n\nUse the exact placeholder keyword from the Word template, " +
    "e.g. ICMSupplierAddress.") || "").trim();
  if (!name) return;
  if (ci.reserved.includes(name)) {
    $("#ci-error").textContent =
      `${name} is filled in automatically and cannot be a column.`;
    return;
  }
  if (ci.fields.some((f) => f.toLowerCase() === name.toLowerCase())) {
    $("#ci-error").textContent = `There is already a ${name} column.`;
    return;
  }
  const rows = ciSnapshotRows();
  ci.fields.push(name);
  rows.forEach((r) => { r[name] = ""; });
  $("#ci-error").textContent = "";
  ciRebuild(rows);
  $("#ci-table").parentElement.scrollLeft = 99999;
}

function ciAddRow(row) {
  const tr = document.createElement("tr");
  ci.fields.forEach((f, idx) => {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.dataset.field = f;
    input.value = (row && row[f]) || "";
    if (idx === 0) input.placeholder = "e.g. Chicony";
    td.appendChild(input);
    tr.appendChild(td);
  });
  const tdDel = document.createElement("td");
  tdDel.className = "row-del";
  const del = document.createElement("button");
  del.className = "del-btn";
  del.title = "Remove this supplier";
  del.textContent = "×";
  del.onclick = () => { tr.remove(); ciUpdateNote(); };
  tdDel.appendChild(del);
  tr.appendChild(tdDel);
  $("#ci-rows").appendChild(tr);
  return tr;
}

function ciUpdateNote() {
  const n = $$("#ci-rows tr").length;
  $("#ci-note").textContent =
    `${n} supplier${n === 1 ? "" : "s"} · ${ci.fields.length - 1} keyword column(s)`;
}

function ciRenderShared(shared) {
  const box = $("#ci-shared");
  box.innerHTML = "";
  Object.keys(shared || {}).forEach((key) => {
    const field = document.createElement("div");
    field.className = "shared-field";
    const label = document.createElement("label");
    label.textContent = ci.sharedLabels[key] || key;
    const code = document.createElement("code");
    code.textContent = key;
    label.appendChild(code);
    const input = document.createElement("input");
    input.type = "text";
    input.dataset.shared = key;
    input.value = shared[key] || "";
    input.placeholder = "leave blank to keep the template's own text";
    field.append(label, input);
    box.appendChild(field);
  });
}

async function openContractInfo() {
  let data;
  try { data = await api("/api/contract-info"); }
  catch (err) { toast("Could not load contract info: " + err.message); return; }

  ci.fields = (data.fields || []).slice();
  ci.keyField = data.keyField || ci.fields[0] || "Supplier Name";
  ci.reserved = data.reserved || [];
  ci.sharedLabels = data.sharedLabels || {};

  $("#ci-error").textContent = "";
  ciRenderShared(data.shared || {});
  ciRebuild(data.rows || []);
  showModal("modal-contract-info");
}

async function saveContractInfo() {
  const key = ci.fields[0];
  const rows = ciSnapshotRows().map((r) => {
    const out = {};
    ci.fields.forEach((f) => { out[f] = (r[f] || "").trim(); });
    return out;
  });
  const named = rows.filter((r) => r[key]);
  const partial = rows.filter((r) => !r[key] && ci.fields.slice(1).some((f) => r[f]));
  if (partial.length) {
    $("#ci-error").textContent =
      `${partial.length} row(s) have details but no ${key} — fill it in or remove the row.`;
    return;
  }
  const seen = new Set();
  const dupes = named.map((r) => r[key])
    .filter((n) => (seen.has(n.toLowerCase()) ? true : (seen.add(n.toLowerCase()), false)));
  if (dupes.length) {
    $("#ci-error").textContent = `Duplicate supplier: ${[...new Set(dupes)].join(", ")}`;
    return;
  }
  const shared = {};
  $$("[data-shared]").forEach((i) => { shared[i.dataset.shared] = i.value.trim(); });

  try {
    const res = await api("/api/contract-info",
                          { fields: ci.fields, rows: named, shared });
    closeModal("modal-contract-info");
    toast(`Contract info saved (${res.rows.length} supplier(s), ` +
          `${res.fields.length - 1} keyword column(s)).`, "ok", 3500);
  } catch (err) { $("#ci-error").textContent = err.message; }
}

/* ── Settings: Contract Template ──────────────────────────────────── */
const tplState = { items: [], active: "", selected: "", wordAvailable: true };

function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  return d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}

function fmtSize(bytes) {
  if (!bytes && bytes !== 0) return "";
  return bytes < 1024 * 1024
    ? `${Math.round(bytes / 1024)} KB`
    : `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

async function openTemplates() {
  await refreshTemplates();
  showModal("modal-template");
}

async function refreshTemplates(selectId) {
  let data;
  try { data = await api("/api/templates"); }
  catch (err) { toast("Could not load templates: " + err.message); return; }

  tplState.items = data.items || [];
  tplState.active = data.active || "";
  tplState.wordAvailable = data.word_available !== false;
  tplState.selected = selectId || tplState.active ||
    (tplState.items[0] && tplState.items[0].id) || "";

  $("#tpl-count").textContent = tplState.items.length
    ? `${tplState.items.length} template${tplState.items.length === 1 ? "" : "s"}`
    : "No templates yet";
  $("#tpl-error").textContent = "";

  const list = $("#tpl-list");
  list.innerHTML = "";
  tplState.items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "tpl-row";
    row.dataset.selected = String(item.id === tplState.selected);
    row.onclick = (e) => {
      if (e.target.closest("input, button")) return;
      selectTemplate(item.id);
    };

    const radio = document.createElement("input");
    radio.type = "radio";
    radio.className = "tpl-radio";
    radio.name = "tpl-active";
    radio.checked = item.id === tplState.active;
    radio.title = "Use this template";
    radio.onchange = () => activateTemplate(item.id);

    const name = document.createElement("input");
    name.type = "text";
    name.className = "tpl-name";
    name.value = item.name || item.file;
    name.title = item.file;
    name.onchange = () => renameTemplate(item.id, name.value, name);

    const meta = document.createElement("span");
    meta.className = "tpl-meta";
    meta.textContent = `added ${fmtDate(item.added)} · ${fmtSize(item.size)}`;

    const right = document.createElement("span");
    right.style.display = "flex";
    right.style.alignItems = "center";
    right.style.gap = "8px";
    if (item.id === tplState.active) {
      const tag = document.createElement("span");
      tag.className = "tpl-inuse";
      tag.textContent = "In use";
      right.appendChild(tag);
    }
    const del = document.createElement("button");
    del.className = "del-btn";
    del.textContent = "×";
    del.title = "Delete this template";
    del.onclick = () => deleteTemplate(item);
    right.appendChild(del);

    row.append(radio, name, meta, right);
    list.appendChild(row);
  });

  if (!tplState.items.length) {
    const empty = document.createElement("div");
    empty.className = "tpl-frame-empty";
    empty.style.position = "static";
    empty.textContent = "No template yet — upload a .docx to get started.";
    list.appendChild(empty);
  }

  renderTemplatePreview();
}

function selectTemplate(id) {
  tplState.selected = id;
  $$("#tpl-list .tpl-row").forEach((r, i) => {
    r.dataset.selected = String(tplState.items[i] && tplState.items[i].id === id);
  });
  renderTemplatePreview();
}

async function activateTemplate(id) {
  try {
    await api("/api/templates/activate", { id });
    await refreshTemplates(id);
    toast("Template switched.", "ok", 2500);
  } catch (err) { $("#tpl-error").textContent = err.message; }
}

async function renameTemplate(id, name, input) {
  const trimmed = name.trim();
  if (!trimmed) { await refreshTemplates(tplState.selected); return; }
  try {
    await api("/api/templates/rename", { id, name: trimmed });
    const item = tplState.items.find((i) => i.id === id);
    if (item) item.name = trimmed;
    if (input) input.blur();
  } catch (err) { $("#tpl-error").textContent = err.message; }
}

async function deleteTemplate(item) {
  const inUse = item.id === tplState.active;
  const msg = inUse
    ? `Delete "${item.name}"? It is the template currently in use.`
    : `Delete "${item.name}"?`;
  if (!window.confirm(msg)) return;
  try {
    await api("/api/templates/delete", { id: item.id });
    await refreshTemplates();
    toast("Template deleted.", "ok", 2500);
  } catch (err) { $("#tpl-error").textContent = err.message; }
}

async function renderTemplatePreview() {
  const frame = $("#tpl-frame");
  const empty = $("#tpl-frame-empty");
  const state = $("#tpl-preview-state");
  const phBox = $("#tpl-placeholders");
  const item = tplState.items.find((i) => i.id === tplState.selected);

  frame.hidden = true;
  frame.removeAttribute("src");
  phBox.innerHTML = "";

  if (!item) {
    $("#tpl-preview-title").textContent = "Preview";
    empty.hidden = false;
    empty.textContent = "Select a template to preview it.";
    state.textContent = "";
    return;
  }

  $("#tpl-preview-title").textContent = item.name || item.file;

  // Placeholder check first: it is instant and is the thing worth knowing.
  try {
    const info = await api("/api/templates/inspect?id=" + encodeURIComponent(item.id));
    (info.found || []).forEach((p) => phBox.appendChild(phChip(p, true)));
    (info.missing || []).forEach((p) => phBox.appendChild(phChip(p, false)));
    if (info.missing && info.missing.length) {
      state.textContent = `${info.missing.length} placeholder(s) missing from this template`;
    } else if (info.ok) {
      state.textContent = `all placeholders present · ${info.tables} table(s)`;
    }
  } catch (_) { /* preview still worth trying */ }

  if (!tplState.wordAvailable) {
    empty.hidden = false;
    empty.textContent =
      "Microsoft Word is not available on this machine, so the document cannot be " +
      "rendered here. The placeholder check above still applies.";
    return;
  }

  empty.hidden = false;
  empty.textContent = "Rendering with Word… the first time takes about 15 seconds.";
  const token = item.id;
  try {
    const res = await fetch("/api/templates/preview?id=" + encodeURIComponent(item.id));
    if (tplState.selected !== token) return;        // user moved on
    if (!res.ok) {
      let detail = "";
      try { detail = (await res.json()).error || ""; } catch (_) {}
      empty.hidden = false;
      empty.textContent = detail || "Could not render this document.";
      return;
    }
    const blob = await res.blob();
    if (tplState.selected !== token) return;
    frame.src = URL.createObjectURL(blob) + "#toolbar=1&view=FitH";
    frame.hidden = false;
    empty.hidden = true;
  } catch (err) {
    empty.hidden = false;
    empty.textContent = "Could not render this document: " + err.message;
  }
}

function phChip(name, found) {
  const el = document.createElement("span");
  el.className = "ph";
  el.dataset.found = String(found);
  el.textContent = (found ? "✓ " : "✗ ") + name;
  return el;
}

async function uploadTemplate(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".docx")) {
    $("#tpl-error").textContent = "Only .docx files are supported.";
    return;
  }
  $("#tpl-error").textContent = "";
  $("#tpl-count").textContent = "Uploading…";
  const qs = `?filename=${encodeURIComponent(file.name)}` +
             `&name=${encodeURIComponent(file.name.replace(/\.docx$/i, ""))}`;
  try {
    const res = await fetch("/api/templates/upload" + qs, {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream" },
      body: file,
    });
    const data = await res.json();
    if (!res.ok || data.ok === false) throw new Error(data.error || "Upload failed");
    await refreshTemplates(data.template.id);
    toast(`"${data.template.name}" uploaded and set as the template in use.`, "ok", 4000);
  } catch (err) {
    $("#tpl-error").textContent = err.message;
    await refreshTemplates(tplState.selected);
  }
}

/* ── source file problems ─────────────────────────────────────────── */
let lastIssues = [];

function showIssues(list) {
  lastIssues = list || [];
  if (!lastIssues.length) return false;

  const errors = lastIssues.filter((i) => i.severity === "error").length;
  const warns = lastIssues.length - errors;
  const parts = [];
  if (errors) parts.push(`${errors} blocking`);
  if (warns) parts.push(`${warns} warning${warns > 1 ? "s" : ""}`);
  $("#issues-sub").textContent =
    `${parts.join(" + ")} — these Excel files need fixing at source.`;

  // group by the workbook the problem belongs to
  const byFile = new Map();
  lastIssues.forEach((i) => {
    const key = i.file || i.supplier || "(unknown file)";
    if (!byFile.has(key)) byFile.set(key, []);
    byFile.get(key).push(i);
  });

  const box = $("#issue-list");
  box.innerHTML = "";
  for (const [file, items] of byFile) {
    const worst = items.some((i) => i.severity === "error") ? "error" : "warning";
    const card = document.createElement("div");
    card.className = "issue-card";
    card.dataset.severity = worst;

    const head = document.createElement("div");
    head.className = "issue-file";
    const name = document.createElement("b");
    name.textContent = file;
    head.appendChild(name);
    const sup = items.find((i) => i.supplier)?.supplier;
    if (sup) {
      const tag = document.createElement("span");
      tag.className = "issue-supplier";
      tag.textContent = sup;
      head.appendChild(tag);
    }
    card.appendChild(head);

    const path = items.find((i) => i.path)?.path;
    if (path) {
      const p = document.createElement("div");
      p.className = "issue-path";
      p.textContent = path;
      p.title = "Full path to the file you need to open";
      card.appendChild(p);
    }

    items.forEach((i) => {
      const row = document.createElement("div");
      row.className = "issue-row";
      row.dataset.severity = i.severity;
      const label = document.createElement("span");
      label.className = "issue-kind";
      label.textContent = i.label + (i.sheet ? ` · ${i.sheet}` : "");
      const detail = document.createElement("span");
      detail.className = "issue-detail";
      detail.textContent = i.detail;
      row.append(label, detail);
      card.appendChild(row);
    });

    box.appendChild(card);
  }
  showModal("modal-issues");
  return true;
}

function issuesAsText() {
  return lastIssues
    .map((i) => [
      `[${i.severity.toUpperCase()}] ${i.label}`,
      i.file ? `  File: ${i.file}` : "",
      i.path ? `  Path: ${i.path}` : "",
      i.sheet ? `  Sheet: ${i.sheet}` : "",
      i.supplier ? `  Supplier: ${i.supplier}` : "",
      `  ${i.detail}`,
    ].filter(Boolean).join("\n"))
    .join("\n\n");
}

/* ── blank-field warning ──────────────────────────────────────────── */
function openBlanksModal(blanks) {
  const list = $("#blank-list");
  list.innerHTML = "";
  Object.keys(blanks).sort().forEach((sup) => {
    const row = document.createElement("div");
    row.className = "blank-item";
    const b = document.createElement("b");
    b.textContent = sup;
    const span = document.createElement("span");
    span.textContent = (blanks[sup] || []).join(", ");
    row.append(b, span);
    list.appendChild(row);
  });
  showModal("modal-blanks");
}

/* ── folder browser ───────────────────────────────────────────────── */
async function openBrowse(key) {
  state.browseKey = key;
  const cfg = state.config || {};
  const title = (cfg.segment_title || {})[key] || key;
  const keyword = (cfg.segment_keyword || {})[key] || "Master price table";
  $("#browse-target").textContent = title;
  $("#browse-hint").innerHTML =
    `Pick the folder that <b>contains</b> a <code>${keyword}_…</code> sub-folder ` +
    `&mdash; not that sub-folder itself.`;
  showModal("modal-browse");
  await navigate($("#" + key).value.trim());
}

async function navigate(path) {
  let data;
  try {
    data = await api("/api/browse?path=" + encodeURIComponent(path || ""));
  } catch (err) { toast(err.message); return; }
  renderBrowse(data);
}

function renderBrowse(d) {
  state.browsePath = d.path || "";
  $("#browse-path").value = d.path || "";
  $("#browse-up").disabled = !d.parent;
  $("#browse-up").dataset.parent = d.parent || "";

  const crumbs = $("#browse-crumbs");
  crumbs.innerHTML = "";
  (d.crumbs || []).forEach((c, i, arr) => {
    if (i) {
      const sep = document.createElement("span");
      sep.className = "crumb-sep";
      sep.textContent = "/";
      crumbs.appendChild(sep);
    }
    const el = document.createElement("span");
    el.className = "crumb" + (i === arr.length - 1 ? " last" : "");
    el.textContent = c.name;
    el.onclick = () => navigate(c.path);
    crumbs.appendChild(el);
  });
  crumbs.scrollLeft = crumbs.scrollWidth;

  const drives = $("#browse-drives");
  drives.innerHTML = "";
  (d.drives || []).forEach((drv) => {
    const el = document.createElement("span");
    el.className = "drive";
    el.textContent = drv;
    el.onclick = () => navigate(drv);
    drives.appendChild(el);
  });

  const dirs = $("#browse-dirs");
  dirs.innerHTML = "";
  if (d.error) {
    const e = document.createElement("div");
    e.className = "empty";
    e.textContent = d.error;
    dirs.appendChild(e);
    return;
  }
  if (!(d.dirs || []).length) {
    const e = document.createElement("div");
    e.className = "empty";
    e.textContent = d.path ? "No sub-folders." : "Pick a drive to start.";
    dirs.appendChild(e);
    return;
  }
  const keyword = ((state.config || {}).segment_keyword || {})[state.browseKey] || "";
  const low = keyword.toLowerCase();
  d.dirs.forEach((name) => {
    const el = document.createElement("div");
    el.className = "dir" + (low && name.toLowerCase().includes(low) ? " match" : "");
    el.textContent = name;
    el.onclick = () => navigate(`${d.path.replace(/\/$/, "")}/${name}`);
    dirs.appendChild(el);
  });
  updateBrowseVerdict(d.dirs, keyword);
}

function updateBrowseVerdict(dirNames, keyword) {
  const el = $("#browse-verdict");
  const low = (keyword || "").toLowerCase();
  const hit = low && (dirNames || []).find((n) => n.toLowerCase().includes(low));
  if (!state.browsePath) {
    el.dataset.state = "";
    $(".mark", el).textContent = "·";
    $(".detail", el).textContent = "";
  } else if (hit) {
    el.dataset.state = "ok";
    $(".mark", el).textContent = "✓";
    $(".detail", el).textContent = `Contains “${hit}” — this is the folder to select`;
  } else {
    el.dataset.state = "bad";
    $(".mark", el).textContent = "✕";
    $(".detail", el).textContent = `No “${keyword}_…” sub-folder here`;
  }
}

function applyBrowse() {
  const value = $("#browse-path").value.trim() || state.browsePath;
  if (state.browseKey && value) {
    $("#" + state.browseKey).value = value;
    saveConfigSoon();
  }
  closeModal("modal-browse");
}

/* ── wiring ───────────────────────────────────────────────────────── */
function wire() {
  PATH_KEYS.forEach((k) => { $("#" + k).addEventListener("input", saveConfigSoon); });
  $$("[data-browse]").forEach((b) => {
    b.onclick = () => openBrowse(b.dataset.browse);
  });
  $$("[data-close]").forEach((b) => { b.onclick = () => closeModal(); });
  $("#overlay").onclick = () => closeModal();

  $("#btn-consolidate").onclick = () => startConsolidate(false);
  $("#btn-runall").onclick = () => startConsolidate(true);

  $("#btn-formdata").onclick = async () => {
    const cfg = await loadConfig();
    if (!cfg.rebate_raw_exists) {
      toast("rebate raw.xlsx not found. Run “Consolidate Rebate Data” and select a FY first.");
      return;
    }
    state.runAll = false;
    openQuarterModal();
  };

  $("#btn-report").onclick = async () => {
    const cfg = await loadConfig();
    if (!cfg.form_input_exists) {
      toast("“rebate form input” folder not found. Run “Generate Form Data” first.");
      return;
    }
    state.runAll = false;
    await openReportModal();
  };

  $("#btn-clear").onclick = () => api("/api/log/clear", {}).catch(() => {});
  $("#btn-copy").onclick = async () => {
    const text = $$("#log .line").map((l) => l.textContent).join("\n");
    try {
      await navigator.clipboard.writeText(text);
      toast("Log copied.", "ok", 2500);
    } catch (_) { toast("Clipboard blocked by the browser."); }
  };

  $("#fy-select").onchange = updateFyInfo;
  $("#fy-generate").onclick = runStage5;
  $("#quarter-generate").onclick = runStage6;
  $("#report-generate").onclick = runStage7;

  $("#sup-all").onchange = (e) => {
    $$(".sup-check").forEach((c) => { c.checked = e.target.checked; });
  };

  $("#browse-up").onclick = (e) => {
    const p = e.currentTarget.dataset.parent;
    navigate(p || "");
  };
  // Settings menu
  const menu = $("#settings-menu");
  const gear = $("#btn-settings");
  const closeMenu = () => { menu.hidden = true; gear.setAttribute("aria-expanded", "false"); };
  gear.onclick = (e) => {
    e.stopPropagation();
    const open = menu.hidden;
    menu.hidden = !open;
    gear.setAttribute("aria-expanded", String(open));
  };
  document.addEventListener("click", (e) => {
    if (!menu.hidden && !e.target.closest(".settings-wrap")) closeMenu();
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeMenu(); });
  $$("[data-settings]").forEach((b) => {
    b.onclick = () => {
      closeMenu();
      const which = b.dataset.settings;
      if (which === "contract-info") openContractInfo();
      else if (which === "exit") confirmExit();
      else openTemplates();
    };
  });

  $("#btn-open-work").onclick = async () => {
    try {
      const res = await api("/api/open-folder", { which: "work" });
      toast(res.path, "ok", 5000);
    } catch (err) { toast(err.message); }
  };

  $("#btn-tracker").onclick = openTracker;
  $("#tracker-refresh").onclick = refreshTracker;

  $("#closed-restart").onclick = () => { $("#closed-hint").hidden = false; };

  // Idle shutdown: any real interaction counts as activity.
  ["mousemove", "mousedown", "keydown", "wheel", "touchstart", "focus"]
    .forEach((evt) => window.addEventListener(evt, markActive,
                                              { passive: true, capture: true }));
  setInterval(idleTick, IDLE_TICK_MS);

  $("#ci-add").onclick = () => {
    const tr = ciAddRow(null);
    ciUpdateNote();
    const first = tr.querySelector("input");
    if (first) first.focus();
  };
  $("#ci-add-col").onclick = ciAddColumn;
  $("#ci-save").onclick = saveContractInfo;

  $("#tpl-expand").onclick = () => {
    const modal = $("#modal-template");
    const on = modal.classList.toggle("expanded");
    $("#tpl-expand").innerHTML = on ? "&#8600; Collapse" : "&#8599; Expand";
  };

  $("#tpl-upload").onclick = () => $("#tpl-file").click();
  $("#tpl-file").onchange = (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    uploadTemplate(file);
  };

  $("#issues-copy").onclick = async () => {
    try {
      await navigator.clipboard.writeText(issuesAsText());
      toast("Problem list copied.", "ok", 2500);
    } catch (_) { toast("Clipboard blocked by the browser."); }
  };

  $("#browse-select").onclick = applyBrowse;
  $("#browse-path").onkeydown = (e) => {
    if (e.key === "Enter") { e.preventDefault(); navigate($("#browse-path").value.trim()); }
  };
}

/* ── boot ─────────────────────────────────────────────────────────── */
/* Tell the local server what this page actually looks like once it has
   booted. Written to data/last-client.json so a bad launch can be diagnosed
   afterwards. Never leaves this machine. */
function reportClientState(configLoaded, err) {
  const visibleModals = $$(".modal")
    .filter((m) => !m.hidden && m.offsetParent !== null)
    .map((m) => m.id);
  const pathsFilled = {};
  PATH_KEYS.forEach((k) => { pathsFilled[k] = !!($("#" + k) || {}).value; });
  const payload = {
    url: location.href,
    configLoaded,
    error: err || null,
    visibleModals,
    overlayVisible: !$("#overlay").hidden,
    pathsFilled,
    stylesheetApplied: getComputedStyle(document.body).backgroundColor,
    modalCount: $$(".modal").length,
    userAgent: navigator.userAgent,
  };
  fetch("/api/client-report", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  }).catch(() => {});
}

(async function boot() {
  let configLoaded = false;
  let bootError = null;
  wire();
  try {
    await loadConfig();
    configLoaded = true;
    await checkPaths();
  } catch (err) {
    bootError = String(err && err.message ? err.message : err);
    toast("Failed to load config: " + bootError);
  }
  connect();
  setTimeout(() => reportClientState(configLoaded, bootError), 400);
})();
