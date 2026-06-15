import { loadState, saveState, getDaysInMonth } from "./store.js";
import {
  parsePreferenceSheet,
  buildTemplateWorkbook,
  exportShiftWorkbook,
  downloadWorkbook,
  readWorkbookFromFile,
} from "./excel.js";
import { generateShift, formatCellDisplay } from "./scheduler.js";

let state = loadState();

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

function init() {
  bindTabs();
  bindBasic();
  bindWorkers();
  bindTeams();
  bindConstraints();
  bindExcel();
  bindShift();
  renderAll();
}

function bindTabs() {
  $$(".tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      $$(".tab").forEach((b) => b.classList.remove("active"));
      $$(".panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      $(`#panel-${btn.dataset.tab}`).classList.add("active");
    });
  });
}

function bindBasic() {
  $("#use-shift-types").addEventListener("change", () => {
    state.useShiftTypes = $("#use-shift-types").checked;
    $("#shift-types-block").classList.toggle("hidden", !state.useShiftTypes);
    persist();
  });

  $("#add-shift-type").addEventListener("click", () => {
    const name = $("#new-shift-type").value.trim();
    if (!name) return;
    if (!state.shiftTypes.includes(name)) state.shiftTypes.push(name);
    $("#new-shift-type").value = "";
    renderShiftTypes();
    persist();
  });

  $("#save-settings").addEventListener("click", () => {
    readBasicFromForm();
    persist();
    alert("設定を保存しました。");
  });

  ["target-year", "target-month", "max-consecutive-work"].forEach((id) => {
    document.getElementById(id).addEventListener("change", () => {
      readBasicFromForm();
      persist();
    });
  });
}

function bindWorkers() {
  $("#add-worker").addEventListener("click", () => {
    const name = $("#new-worker-name").value.trim();
    if (!name) return;
    if (state.workers.some((w) => w.name === name)) {
      alert("同じ名前の勤務者が既にいます。");
      return;
    }
    const teamSel = $("#new-worker-team");
    const teamId = teamSel.value || null;
    state.workers.push({
      id: crypto.randomUUID(),
      name,
      teamId,
      isSupervisor: false,
      monthlyOffDays: 8,
    });
    $("#new-worker-name").value = "";
    renderWorkers();
    renderTeams();
    persist();
  });
}

function bindTeams() {
  $("#add-team").addEventListener("click", () => {
    const name = $("#new-team-name").value.trim();
    if (!name) return;
    if (state.teams.some((t) => t.name === name)) {
      alert("同じチーム名があります。");
      return;
    }
    const id = crypto.randomUUID();
    state.teams.push({ id, name });
    state.teamConstraints[id] = { min: 0, max: 99 };
    $("#new-team-name").value = "";
    renderTeams();
    renderWorkers();
    renderTeamAssignment();
    persist();
  });
}

function assignWorkerToTeam(workerId, teamId) {
  const worker = state.workers.find((w) => w.id === workerId);
  if (!worker) return;
  worker.teamId = teamId || null;
  persist();
  renderWorkers();
  renderTeams();
  renderTeamAssignment();
}

function getWorkersInTeam(teamId) {
  return state.workers.filter((w) => w.teamId === teamId);
}

function createTeamSelect(worker, extraClass = "") {
  const teamSel = document.createElement("select");
  teamSel.className = `team-select ${extraClass}`.trim();
  const optNone = document.createElement("option");
  optNone.value = "";
  optNone.textContent = "（未所属）";
  teamSel.appendChild(optNone);
  state.teams.forEach((t) => {
    const o = document.createElement("option");
    o.value = t.id;
    o.textContent = t.name;
    if (worker.teamId === t.id) o.selected = true;
    teamSel.appendChild(o);
  });
  if (state.teams.length === 0) {
    teamSel.disabled = true;
    teamSel.title = "先にチームを作成してください";
  }
  teamSel.addEventListener("change", () => {
    assignWorkerToTeam(worker.id, teamSel.value || null);
  });
  return teamSel;
}

function bindConstraints() {
  ["daily-min", "daily-max", "supervisor-min", "supervisor-max"].forEach((id) => {
    document.getElementById(id).addEventListener("change", () => {
      readConstraintsFromForm();
      persist();
    });
  });
}

function bindExcel() {
  $("#download-template").addEventListener("click", () => {
    readBasicFromForm();
    const wb = buildTemplateWorkbook(state);
    downloadWorkbook(wb, `休み希望テンプレート_${state.year}年${state.month}月.xlsx`);
  });

  $("#excel-upload").addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const status = $("#excel-status");
    status.textContent = "読み込み中…";
    status.className = "status-msg";

    try {
      readBasicFromForm();
      const wb = await readWorkbookFromFile(file);
      const names = state.workers.map((w) => w.name);
      const { preferences, warnings } = parsePreferenceSheet(wb, names, state.year, state.month);
      state.preferences = preferences;
      persist();
      renderPreferencePreview();

      if (warnings.length) {
        status.textContent = `読み込みました（警告: ${warnings.join(" / ")}）`;
        status.className = "status-msg warn";
      } else {
        status.textContent = "休み希望を読み込みました。";
        status.className = "status-msg success";
      }
    } catch (err) {
      status.textContent = `読み込みに失敗しました: ${err.message}`;
      status.className = "status-msg error";
    }
    e.target.value = "";
  });
}

function bindShift() {
  $("#generate-shift").addEventListener("click", () => {
    readAllFromForm();
    persist();
    const status = $("#generate-status");
    status.textContent = "生成中…";
    status.className = "status-msg";

    setTimeout(() => {
      const result = generateShift(state);
      if (!result.ok) {
        status.textContent = result.messages.join(" ");
        status.className = "status-msg error";
        return;
      }

      state.lastResult = result;
      persist();
      status.textContent = result.messages.join(" ");
      status.className = result.messages.some((m) => m.includes("超え"))
        ? "status-msg warn"
        : "status-msg success";
      renderShiftResult(result);
      $("#export-shift").disabled = false;
      $("#print-shift").disabled = false;
    }, 30);
  });

  $("#print-shift").addEventListener("click", () => {
    if (!state.lastResult) return;
    prepareShiftPrintScale();
    window.print();
  });

  window.addEventListener("beforeprint", () => {
    if (state.lastResult) prepareShiftPrintScale();
  });

  window.addEventListener("afterprint", () => {
    $("#shift-print-root")?.style.removeProperty("--print-scale");
  });

  $("#export-shift").addEventListener("click", () => {
    if (!state.lastResult) return;
    const wb = exportShiftWorkbook(state.lastResult, state);
    downloadWorkbook(wb, `シフト_${state.year}年${state.month}月.xlsx`);
  });

  $("#clear-result").addEventListener("click", () => {
    state.lastResult = null;
    persist();
    $("#shift-print-root").classList.add("hidden");
    $("#shift-table-wrap").classList.add("hidden");
    $("#shift-summary").classList.add("hidden");
    $("#generate-status").textContent = "";
    $("#export-shift").disabled = true;
    $("#print-shift").disabled = true;
  });
}

/** A4横・1枚に収まるよう印刷用スケールを計算 */
function prepareShiftPrintScale() {
  const root = $("#shift-print-root");
  if (!root) return;

  root.style.setProperty("--print-scale", "1");
  const widthMm = 297 - 16;
  const heightMm = 210 - 16;
  const pxPerMm = 96 / 25.4;
  const maxW = widthMm * pxPerMm;
  const maxH = heightMm * pxPerMm;

  const scaleW = maxW / root.scrollWidth;
  const scaleH = maxH / root.scrollHeight;
  const scale = Math.min(scaleW, scaleH, 1);
  root.style.setProperty("--print-scale", String(scale));
}

function readBasicFromForm() {
  state.year = parseInt($("#target-year").value, 10);
  state.month = parseInt($("#target-month").value, 10);
  state.maxConsecutiveWork = parseInt($("#max-consecutive-work").value, 10);
  state.useShiftTypes = $("#use-shift-types").checked;
}

function readConstraintsFromForm() {
  state.constraints.dailyMin = parseInt($("#daily-min").value, 10);
  state.constraints.dailyMax = parseInt($("#daily-max").value, 10);
  state.constraints.supervisorMin = parseInt($("#supervisor-min").value, 10);
  state.constraints.supervisorMax = parseInt($("#supervisor-max").value, 10);
}

function readAllFromForm() {
  readBasicFromForm();
  readConstraintsFromForm();
}

function persist() {
  saveState(state);
}

function renderAll() {
  renderBasic();
  renderShiftTypes();
  renderWorkers();
  renderTeams();
  renderTeamAssignment();
  renderConstraints();
  renderPreferencePreview();
  if (state.lastResult) {
    renderShiftResult(state.lastResult);
    $("#export-shift").disabled = false;
  }
}

function renderBasic() {
  $("#target-year").value = state.year;
  $("#target-month").value = state.month;
  $("#max-consecutive-work").value = state.maxConsecutiveWork;
  $("#use-shift-types").checked = state.useShiftTypes;
  $("#shift-types-block").classList.toggle("hidden", !state.useShiftTypes);
}

function renderShiftTypes() {
  const ul = $("#shift-type-list");
  ul.innerHTML = "";
  state.shiftTypes.forEach((name, idx) => {
    const li = document.createElement("li");
    li.textContent = name;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = "×";
    btn.title = "削除";
    btn.addEventListener("click", () => {
      state.shiftTypes.splice(idx, 1);
      renderShiftTypes();
      persist();
    });
    li.appendChild(btn);
    ul.appendChild(li);
  });
}

function renderWorkers() {
  const tbody = $("#workers-tbody");
  tbody.innerHTML = "";
  state.workers.forEach((w, idx) => {
    const tr = document.createElement("tr");

    const nameTd = document.createElement("td");
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.value = w.name;
    nameInput.addEventListener("change", () => {
      const old = w.name;
      w.name = nameInput.value.trim() || w.name;
      if (state.preferences[old]) {
        state.preferences[w.name] = state.preferences[old];
        delete state.preferences[old];
      }
      persist();
    });
    nameTd.appendChild(nameInput);

    const teamTd = document.createElement("td");
    teamTd.appendChild(createTeamSelect(w));

    const supTd = document.createElement("td");
    const supChk = document.createElement("input");
    supChk.type = "checkbox";
    supChk.checked = w.isSupervisor;
    supChk.addEventListener("change", () => {
      w.isSupervisor = supChk.checked;
      persist();
    });
    supTd.appendChild(supChk);

    const offTd = document.createElement("td");
    const offInput = document.createElement("input");
    offInput.type = "number";
    offInput.min = 0;
    offInput.max = 31;
    offInput.value = w.monthlyOffDays;
    offInput.addEventListener("change", () => {
      w.monthlyOffDays = parseInt(offInput.value, 10) || 0;
      persist();
    });
    offTd.appendChild(offInput);

    const actTd = document.createElement("td");
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "btn-icon";
    delBtn.textContent = "削除";
    delBtn.addEventListener("click", () => {
      if (!confirm(`${w.name} を削除しますか？`)) return;
      delete state.preferences[w.name];
      state.workers.splice(idx, 1);
      renderWorkers();
      renderTeams();
      renderTeamAssignment();
      persist();
    });
    actTd.appendChild(delBtn);

    tr.append(nameTd, teamTd, supTd, offTd, actTd);
    tbody.appendChild(tr);
  });

  const hint = $("#workers-team-hint");
  if (!state.teams.length) {
    hint.textContent = "チームがまだありません。「チーム」タブでチームを作成し、メンバーを割り振ってください。";
    hint.classList.remove("hidden");
  } else {
    hint.classList.add("hidden");
  }
  renderNewWorkerTeamSelect();
  renderTeamWorkerStats();
}

function renderNewWorkerTeamSelect() {
  const sel = $("#new-worker-team");
  if (!sel) return;
  const prev = sel.value;
  sel.innerHTML = "";
  const optNone = document.createElement("option");
  optNone.value = "";
  optNone.textContent = "（未所属）";
  sel.appendChild(optNone);
  state.teams.forEach((t) => {
    const o = document.createElement("option");
    o.value = t.id;
    o.textContent = t.name;
    sel.appendChild(o);
  });
  if (prev && [...sel.options].some((o) => o.value === prev)) {
    sel.value = prev;
  } else if (state.teams[0]) {
    sel.value = state.teams[0].id;
  }
  sel.disabled = state.teams.length === 0;
}

function renderTeamAssignment() {
  renderTeamWorkerStats();

  const tbody = $("#team-assignment-tbody");
  if (!tbody) return;

  tbody.innerHTML = "";
  state.workers.forEach((w) => {
    const tr = document.createElement("tr");
    const nameTd = document.createElement("td");
    nameTd.textContent = w.name;
    const teamTd = document.createElement("td");
    teamTd.appendChild(createTeamSelect(w));
    tr.append(nameTd, teamTd);
    tbody.appendChild(tr);
  });

  const noWorkers = $("#teams-assignment-empty");
  const noTeams = $("#teams-no-team-hint");
  const matrixWrap = $("#team-assignment-matrix");

  if (!state.workers.length) {
    noWorkers?.classList.remove("hidden");
    $("#team-assignment-list-wrap")?.classList.add("hidden");
    matrixWrap?.classList.add("hidden");
    return;
  }
  noWorkers?.classList.add("hidden");
  $("#team-assignment-list-wrap")?.classList.remove("hidden");

  if (!state.teams.length) {
    noTeams?.classList.remove("hidden");
    matrixWrap?.classList.add("hidden");
    return;
  }
  noTeams?.classList.add("hidden");

  renderTeamAssignmentMatrix(matrixWrap);
}

function renderTeamAssignmentMatrix(wrap) {
  if (!wrap) return;
  let html =
    "<table class='data-table team-matrix-table'><thead><tr><th>勤務者</th>";
  state.teams.forEach((t) => {
    html += `<th>${escapeHtml(t.name)}</th>`;
  });
  html += "<th>未所属</th></tr></thead><tbody>";

  state.workers.forEach((w) => {
    html += `<tr><td>${escapeHtml(w.name)}</td>`;
    state.teams.forEach((t) => {
      const checked = w.teamId === t.id ? "checked" : "";
      html += `<td><input type="radio" name="team-${escapeHtml(w.id)}" value="${escapeHtml(t.id)}" ${checked} aria-label="${escapeHtml(w.name)} → ${escapeHtml(t.name)}" /></td>`;
    });
    const noneChecked = !w.teamId ? "checked" : "";
    html += `<td><input type="radio" name="team-${escapeHtml(w.id)}" value="" ${noneChecked} aria-label="${escapeHtml(w.name)} → 未所属" /></td></tr>`;
  });
  html += "</tbody></table>";
  wrap.innerHTML = html;
  wrap.classList.remove("hidden");

  wrap.querySelectorAll('input[type="radio"]').forEach((input) => {
    input.addEventListener("change", () => {
      if (!input.checked) return;
      const workerId = input.name.replace(/^team-/, "");
      assignWorkerToTeam(workerId, input.value || null);
    });
  });
}

function renderTeamWorkerStats() {
  const total = state.workers.length;
  const assigned = state.workers.filter((w) => w.teamId).length;
  const unassigned = total - assigned;

  const totalEl = $("#workers-total-count");
  const assignedEl = $("#workers-assigned-count");
  const unassignedEl = $("#workers-unassigned-count");
  if (totalEl) totalEl.textContent = String(total);
  if (assignedEl) assignedEl.textContent = String(assigned);
  if (unassignedEl) unassignedEl.textContent = String(unassigned);
}

function renderTeams() {
  renderTeamWorkerStats();

  const tbody = $("#teams-tbody");
  tbody.innerHTML = "";
  state.teams.forEach((t, idx) => {
    const tr = document.createElement("tr");
    const tc = state.teamConstraints[t.id] || { min: 0, max: 99 };

    const nameTd = document.createElement("td");
    const nameInput = document.createElement("input");
    nameInput.value = t.name;
    nameInput.addEventListener("change", () => {
      t.name = nameInput.value.trim() || t.name;
      renderWorkers();
      persist();
    });
    nameTd.appendChild(nameInput);

    const membersTd = document.createElement("td");
    membersTd.className = "team-member-names";
    const members = getWorkersInTeam(t.id);
    membersTd.textContent = members.length ? members.map((w) => w.name).join("、") : "—";
    membersTd.title = members.length ? members.map((w) => w.name).join("\n") : "メンバー未割り当て";

    const minTd = document.createElement("td");
    const minInput = document.createElement("input");
    minInput.type = "number";
    minInput.min = 0;
    minInput.value = tc.min;
    minInput.addEventListener("change", () => {
      state.teamConstraints[t.id].min = parseInt(minInput.value, 10) || 0;
      persist();
    });
    minTd.appendChild(minInput);

    const maxTd = document.createElement("td");
    const maxInput = document.createElement("input");
    maxInput.type = "number";
    maxInput.min = 0;
    maxInput.value = tc.max;
    maxInput.addEventListener("change", () => {
      state.teamConstraints[t.id].max = parseInt(maxInput.value, 10) || 0;
      persist();
    });
    maxTd.appendChild(maxInput);

    const actTd = document.createElement("td");
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "btn-icon";
    delBtn.textContent = "削除";
    delBtn.addEventListener("click", () => {
      if (!confirm(`チーム「${t.name}」を削除しますか？`)) return;
      state.workers.forEach((w) => {
        if (w.teamId === t.id) w.teamId = null;
      });
      delete state.teamConstraints[t.id];
      state.teams.splice(idx, 1);
      renderTeams();
      renderWorkers();
      renderTeamAssignment();
      persist();
    });
    actTd.appendChild(delBtn);

    tr.append(nameTd, membersTd, minTd, maxTd, actTd);
    tbody.appendChild(tr);
  });

  renderTeamAssignment();
  renderNewWorkerTeamSelect();
}

function renderConstraints() {
  $("#daily-min").value = state.constraints.dailyMin;
  $("#daily-max").value = state.constraints.dailyMax;
  $("#supervisor-min").value = state.constraints.supervisorMin;
  $("#supervisor-max").value = state.constraints.supervisorMax;
}

function renderPreferencePreview() {
  const wrap = $("#preference-preview");
  const prefs = state.preferences;
  const keys = Object.keys(prefs);
  if (!keys.length) {
    wrap.classList.add("hidden");
    wrap.innerHTML = "";
    return;
  }

  const days = getDaysInMonth(state.year, state.month);
  let html = "<table class='data-table'><thead><tr><th>勤務者</th>";
  for (let d = 1; d <= Math.min(days, 10); d++) html += `<th>${d}</th>`;
  if (days > 10) html += "<th>…</th>";
  html += "</tr></thead><tbody>";

  keys.slice(0, 8).forEach((name) => {
    html += `<tr><td>${escapeHtml(name)}</td>`;
    for (let d = 1; d <= Math.min(days, 10); d++) {
      html += `<td>${prefs[name][d] ? "休" : ""}</td>`;
    }
    if (days > 10) html += "<td></td>";
    html += "</tr>";
  });
  html += "</tbody></table>";
  if (keys.length > 8) html += `<p class="hint">他 ${keys.length - 8} 名</p>`;
  wrap.innerHTML = html;
  wrap.classList.remove("hidden");
}

function renderShiftResult(result) {
  const { year, month, assignments, workers, stats } = result;
  const days = getDaysInMonth(year, month);
  const useTypes = state.useShiftTypes;

  const summary = $("#shift-summary");
  summary.innerHTML = "";
  const okDays = stats.daily.filter(
    (d) =>
      d.total >= state.constraints.dailyMin &&
      d.total <= state.constraints.dailyMax
  ).length;
  summary.innerHTML = `
    <div class="card-stat">${year}年${month}月 · ${days}日間</div>
    <div class="card-stat">勤務者 ${workers.length} 名</div>
    <div class="card-stat">日別人数制約 OK: ${okDays}/${days} 日</div>
  `;
  summary.classList.remove("hidden");

  const wrap = $("#shift-table-wrap");
  let html = "<table class='data-table'><thead><tr><th class='sticky-col'>勤務者</th>";
  for (let d = 1; d <= days; d++) {
    const dow = ["日", "月", "火", "水", "木", "金", "土"][new Date(year, month - 1, d).getDay()];
    html += `<th>${d}<br><small>${dow}</small></th>`;
  }
  html += "</tr></thead><tbody>";

  workers.forEach((w) => {
    html += `<tr><td class="sticky-col">${escapeHtml(w.name)}${w.isSupervisor ? " ★" : ""}</td>`;
    for (let d = 1; d <= days; d++) {
      const cell = assignments[w.id][d];
      const label = formatCellDisplay(cell, useTypes);
      let cls = "cell-off";
      if (cell?.type === "work") {
        cls = w.isSupervisor ? "cell-supervisor" : "cell-work";
      }
      html += `<td class="${cls}" title="${escapeHtml(label)}">${escapeHtml(label)}</td>`;
    }
    html += "</tr>";
  });
  html += "</tbody></table>";
  wrap.innerHTML = html;
  wrap.classList.remove("hidden");
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

init();
