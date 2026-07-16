import { loadState, saveState, getDaysInMonth, APP_VERSION } from "./store.js";
import {
  getWorkerSections,
  moveWorkerWithinSection,
  moveTeamOrder,
  moveSubGroupOrder,
  canMoveSubGroupUp,
  canMoveSubGroupDown,
} from "./worker-groups.js";
import {
  parsePreferenceSheet,
  buildTemplateWorkbook,
  exportShiftWorkbook,
  downloadWorkbook,
  readWorkbookFromFile,
  normalizePreferences,
  formatPreferencePreview,
} from "./excel.js";
import {
  generateShift,
  formatCellDisplay,
  countWorkerOffDaysFromAssignments,
  getConferenceTeamsOnDay,
  getConferenceSubGroupsOnDay,
  getWorkerConferenceGroup,
  getConferenceGroupStyle,
  buildConferenceColorMap,
  normalizeConferenceDays,
} from "./scheduler.js";

let state = loadState();
state.preferences = normalizePreferences(state.preferences);

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

function init() {
  try {
    if (location.protocol === "file:") {
      showBootError(
        "このアプリは file:// では動作しません。ターミナルで npm start を実行し、http://localhost:3456 を開いてください。"
      );
      return;
    }

    bindTabs();
    bindBasic();
    bindWorkers();
    bindTeams();
    bindConstraints();
    bindExcel();
    bindShift();
    renderAll();
    showAppVersion();
  } catch (err) {
    console.error(err);
    showBootError(`アプリの起動に失敗しました: ${err.message}`);
  }
}

function showBootError(message) {
  const box = document.createElement("div");
  box.className = "boot-error";
  box.textContent = message;
  document.body.prepend(box);
}

function showAppVersion() {
  const footer = document.querySelector(".app-footer p");
  if (footer) {
    footer.textContent = `データはこのブラウザのローカルストレージに保存されます。バージョン ${APP_VERSION}`;
  }
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
    readAllFromForm();
    persist();
    alert("設定を保存しました。");
  });

  ["target-year", "target-month", "max-consecutive-work"].forEach((id) => {
    document.getElementById(id).addEventListener("change", () => {
      readAllFromForm();
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
    const subSel = $("#new-worker-subgroup");
    const teamId = teamSel.value || null;
    const subGroupId = subSel.value || null;
    state.workers.push({
      id: crypto.randomUUID(),
      name,
      teamId,
      subGroupId: teamId && subGroupId ? subGroupId : null,
      isSupervisor: false,
      isGroupSupervisor: false,
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
      alert("同じメイングループ名があります。");
      return;
    }
    const id = crypto.randomUUID();
    state.teams.push({ id, name });
    state.teamConstraints[id] = {
      min: 0,
      max: 99,
      useConferenceDay: false,
      supervisorMin: 0,
      supervisorMax: 99,
    };
    $("#new-team-name").value = "";
    renderTeams();
    renderWorkers();
    renderTeamAssignment();
    persist();
  });

  $("#add-subgroup").addEventListener("click", () => {
    const teamId = $("#new-subgroup-team").value;
    const name = $("#new-subgroup-name").value.trim();
    if (!teamId) {
      alert("メイングループを選択してください。");
      return;
    }
    if (!name) return;
    if (state.subGroups.some((sg) => sg.teamId === teamId && sg.name === name)) {
      alert("同じサブグループ名があります。");
      return;
    }
    const id = crypto.randomUUID();
    state.subGroups.push({ id, name, teamId });
    state.subGroupConstraints[id] = { min: 0, max: 99, useConferenceDay: false };
    $("#new-subgroup-name").value = "";
    renderTeams();
    renderWorkers();
    renderTeamAssignment();
    persist();
  });
}

function ensureTeamConstraint(teamId) {
  if (!state.teamConstraints[teamId]) {
    state.teamConstraints[teamId] = {
      min: 0,
      max: 99,
      useConferenceDay: false,
      supervisorMin: 0,
      supervisorMax: 99,
    };
  }
  const tc = state.teamConstraints[teamId];
  if (tc.useConferenceDay == null) tc.useConferenceDay = false;
  if (tc.supervisorMin == null) tc.supervisorMin = 0;
  if (tc.supervisorMax == null) tc.supervisorMax = 99;
  return tc;
}

function ensureSubGroupConstraint(subGroupId) {
  if (!state.subGroupConstraints[subGroupId]) {
    state.subGroupConstraints[subGroupId] = { min: 0, max: 99, useConferenceDay: false };
  }
  if (state.subGroupConstraints[subGroupId].useConferenceDay == null) {
    state.subGroupConstraints[subGroupId].useConferenceDay = false;
  }
  return state.subGroupConstraints[subGroupId];
}

function assignWorkerToTeam(workerId, teamId) {
  const worker = state.workers.find((w) => w.id === workerId);
  if (!worker) return;
  worker.teamId = teamId || null;
  if (!worker.teamId) worker.isGroupSupervisor = false;
  if (worker.subGroupId) {
    const sg = state.subGroups.find((g) => g.id === worker.subGroupId);
    if (!sg || sg.teamId !== worker.teamId) worker.subGroupId = null;
  }
  persist();
  renderWorkers();
  renderTeams();
  renderTeamAssignment();
  renderConstraints();
}

function assignWorkerToSubGroup(workerId, subGroupId) {
  const worker = state.workers.find((w) => w.id === workerId);
  if (!worker) return;
  if (!subGroupId) {
    worker.subGroupId = null;
  } else {
    const sg = state.subGroups.find((g) => g.id === subGroupId);
    if (!sg) return;
    worker.teamId = sg.teamId;
    worker.subGroupId = subGroupId;
  }
  persist();
  renderWorkers();
  renderTeams();
  renderTeamAssignment();
}

function getWorkersInTeam(teamId) {
  return state.workers.filter((w) => w.teamId === teamId);
}

function getWorkersInSubGroup(subGroupId) {
  return state.workers.filter((w) => w.subGroupId === subGroupId);
}

function getSubGroupsForTeam(teamId) {
  return state.subGroups.filter((sg) => sg.teamId === teamId);
}

function deleteSubGroupsForTeam(teamId) {
  const ids = state.subGroups.filter((sg) => sg.teamId === teamId).map((sg) => sg.id);
  state.subGroups = state.subGroups.filter((sg) => sg.teamId !== teamId);
  ids.forEach((id) => delete state.subGroupConstraints[id]);
  state.workers.forEach((w) => {
    if (w.subGroupId && ids.includes(w.subGroupId)) w.subGroupId = null;
  });
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
    teamSel.title = "先にメイングループを作成してください";
  }
  teamSel.addEventListener("change", () => {
    assignWorkerToTeam(worker.id, teamSel.value || null);
  });
  return teamSel;
}

function createSubGroupSelect(worker, extraClass = "") {
  const subSel = document.createElement("select");
  subSel.className = `subgroup-select ${extraClass}`.trim();
  const optNone = document.createElement("option");
  optNone.value = "";
  optNone.textContent = "（未所属）";
  subSel.appendChild(optNone);

  const teamSubs = getSubGroupsForTeam(worker.teamId);
  teamSubs.forEach((sg) => {
    const o = document.createElement("option");
    o.value = sg.id;
    o.textContent = sg.name;
    if (worker.subGroupId === sg.id) o.selected = true;
    subSel.appendChild(o);
  });

  if (!worker.teamId || teamSubs.length === 0) {
    subSel.disabled = true;
    subSel.title = worker.teamId ? "先にサブグループを作成してください" : "先にメイングループを選択してください";
  }

  subSel.addEventListener("change", () => {
    assignWorkerToSubGroup(worker.id, subSel.value || null);
  });
  return subSel;
}

function bindConstraints() {
  ["supervisor-min", "supervisor-max"].forEach((id) => {
    document.getElementById(id).addEventListener("change", () => {
      readConstraintsFromForm();
      renderSupervisorStats();
      persist();
    });
  });
}

function bindExcel() {
  $("#download-template").addEventListener("click", () => {
    readAllFromForm();
    const wb = buildTemplateWorkbook(state);
    downloadWorkbook(wb, `勤務希望テンプレート_${state.year}年${state.month}月.xlsx`);
  });

  $("#excel-upload").addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const status = $("#excel-status");
    status.textContent = "読み込み中…";
    status.className = "status-msg";

    try {
      readAllFromForm();
      const wb = await readWorkbookFromFile(file);
      const names = state.workers.map((w) => w.name);
      const { preferences, monthlyOffDaysByName, warnings } = parsePreferenceSheet(wb, names, state.year, state.month);
      state.preferences = normalizePreferences(preferences);
      let offDaysUpdated = 0;
      for (const w of state.workers) {
        if (monthlyOffDaysByName[w.name] != null) {
          w.monthlyOffDays = monthlyOffDaysByName[w.name];
          offDaysUpdated++;
        }
      }
      persist();
      renderWorkers();
      renderPreferencePreview();

      const statusParts = [];
      if (offDaysUpdated > 0) {
        statusParts.push(`月間休み日数 ${offDaysUpdated}名分を反映`);
      }
      if (warnings.length) {
        status.textContent = `読み込みました（${statusParts.join("、")}${statusParts.length ? " / " : ""}警告: ${warnings.join(" / ")}）`;
        status.className = "status-msg warn";
      } else if (statusParts.length) {
        status.textContent = `休み希望と${statusParts.join("、")}しました。`;
        status.className = "status-msg success";
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
        $("#shift-print-root")?.classList.add("hidden");
        return;
      }

      state.lastResult = { ...result, workers: state.workers.map((w) => ({ ...w })) };
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
  state.constraints.supervisorMin = parseInt($("#supervisor-min").value, 10);
  state.constraints.supervisorMax = parseInt($("#supervisor-max").value, 10);
}

function readWorkersFromForm() {
  const tbody = $("#workers-tbody");
  if (!tbody) return;

  tbody.querySelectorAll("tr[data-worker-id]").forEach((tr) => {
    const w = state.workers.find((x) => x.id === tr.dataset.workerId);
    if (!w) return;

    const nameInput = tr.querySelector('input[type="text"]');
    const offInput = tr.querySelector(".worker-off-days");
    const supChk = tr.querySelector(".chk-supervisor");
    const groupSupChk = tr.querySelector(".chk-group-supervisor");
    const teamSel = tr.querySelector(".team-select");
    const subSel = tr.querySelector(".subgroup-select");

    if (nameInput) {
      const newName = nameInput.value.trim();
      if (newName && newName !== w.name) {
        const old = w.name;
        if (state.preferences[old]) {
          state.preferences[newName] = state.preferences[old];
          delete state.preferences[old];
        }
        w.name = newName;
      }
    }
    if (offInput) w.monthlyOffDays = parseFloat(offInput.value) || 0;
    if (supChk) w.isSupervisor = supChk.checked;
    if (groupSupChk) w.isGroupSupervisor = groupSupChk.checked;
    if (teamSel) {
      w.teamId = teamSel.value || null;
      if (!w.teamId) w.isGroupSupervisor = false;
      if (w.subGroupId) {
        const sg = state.subGroups.find((g) => g.id === w.subGroupId);
        if (!sg || sg.teamId !== w.teamId) w.subGroupId = null;
      }
    }
    if (subSel && !subSel.disabled) w.subGroupId = subSel.value || null;
  });
}

function readAllFromForm() {
  readBasicFromForm();
  readConstraintsFromForm();
  readWorkersFromForm();
}

function persist() {
  if (!saveState(state)) {
    alert("設定の保存に失敗しました。ブラウザのストレージ容量を確認してください。");
  }
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
  readWorkersFromForm();
  const tbody = $("#workers-tbody");
  tbody.innerHTML = "";
  const sections = getWorkerSections(state.workers, state.teams, state.subGroups);

  sections.forEach((section) => {
    const headerTr = document.createElement("tr");
    headerTr.className = "group-header-row";
    const headerTd = document.createElement("td");
    headerTd.colSpan = 8;
    headerTd.textContent = section.label;
    headerTr.appendChild(headerTd);
    tbody.appendChild(headerTr);

    section.members.forEach((w) => {
      tbody.appendChild(createWorkerRow(w, section));
    });
  });

  const hint = $("#workers-team-hint");
  if (!state.teams.length) {
    hint.textContent = "メイングループがまだありません。「グループ」タブでメイングループを作成し、メンバーを割り振ってください。";
    hint.classList.remove("hidden");
  } else {
    hint.classList.add("hidden");
  }
  renderNewWorkerTeamSelect();
  renderNewWorkerSubGroupSelect();
  renderTeamWorkerStats();
  renderSupervisorStats();
}

function createWorkerRow(w, section) {
  const tr = document.createElement("tr");
  tr.dataset.workerId = w.id;

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

  const subTd = document.createElement("td");
  subTd.appendChild(createSubGroupSelect(w));

  const supTd = document.createElement("td");
  const supChk = document.createElement("input");
  supChk.type = "checkbox";
  supChk.className = "chk-supervisor";
  supChk.title = "全体責任者";
  supChk.checked = w.isSupervisor;
  supChk.addEventListener("change", () => {
    w.isSupervisor = supChk.checked;
    renderWorkers();
    renderConstraints();
    persist();
  });
  supTd.appendChild(supChk);

  const groupSupTd = document.createElement("td");
  const groupSupChk = document.createElement("input");
  groupSupChk.type = "checkbox";
  groupSupChk.className = "chk-group-supervisor";
  groupSupChk.title = "所属メイングループの責任者（全体責任者と重複可）";
  groupSupChk.checked = w.isGroupSupervisor;
  groupSupChk.disabled = !w.teamId;
  groupSupChk.addEventListener("change", () => {
    w.isGroupSupervisor = groupSupChk.checked;
    renderWorkers();
    renderTeams();
    renderConstraints();
    persist();
  });
  groupSupTd.appendChild(groupSupChk);

  const offTd = document.createElement("td");
  const offInput = document.createElement("input");
  offInput.className = "worker-off-days";
  offInput.type = "number";
  offInput.min = 0;
  offInput.max = 31;
  offInput.step = 0.5;
  offInput.value = w.monthlyOffDays;
  const syncOffDays = () => {
    w.monthlyOffDays = parseFloat(offInput.value) || 0;
    persist();
  };
  offInput.addEventListener("input", syncOffDays);
  offInput.addEventListener("change", syncOffDays);
  offTd.appendChild(offInput);

  const orderTd = document.createElement("td");
  orderTd.className = "order-cell";
  const memberIdx = section.members.findIndex((m) => m.id === w.id);
  const upBtn = document.createElement("button");
  upBtn.type = "button";
  upBtn.className = "btn-icon";
  upBtn.textContent = "↑";
  upBtn.title = "グループ内で上へ";
  upBtn.disabled = memberIdx <= 0;
  upBtn.addEventListener("click", () => moveWorker(w.id, -1));
  const downBtn = document.createElement("button");
  downBtn.type = "button";
  downBtn.className = "btn-icon";
  downBtn.textContent = "↓";
  downBtn.title = "グループ内で下へ";
  downBtn.disabled = memberIdx >= section.members.length - 1;
  downBtn.addEventListener("click", () => moveWorker(w.id, 1));
  orderTd.append(upBtn, downBtn);

  const actTd = document.createElement("td");
  const delBtn = document.createElement("button");
  delBtn.type = "button";
  delBtn.className = "btn-icon";
  delBtn.textContent = "削除";
  delBtn.addEventListener("click", () => {
    if (!confirm(`${w.name} を削除しますか？`)) return;
    delete state.preferences[w.name];
    state.workers = state.workers.filter((x) => x.id !== w.id);
    renderWorkers();
    renderTeams();
    renderTeamAssignment();
    persist();
  });
  actTd.appendChild(delBtn);

  tr.append(nameTd, teamTd, subTd, supTd, groupSupTd, offTd, orderTd, actTd);
  return tr;
}

function createOrderButtons({ onUp, onDown, upDisabled, downDisabled, upTitle = "上へ", downTitle = "下へ" }) {
  const orderTd = document.createElement("td");
  orderTd.className = "order-cell";
  const upBtn = document.createElement("button");
  upBtn.type = "button";
  upBtn.className = "btn-icon";
  upBtn.textContent = "↑";
  upBtn.title = upTitle;
  upBtn.disabled = upDisabled;
  upBtn.addEventListener("click", onUp);
  const downBtn = document.createElement("button");
  downBtn.type = "button";
  downBtn.className = "btn-icon";
  downBtn.textContent = "↓";
  downBtn.title = downTitle;
  downBtn.disabled = downDisabled;
  downBtn.addEventListener("click", onDown);
  orderTd.append(upBtn, downBtn);
  return orderTd;
}

function moveWorker(workerId, delta) {
  if (!moveWorkerWithinSection(state.workers, workerId, delta, state.teams, state.subGroups)) return;
  renderWorkers();
  renderTeamAssignment();
  if (state.lastResult) renderShiftResult(state.lastResult);
  persist();
}

function moveTeam(idx, delta) {
  if (!moveTeamOrder(state.teams, idx, delta)) return;
  renderTeams();
  renderWorkers();
  renderTeamAssignment();
  persist();
}

function moveSubGroup(subGroupId, delta) {
  if (!moveSubGroupOrder(state.subGroups, subGroupId, delta)) return;
  renderTeams();
  renderWorkers();
  renderTeamAssignment();
  persist();
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
  sel.onchange = () => renderNewWorkerSubGroupSelect();
}

function renderNewWorkerSubGroupSelect() {
  const teamSel = $("#new-worker-team");
  const subSel = $("#new-worker-subgroup");
  if (!subSel) return;
  const teamId = teamSel?.value || "";
  const prev = subSel.value;
  subSel.innerHTML = "";
  const optNone = document.createElement("option");
  optNone.value = "";
  optNone.textContent = "（未所属）";
  subSel.appendChild(optNone);

  if (teamId) {
    getSubGroupsForTeam(teamId).forEach((sg) => {
      const o = document.createElement("option");
      o.value = sg.id;
      o.textContent = sg.name;
      subSel.appendChild(o);
    });
  }

  if (prev && [...subSel.options].some((o) => o.value === prev)) {
    subSel.value = prev;
  }
  subSel.disabled = !teamId || getSubGroupsForTeam(teamId).length === 0;
}

function renderTeamAssignment() {
  renderTeamWorkerStats();

  const tbody = $("#team-assignment-tbody");
  if (!tbody) return;

  tbody.innerHTML = "";
  getWorkerSections(state.workers, state.teams, state.subGroups).forEach((section) => {
    const headerTr = document.createElement("tr");
    headerTr.className = "group-header-row";
    const headerTd = document.createElement("td");
    headerTd.colSpan = 3;
    headerTd.textContent = section.label;
    headerTr.appendChild(headerTd);
    tbody.appendChild(headerTr);

    section.members.forEach((w) => {
      const tr = document.createElement("tr");
      const nameTd = document.createElement("td");
      nameTd.textContent = w.name;
      const teamTd = document.createElement("td");
      teamTd.appendChild(createTeamSelect(w));
      const subTd = document.createElement("td");
      subTd.appendChild(createSubGroupSelect(w));
      tr.append(nameTd, teamTd, subTd);
      tbody.appendChild(tr);
    });
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
    const tc = ensureTeamConstraint(t.id);

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
    const memberLabels = members.map((w) => {
      const marks = [];
      if (w.isSupervisor) marks.push("★");
      if (w.isGroupSupervisor) marks.push("◆");
      return marks.length ? `${w.name}${marks.join("")}` : w.name;
    });
    membersTd.textContent = memberLabels.length ? memberLabels.join("、") : "—";
    membersTd.title = members.length
      ? members
          .map((w) => {
            const marks = [];
            if (w.isSupervisor) marks.push("全体責任者");
            if (w.isGroupSupervisor) marks.push("グループ責任者");
            return marks.length ? `${w.name}（${marks.join("・")}）` : w.name;
          })
          .join("\n")
      : "メンバー未割り当て";

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

    const supMinTd = document.createElement("td");
    const supMinInput = document.createElement("input");
    supMinInput.type = "number";
    supMinInput.min = 0;
    supMinInput.value = tc.supervisorMin ?? 0;
    supMinInput.title = "グループ責任者の1日あたり下限";
    supMinInput.addEventListener("change", () => {
      state.teamConstraints[t.id].supervisorMin = parseInt(supMinInput.value, 10) || 0;
      renderConstraints();
      persist();
    });
    supMinTd.appendChild(supMinInput);

    const supMaxTd = document.createElement("td");
    const supMaxInput = document.createElement("input");
    supMaxInput.type = "number";
    supMaxInput.min = 0;
    supMaxInput.value = tc.supervisorMax ?? 99;
    supMaxInput.title = "グループ責任者の1日あたり上限";
    supMaxInput.addEventListener("change", () => {
      state.teamConstraints[t.id].supervisorMax = parseInt(supMaxInput.value, 10) || 0;
      renderConstraints();
      persist();
    });
    supMaxTd.appendChild(supMaxInput);

    const confTd = document.createElement("td");
    const confChk = document.createElement("input");
    confChk.type = "checkbox";
    confChk.checked = Boolean(tc.useConferenceDay);
    confChk.title = "月4回・隔週優先でカンファレンス日を設定（参加者が多い日を優先）";
    confChk.addEventListener("change", () => {
      tc.useConferenceDay = confChk.checked;
      persist();
    });
    confTd.appendChild(confChk);

    const orderTd = createOrderButtons({
      onUp: () => moveTeam(idx, -1),
      onDown: () => moveTeam(idx, 1),
      upDisabled: idx === 0,
      downDisabled: idx === state.teams.length - 1,
      upTitle: "メイングループを上へ",
      downTitle: "メイングループを下へ",
    });

    const actTd = document.createElement("td");
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "btn-icon";
    delBtn.textContent = "削除";
    delBtn.addEventListener("click", () => {
      if (!confirm(`メイングループ「${t.name}」を削除しますか？`)) return;
      state.workers.forEach((w) => {
        if (w.teamId === t.id) {
          w.teamId = null;
          w.subGroupId = null;
          w.isGroupSupervisor = false;
        }
      });
      deleteSubGroupsForTeam(t.id);
      delete state.teamConstraints[t.id];
      state.teams.splice(idx, 1);
      renderTeams();
      renderWorkers();
      renderTeamAssignment();
      renderConstraints();
      persist();
    });
    actTd.appendChild(delBtn);

    tr.append(nameTd, membersTd, minTd, maxTd, supMinTd, supMaxTd, confTd, orderTd, actTd);
    tbody.appendChild(tr);
  });

  renderSubGroups();
  renderNewSubgroupTeamSelect();
  renderTeamAssignment();
  renderNewWorkerTeamSelect();
}

function renderSubGroups() {
  const tbody = $("#subgroups-tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  state.subGroups.forEach((sg, idx) => {
    const tr = document.createElement("tr");
    const sgc = ensureSubGroupConstraint(sg.id);
    const team = state.teams.find((t) => t.id === sg.teamId);

    const teamTd = document.createElement("td");
    teamTd.textContent = team?.name || "（削除済み）";

    const nameTd = document.createElement("td");
    const nameInput = document.createElement("input");
    nameInput.value = sg.name;
    nameInput.addEventListener("change", () => {
      sg.name = nameInput.value.trim() || sg.name;
      renderWorkers();
      persist();
    });
    nameTd.appendChild(nameInput);

    const membersTd = document.createElement("td");
    membersTd.className = "team-member-names";
    const members = getWorkersInSubGroup(sg.id);
    membersTd.textContent = members.length ? members.map((w) => w.name).join("、") : "—";

    const minTd = document.createElement("td");
    const minInput = document.createElement("input");
    minInput.type = "number";
    minInput.min = 0;
    minInput.value = sgc.min;
    minInput.addEventListener("change", () => {
      sgc.min = parseInt(minInput.value, 10) || 0;
      persist();
    });
    minTd.appendChild(minInput);

    const maxTd = document.createElement("td");
    const maxInput = document.createElement("input");
    maxInput.type = "number";
    maxInput.min = 0;
    maxInput.value = sgc.max;
    maxInput.addEventListener("change", () => {
      sgc.max = parseInt(maxInput.value, 10) || 0;
      persist();
    });
    maxTd.appendChild(maxInput);

    const confTd = document.createElement("td");
    const confChk = document.createElement("input");
    confChk.type = "checkbox";
    confChk.checked = Boolean(sgc.useConferenceDay);
    confChk.title = "月4回・隔週優先でカンファレンス日を設定（参加者が多い日を優先）";
    confChk.addEventListener("change", () => {
      sgc.useConferenceDay = confChk.checked;
      persist();
    });
    confTd.appendChild(confChk);

    const orderTd = createOrderButtons({
      onUp: () => moveSubGroup(sg.id, -1),
      onDown: () => moveSubGroup(sg.id, 1),
      upDisabled: !canMoveSubGroupUp(state.subGroups, sg.id),
      downDisabled: !canMoveSubGroupDown(state.subGroups, sg.id),
      upTitle: "サブグループを上へ（同一メイングループ内）",
      downTitle: "サブグループを下へ（同一メイングループ内）",
    });

    const actTd = document.createElement("td");
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "btn-icon";
    delBtn.textContent = "削除";
    delBtn.addEventListener("click", () => {
      if (!confirm(`サブグループ「${sg.name}」を削除しますか？`)) return;
      state.workers.forEach((w) => {
        if (w.subGroupId === sg.id) w.subGroupId = null;
      });
      delete state.subGroupConstraints[sg.id];
      state.subGroups.splice(idx, 1);
      renderTeams();
      renderWorkers();
      renderTeamAssignment();
      persist();
    });
    actTd.appendChild(delBtn);

    tr.append(teamTd, nameTd, membersTd, minTd, maxTd, confTd, orderTd, actTd);
    tbody.appendChild(tr);
  });
}

function renderNewSubgroupTeamSelect() {
  const sel = $("#new-subgroup-team");
  if (!sel) return;
  const prev = sel.value;
  sel.innerHTML = "";
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

function renderConstraints() {
  $("#supervisor-min").value = state.constraints.supervisorMin;
  $("#supervisor-max").value = state.constraints.supervisorMax;
  renderSupervisorStats();
  renderGroupSupervisorConstraintsSummary();
}

function renderGroupSupervisorConstraintsSummary() {
  const wrap = $("#group-supervisor-constraints-summary");
  if (!wrap) return;
  if (!state.teams.length) {
    wrap.innerHTML = `<h3 class="section-subheading">グループ責任者（1日あたり）</h3><p class="hint">メイングループが未作成です。「グループ」タブで設定してください。</p>`;
    return;
  }
  let html = `<h3 class="section-subheading">グループ責任者（1日あたり）</h3>
    <p class="hint">下限・上限の変更は「グループ」タブのメイングループ表で行います。</p>
    <div class="table-wrap"><table class="data-table"><thead><tr>
      <th>メイングループ</th><th>登録責任者</th><th>下限</th><th>上限</th>
    </tr></thead><tbody>`;
  for (const team of state.teams) {
    const tc = ensureTeamConstraint(team.id);
    const supers = state.workers.filter((w) => w.isGroupSupervisor && w.teamId === team.id);
    const names = supers.map((w) => w.name).join("、") || "—";
    const over = supers.length > (tc.supervisorMax ?? 99);
    html += `<tr${over ? ' class="supervisor-stats-warn"' : ""}>
      <td>${escapeHtml(team.name)}</td>
      <td title="${escapeHtml(names)}">${supers.length} 名（${escapeHtml(names)}）</td>
      <td>${tc.supervisorMin ?? 0}</td>
      <td>${tc.supervisorMax ?? 99}</td>
    </tr>`;
  }
  html += "</tbody></table></div>";
  wrap.innerHTML = html;
}

function renderSupervisorStats() {
  const supervisors = state.workers.filter((w) => w.isSupervisor);
  const groupSupervisors = state.workers.filter((w) => w.isGroupSupervisor);
  const count = supervisors.length;
  const names = supervisors.map((w) => w.name);
  const namesText = names.length ? names.join("、") : "—";
  const dailyMax = parseInt($("#supervisor-max")?.value, 10);
  const max = Number.isNaN(dailyMax) ? state.constraints.supervisorMax ?? 2 : dailyMax;

  const countEl = $("#supervisor-registered-count");
  if (countEl) countEl.textContent = String(count);

  const groupCountEl = $("#group-supervisor-registered-count");
  if (groupCountEl) groupCountEl.textContent = String(groupSupervisors.length);

  const countConstraintsEl = $("#supervisor-registered-count-constraints");
  if (countConstraintsEl) countConstraintsEl.textContent = String(count);

  const namesEl = $("#supervisor-registered-names");
  if (namesEl) {
    namesEl.textContent = namesText;
    namesEl.title = names.length ? names.join("\n") : "";
  }

  const maxDisplayEl = $("#supervisor-daily-max-display");
  if (maxDisplayEl) maxDisplayEl.textContent = String(max);

  const barWorkers = document.querySelector("#panel-workers .supervisor-stats-bar");
  const barConstraints = $("#supervisor-stats-constraints");
  const overMax = count > max;
  for (const bar of [barWorkers, barConstraints]) {
    if (!bar) continue;
    bar.classList.toggle("supervisor-stats-warn", overMax);
  }

  const hintEl = $("#supervisor-registered-hint");
  if (!hintEl) return;
  if (!count) {
    hintEl.textContent =
      "現在、全体責任者として登録されている勤務者はいません。「勤務者」タブのチェックで設定します。";
    hintEl.className = "hint warn-hint";
  } else if (overMax) {
    hintEl.textContent = `登録 ${count} 名ですが、1日あたりの上限は ${max} 名です。上限を ${count} 名以上にするか、全体責任者登録を減らしてください。`;
    hintEl.className = "hint warn-hint";
  } else {
    hintEl.textContent = `登録者: ${namesText}。上の「上限」は1日に出勤する全体責任者の人数です（登録人数とは別）。`;
    hintEl.className = "hint";
  }
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
      html += `<td>${formatPreferencePreview(prefs[name][d])}</td>`;
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
  const { year, month, assignments, stats, conferenceDays: rawConf = {}, teams = state.teams } = result;
  const workers = state.workers;
  const conferenceDays = normalizeConferenceDays(rawConf);
  const subGroups = state.subGroups ?? [];
  const confColorMap = buildConferenceColorMap(
    teams,
    subGroups,
    state.teamConstraints,
    state.subGroupConstraints
  );
  const days = getDaysInMonth(year, month);
  const useTypes = state.useShiftTypes;
  const root = $("#shift-print-root");
  if (root) root.classList.remove("hidden");

  const printTitle = $("#shift-print-title");
  const printMeta = $("#shift-print-meta");
  if (printTitle) printTitle.textContent = `シフト表 ${year}年${month}月`;
  if (printMeta) {
    printMeta.textContent = `勤務者 ${workers.length} 名`;
  }

  const summary = $("#shift-summary");
  summary.innerHTML = "";
  let summaryHtml = `
    <div class="card-stat">${year}年${month}月 · ${days}日間</div>
    <div class="card-stat">勤務者 ${workers.length} 名</div>
  `;

  teams.forEach((team) => {
    const tc = state.teamConstraints[team.id];
    if (!tc?.useConferenceDay) return;
    const entries = conferenceDays.teams[team.id] || [];
    const style = getConferenceGroupStyle(confColorMap, "team", team.id);
    const cardStyle = style
      ? ` style="border-left:4px solid ${style.color};background:${style.bg}"`
      : "";
    if (!entries.length) {
      summaryHtml += `<div class="card-stat conference-card"${cardStyle}>${escapeHtml(team.name)}: カンファレンス日なし</div>`;
      return;
    }
    const list = entries.map((e) => `${e.day}日(${e.count}名)`).join("、");
    summaryHtml += `<div class="card-stat conference-card"${cardStyle}><strong>${escapeHtml(team.name)}</strong> 会: ${escapeHtml(list)}</div>`;
  });

  subGroups.forEach((sg) => {
    const sgc = state.subGroupConstraints[sg.id];
    if (!sgc?.useConferenceDay) return;
    const team = teams.find((t) => t.id === sg.teamId);
    const label = team ? `${team.name} / ${sg.name}` : sg.name;
    const entries = conferenceDays.subGroups[sg.id] || [];
    const style = getConferenceGroupStyle(confColorMap, "subGroup", sg.id);
    const cardStyle = style
      ? ` style="border-left:4px solid ${style.color};background:${style.bg}"`
      : "";
    if (!entries.length) {
      summaryHtml += `<div class="card-stat conference-card"${cardStyle}>${escapeHtml(label)}: カンファレンス日なし</div>`;
      return;
    }
    const list = entries.map((e) => `${e.day}日(${e.count}名)`).join("、");
    summaryHtml += `<div class="card-stat conference-card"${cardStyle}><strong>${escapeHtml(label)}</strong> 会: ${escapeHtml(list)}</div>`;
  });

  summary.innerHTML = summaryHtml;

  const wrap = $("#shift-table-wrap");
  let html = "<table class='data-table'><thead><tr><th class='sticky-col'>勤務者</th>";
  for (let d = 1; d <= days; d++) {
    const dow = ["日", "月", "火", "水", "木", "金", "土"][new Date(year, month - 1, d).getDay()];
    const confTeams = getConferenceTeamsOnDay(conferenceDays, teams, d);
    const confSubGroups = getConferenceSubGroupsOnDay(conferenceDays, subGroups, d);
    const confItems = [
      ...confTeams.map((t) => ({ type: "team", id: t.id, label: t.name })),
      ...confSubGroups.map((sg) => {
        const team = teams.find((t) => t.id === sg.teamId);
        return {
          type: "subGroup",
          id: sg.id,
          label: team ? `${team.name}/${sg.name}` : sg.name,
        };
      }),
    ];
    const confHint = confItems.length
      ? ` title="${escapeHtml(confItems.map((c) => c.label).join("・"))} カンファレンス日"`
      : "";
    const confBadges = confItems
      .map((item) => {
        const style = getConferenceGroupStyle(confColorMap, item.type, item.id);
        if (!style) return "";
        return `<small class="conf-badge" style="background:${style.color}">会</small>`;
      })
      .join("");
    const thStyle =
      confItems.length === 1
        ? (() => {
            const style = getConferenceGroupStyle(confColorMap, confItems[0].type, confItems[0].id);
            return style ? ` style="background:${style.bg};border-bottom:3px solid ${style.color}"` : "";
          })()
        : "";
    html += `<th${thStyle}${confHint}>${d}<br><small>${dow}</small>${confBadges ? `<br>${confBadges}` : ""}</th>`;
  }
  html += "<th class='shift-off-col'>休み<br><small>日数</small></th></tr></thead><tbody>";

  getWorkerSections(workers, teams, subGroups).forEach((section) => {
    html += `<tr class="group-header-row"><td class="sticky-col" colspan="${days + 2}">${escapeHtml(section.label)}</td></tr>`;
    section.members.forEach((w) => {
      const team = teams.find((t) => t.id === w.teamId);
      const subGroup = subGroups.find((sg) => sg.id === w.subGroupId);
      const groupLabel = [team?.name, subGroup?.name].filter(Boolean).join(" / ");
      const actualOff = countWorkerOffDaysFromAssignments(assignments, w.id, days);
      const targetOff = w.monthlyOffDays ?? 0;
      const offMatch = Math.abs(actualOff - targetOff) < 0.001;
      html += `<tr><td class="sticky-col">${escapeHtml(w.name)}${w.isSupervisor ? " ★" : ""}${w.isGroupSupervisor ? " ◆" : ""}${groupLabel ? `<br><small>${escapeHtml(groupLabel)}</small>` : ""}</td>`;
      for (let d = 1; d <= days; d++) {
        const cell = assignments[w.id][d];
        let label = formatCellDisplay(cell, useTypes);
        const confGroup = getWorkerConferenceGroup(conferenceDays, w, d);
        const pref = state.preferences[w.name]?.[d];
        const isPreferredOff = cell?.type === "off" && (pref === "off" || pref === true);
        let cls = isPreferredOff ? "cell-preferred-off" : "cell-off";
        let cellStyle = "";
        if (cell?.type === "half-off") {
          cls = cell.half === "am" ? "cell-half-am" : "cell-half-pm";
        } else if (cell?.type === "work") {
          cls = w.isSupervisor ? "cell-supervisor" : "cell-work";
          if (confGroup) {
            const style = getConferenceGroupStyle(confColorMap, confGroup.type, confGroup.id);
            if (style) {
              cls += " cell-conference";
              cellStyle = ` style="background:${style.bg};color:${style.color};font-weight:700"`;
            }
          }
        }
        const confName =
          confGroup &&
          (confGroup.type === "team"
            ? teams.find((t) => t.id === confGroup.id)?.name
            : (() => {
                const sg = subGroups.find((s) => s.id === confGroup.id);
                const t = teams.find((x) => x.id === sg?.teamId);
                return sg ? (t ? `${t.name}/${sg.name}` : sg.name) : "";
              })());
        const titleParts = [label];
        if (isPreferredOff) titleParts.push("希望休");
        if (confName) titleParts.push(`${confName} カンファレンス`);
        const title = titleParts.join(" · ");
        html += `<td class="${cls}"${cellStyle} title="${escapeHtml(title)}">${escapeHtml(label)}</td>`;
      }
      html += `<td class="shift-off-col${offMatch ? "" : " shift-off-mismatch"}" title="目標 ${targetOff} 日">${formatOffDaysDisplay(actualOff, targetOff)}</td></tr>`;
    });
  });

  html += "</tbody><tfoot><tr class='shift-total-row'>";
  html += "<td class='sticky-col'>出勤合計</td>";
  for (let d = 1; d <= days; d++) {
    const total =
      stats.daily[d - 1]?.total ??
      workers.filter((w) => assignments[w.id][d]?.type !== "off").length;
    html += `<td class="shift-total-cell">${total}</td>`;
  }
  html += "<td class='shift-off-col'></td></tr></tfoot></table>";

  const legendItems = [];
  teams.forEach((team) => {
    if (!state.teamConstraints[team.id]?.useConferenceDay) return;
    const style = getConferenceGroupStyle(confColorMap, "team", team.id);
    if (!style) return;
    legendItems.push(
      `<span class="conf-legend-item"><span class="conf-legend-swatch" style="background:${style.bg};border-color:${style.color}"></span>${escapeHtml(team.name)}</span>`
    );
  });
  subGroups.forEach((sg) => {
    if (!state.subGroupConstraints[sg.id]?.useConferenceDay) return;
    const team = teams.find((t) => t.id === sg.teamId);
    const label = team ? `${team.name}/${sg.name}` : sg.name;
    const style = getConferenceGroupStyle(confColorMap, "subGroup", sg.id);
    if (!style) return;
    legendItems.push(
      `<span class="conf-legend-item"><span class="conf-legend-swatch" style="background:${style.bg};border-color:${style.color}"></span>${escapeHtml(label)}</span>`
    );
  });
  if (legendItems.length) {
    html += `<p class="shift-conf-legend">${legendItems.join("")}</p>`;
  }

  wrap.innerHTML = html;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatOffDaysDisplay(actual, target) {
  const actualText = Number.isInteger(actual) ? String(actual) : actual.toFixed(1);
  const targetText = Number.isInteger(target) ? String(target) : target.toFixed(1);
  return `${actualText}/${targetText}`;
}

init();
