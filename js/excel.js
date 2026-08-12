import { getDaysInMonth } from "./store.js";
import {
  fillCalendarTemplateSheet,
  detectCalendarDataStartRow,
  detectNameColumn,
  detectOffDaysColumn,
  isGroupSubHeaderRow,
} from "./excel-format.js";
import {
  normalizeConferenceDays,
  buildConferenceColorMap,
  getConferenceGroupStyle,
  getWorkerConferenceGroup,
  getConferenceTeamsOnDay,
  getConferenceSubGroupsOnDay,
} from "./scheduler.js";
import { getSectionHeadcountBounds, getTeamHeadcountBounds, isHeadcountOutOfRange } from "./worker-groups.js";
import {
  parsePreferenceCell,
  isPreferredOffValue,
  getPreferenceValue,
  getWorkerPrefMap,
} from "./preference-markers.js";

export {
  parsePreferenceCell,
  isPreferredOffValue,
  getPreferenceValue,
  getWorkerPrefMap,
} from "./preference-markers.js";

/** 旧形式（true = 休み）および 〇 / ○ などの丸印との互換 */
export function normalizePreferences(preferences) {
  const out = {};
  for (const [name, days] of Object.entries(preferences || {})) {
    out[name] = {};
    for (const [day, value] of Object.entries(days || {})) {
      const kind = parsePreferenceCell(value);
      if (value === true || value === "off" || kind === "off") out[name][day] = "off";
      else if (value === "work" || kind === "work") out[name][day] = "work";
      else if (value === "am-off" || kind === "am-off") out[name][day] = "am-off";
      else if (value === "pm-off" || kind === "pm-off") out[name][day] = "pm-off";
    }
  }
  return out;
}

function resolveWorkerName(excelName, workerNames) {
  const raw = String(excelName ?? "").trim();
  if (!raw) return null;
  if (workerNames.includes(raw)) return raw;
  const norm = (s) => String(s).trim().replace(/[★◆]/g, "").replace(/\s+/g, "");
  const n = norm(raw);
  return workerNames.find((w) => norm(w) === n) || null;
}

function readSheetCell(sheet, r, c, jsonFallback) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = sheet[addr];
  if (cell) {
    if (cell.w != null && String(cell.w).trim() !== "") return cell.w;
    if (cell.v != null && cell.v !== "") return cell.v;
  }
  return jsonFallback ?? "";
}

/**
 * Parse matrix Excel: preferences[name][day] = "off" | "work" | "am-off" | "pm-off"
 */
export function parsePreferenceSheet(workbook, workerNames, year, month) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  if (!rows.length) return { preferences: {}, monthlyOffDaysByName: {}, warnings: ["シートが空です"] };

  const header = rows[0];
  const dataStartRow = detectCalendarDataStartRow(rows);
  const nameCol = detectNameColumn(header);
  const offDaysCol = detectOffDaysColumn(header);
  const dayColEnd = offDaysCol != null ? offDaysCol : header.length;
  const dayCols = [];
  const daysInMonth = getDaysInMonth(year, month);

  for (let c = nameCol + 1; c < dayColEnd; c++) {
    const day = parseDayHeader(header[c], c, year, month);
    if (day >= 1 && day <= daysInMonth) {
      dayCols.push({ col: c, day });
    }
  }

  const warnings = [];
  if (!dayCols.length) {
    warnings.push(
      "日付列を認識できませんでした。1行目の勤務者列の右側に日（1〜31）を入力してください。"
    );
  }

  const preferences = {};
  const monthlyOffDaysByName = {};

  for (let r = dataStartRow; r < rows.length; r++) {
    const row = rows[r];
    const excelName = String(row[nameCol] ?? "").trim();
    if (!excelName) continue;
    if (isGroupSubHeaderRow(row, nameCol)) continue;
    if (excelName === "チーム" || excelName === "備考" || excelName.startsWith("出勤合計") || excelName.includes("未所属") || excelName.includes("責任者")) continue;
    const name = resolveWorkerName(excelName, workerNames);
    if (!name) {
      warnings.push(`未登録の勤務者: ${excelName}`);
      continue;
    }
    if (!preferences[name]) preferences[name] = {};

    if (offDaysCol != null) {
      const offDays = parseMonthlyOffDaysCell(row[offDaysCol]);
      if (offDays != null) monthlyOffDaysByName[name] = offDays;
    }

    for (const { col, day } of dayCols) {
      const kind = parsePreferenceCell(readSheetCell(sheet, r, col, row[col]));
      if (kind === "conflict") {
        warnings.push(`${name} ${day}日: 休みと出勤の両方の指定があります（スキップ）`);
        continue;
      }
      if (kind) preferences[name][day] = kind;
    }
  }

  return { preferences, monthlyOffDaysByName, warnings };
}

function parseMonthlyOffDaysCell(value) {
  if (value == null || value === "") return null;
  const n = parseFloat(String(value).trim());
  return Number.isNaN(n) ? null : Math.max(0, n);
}

function parseDayHeader(cell, colIndex, year, month) {
  if (cell == null || cell === "") return colIndex;
  if (typeof cell === "number" && cell >= 1 && cell <= 31) return Math.floor(cell);

  const s = String(cell).trim();
  const num = parseInt(s, 10);
  if (!Number.isNaN(num) && num >= 1 && num <= 31) return num;

  const m = s.match(/(\d{1,2})\s*日?/);
  if (m) return parseInt(m[1], 10);

  if (cell instanceof Date) {
    return cell.getDate();
  }

  const excelDate = XLSX.SSF?.parse_date_code?.(cell);
  if (excelDate?.d) return excelDate.d;

  return colIndex;
}

export function buildTemplateWorkbook(state) {
  const { year, month, workers } = state;
  const days = getDaysInMonth(year, month);
  const ws = {};
  fillCalendarTemplateSheet(ws, year, month, days, workers, state.teams ?? [], state.subGroups ?? []);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "勤務希望");
  return wb;
}

export function exportShiftWorkbook(result, state) {
  const { year, month, assignments, workers, conferenceDays: rawConf = {} } = result;
  const days = getDaysInMonth(year, month);
  const ws = {};
  const conferenceDays = normalizeConferenceDays(rawConf);
  const teams = state.teams ?? [];
  const subGroups = state.subGroups ?? [];
  const preferences = normalizePreferences(state.preferences);
  const confColorMap = buildConferenceColorMap(
    teams,
    subGroups,
    state.teamConstraints ?? {},
    state.subGroupConstraints ?? {}
  );

  const dataWorkers = workers.map((w) => {
    const byDay = assignments[w.id] || {};
    const pref = getWorkerPrefMap(preferences, w.name) || {};
    const cells = [];
    for (let d = 1; d <= days; d++) {
      const cell = byDay[d];
      const isConference =
        cell?.type !== "off" && Boolean(getWorkerConferenceGroup(conferenceDays, w, d));
      const preferredOff =
        cell?.type === "off" &&
        (cell.preferredOff || isPreferredOffValue(getPreferenceValue(pref, d)));
      const off = !cell || cell.type === "off" || cell.type === "half-off";
      const attending = Boolean(cell && cell.type !== "off");
      cells.push({
        text: formatCellExport(cell, state.useShiftTypes, isConference),
        conference: isConference,
        preferredOff,
        off,
        attending,
      });
    }
    return {
      id: w.id,
      name: w.name,
      teamId: w.teamId,
      subGroupId: w.subGroupId,
      isSupervisor: w.isSupervisor,
      isGroupSupervisor: w.isGroupSupervisor,
      monthlyOffDays: w.monthlyOffDays,
      cells,
    };
  });

  fillCalendarTemplateSheet(ws, year, month, days, dataWorkers, teams, subGroups, {
    variant: "shift",
    getConferenceHeaderLabel(day) {
      const hasTeam = getConferenceTeamsOnDay(conferenceDays, teams, day).length > 0;
      const hasSubGroup = getConferenceSubGroupsOnDay(conferenceDays, subGroups, day).length > 0;
      return hasTeam || hasSubGroup;
    },
    getConferenceStyle(worker, day) {
      const cell = assignments[worker.id]?.[day];
      if (!cell || cell.type === "off") return null;
      const group = getWorkerConferenceGroup(conferenceDays, worker, day);
      if (!group) return null;
      return getConferenceGroupStyle(confColorMap, group.type, group.id);
    },
    getGroupTotalOutOfRange(section, day, total) {
      const bounds = getSectionHeadcountBounds(
        section,
        state.teamConstraints ?? {},
        state.subGroupConstraints ?? {},
        subGroups
      );
      return isHeadcountOutOfRange(total, bounds);
    },
    getTeamTotalOutOfRange(team, day, total) {
      const bounds = getTeamHeadcountBounds(team, state.teamConstraints ?? {});
      return isHeadcountOutOfRange(total, bounds);
    },
    getSupervisorCount(day) {
      return (
        result.stats?.daily?.[day - 1]?.supervisors ??
        workers.filter((w) => w.isSupervisor && assignments[w.id]?.[day]?.type !== "off").length
      );
    },
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `シフト_${year}${month}`);
  return wb;
}

function formatCellExport(cell, useShiftTypes, isConference = false) {
  if (!cell || cell.type === "off") return "休";
  if (cell.type === "half-off") return cell.half === "am" ? "午前休" : "午後休";
  if (useShiftTypes && cell.shiftType) return isConference ? `${cell.shiftType}会` : cell.shiftType;
  return isConference ? "１会" : "１";
}

export function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename, { cellStyles: true });
}

export async function readWorkbookFromFile(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellDates: true });
}

export function formatPreferencePreview(kind) {
  if (isPreferredOffValue(kind) || kind === "off" || kind === true) return "休";
  if (kind === "work") return "出";
  if (kind === "am-off") return "前休";
  if (kind === "pm-off") return "後休";
  return "";
}

/** 希望の休み日数（半休は0.5日） */
export function countPreferenceOffDays(pref, days) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    const v = getPreferenceValue(pref, d);
    if (isPreferredOffValue(v)) n += 1;
    else if (v === "am-off" || v === "pm-off") n += 0.5;
  }
  return n;
}
