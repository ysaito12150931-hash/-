import { getDaysInMonth } from "./store.js";
import {
  fillCalendarTemplateSheet,
  detectCalendarDataStartRow,
  detectNameColumn,
  detectOffDaysColumn,
} from "./excel-format.js";

const OFF_MARKERS = new Set([
  "休",
  "休み",
  "×",
  "x",
  "X",
  "off",
  "OFF",
  "0",
  "希望休",
  "公休",
]);

const WORK_MARKERS = new Set([
  "出勤",
  "出",
  "勤",
  "勤務",
  "work",
  "WORK",
  "1",
  "○",
  "◯",
  "丸",
  "希望出勤",
  "◎",
]);

const AM_OFF_MARKERS = new Set([
  "午前休",
  "前休",
  "午前のみ休み",
  "午前だけ休み",
  "AM休",
  "am休",
  "am-off",
  "AM-OFF",
  "午前休み",
]);

const PM_OFF_MARKERS = new Set([
  "午後休",
  "後休",
  "午後のみ休み",
  "午後だけ休み",
  "PM休",
  "pm休",
  "pm-off",
  "PM-OFF",
  "午後休み",
]);

export function isOffMarker(value) {
  if (value == null || value === "") return false;
  const s = String(value).trim();
  return OFF_MARKERS.has(s);
}

export function isWorkMarker(value) {
  if (value == null || value === "") return false;
  const s = String(value).trim();
  return WORK_MARKERS.has(s);
}

export function isAmOffMarker(value) {
  if (value == null || value === "") return false;
  return AM_OFF_MARKERS.has(String(value).trim());
}

export function isPmOffMarker(value) {
  if (value == null || value === "") return false;
  return PM_OFF_MARKERS.has(String(value).trim());
}

/** @returns {"off"|"work"|"am-off"|"pm-off"|"conflict"|null} */
export function parsePreferenceCell(value) {
  const s = value == null ? "" : String(value).trim();
  if (!s) return null;

  const am = isAmOffMarker(s);
  const pm = isPmOffMarker(s);
  if (am && pm) return "conflict";
  if (am) return "am-off";
  if (pm) return "pm-off";

  const off = isOffMarker(value);
  const work = isWorkMarker(value);
  if (off && work) return "conflict";
  if (off) return "off";
  if (work) return "work";
  return null;
}

/** 旧形式（true = 休み）との互換 */
export function normalizePreferences(preferences) {
  const out = {};
  for (const [name, days] of Object.entries(preferences || {})) {
    out[name] = {};
    for (const [day, value] of Object.entries(days || {})) {
      if (value === true || value === "off") out[name][day] = "off";
      else if (value === "work") out[name][day] = "work";
      else if (value === "am-off") out[name][day] = "am-off";
      else if (value === "pm-off") out[name][day] = "pm-off";
    }
  }
  return out;
}

/**
 * Parse matrix Excel: preferences[name][day] = "off" | "work" | "am-off" | "pm-off"
 */
export function parsePreferenceSheet(workbook, workerNames, year, month) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
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
  const nameSet = new Set(workerNames);

  for (let r = dataStartRow; r < rows.length; r++) {
    const row = rows[r];
    const name = String(row[nameCol] ?? "").trim();
    if (!name) continue;
    if (name === "チーム" || name.includes("未所属") || name.includes("責任者")) continue;
    if (!nameSet.has(name)) {
      warnings.push(`未登録の勤務者: ${name}`);
      continue;
    }
    if (!preferences[name]) preferences[name] = {};

    if (offDaysCol != null) {
      const offDays = parseMonthlyOffDaysCell(row[offDaysCol]);
      if (offDays != null) monthlyOffDaysByName[name] = offDays;
    }

    for (const { col, day } of dayCols) {
      const kind = parsePreferenceCell(row[col]);
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
  const { year, month, assignments, workers } = result;
  const days = getDaysInMonth(year, month);
  const ws = {};

  const dataWorkers = workers.map((w) => {
    const byDay = assignments[w.id] || {};
    const cells = [];
    for (let d = 1; d <= days; d++) {
      cells.push(formatCellExport(byDay[d], state.useShiftTypes));
    }
    return {
      id: w.id,
      name: w.name,
      teamId: w.teamId,
      subGroupId: w.subGroupId,
      isSupervisor: w.isSupervisor,
      monthlyOffDays: w.monthlyOffDays,
      cells,
    };
  });

  fillCalendarTemplateSheet(ws, year, month, days, dataWorkers, state.teams ?? [], state.subGroups ?? []);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `シフト_${year}${month}`);
  return wb;
}

function formatCellExport(cell, useShiftTypes) {
  if (!cell || cell.type === "off") return "休";
  if (cell.type === "half-off") return cell.half === "am" ? "午前休" : "午後休";
  if (useShiftTypes && cell.shiftType) return cell.shiftType;
  return "１";
}

export function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename, { cellStyles: true });
}

export async function readWorkbookFromFile(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellDates: true });
}

export function formatPreferencePreview(kind) {
  if (kind === "off" || kind === true) return "休";
  if (kind === "work") return "出";
  if (kind === "am-off") return "前休";
  if (kind === "pm-off") return "後休";
  return "";
}

/** 希望の休み日数（半休は0.5日） */
export function countPreferenceOffDays(pref, days) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    const v = pref?.[d];
    if (v === "off" || v === true) n += 1;
    else if (v === "am-off" || v === "pm-off") n += 0.5;
  }
  return n;
}
