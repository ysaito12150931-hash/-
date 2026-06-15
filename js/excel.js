import { getDaysInMonth } from "./store.js";
import {
  fillCalendarTemplateSheet,
  detectCalendarDataStartRow,
  detectNameColumn,
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

export function isOffMarker(value) {
  if (value == null || value === "") return false;
  const s = String(value).trim();
  return OFF_MARKERS.has(s);
}

/**
 * Parse matrix Excel: header with day numbers (row 0), optional weekday row (row 1)
 * Returns { workerName: { dayNumber: true } }
 */
export function parsePreferenceSheet(workbook, workerNames, year, month) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows.length) return { preferences: {}, warnings: ["シートが空です"] };

  const header = rows[0];
  const dataStartRow = detectCalendarDataStartRow(rows);
  const nameCol = detectNameColumn(header);
  const dayCols = [];
  const daysInMonth = getDaysInMonth(year, month);

  for (let c = nameCol + 1; c < header.length; c++) {
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
    for (const { col, day } of dayCols) {
      if (isOffMarker(row[col])) {
        preferences[name][day] = true;
      }
    }
  }

  return { preferences, warnings };
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
  fillCalendarTemplateSheet(ws, year, month, days, workers, state.teams ?? []);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "休み希望");
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
    return { id: w.id, name: w.name, teamId: w.teamId, isSupervisor: w.isSupervisor, cells };
  });

  fillCalendarTemplateSheet(ws, year, month, days, dataWorkers, state.teams ?? []);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `シフト_${year}${month}`);
  return wb;
}

function formatCellExport(cell, useShiftTypes) {
  if (!cell || cell.type === "off") return "休";
  if (useShiftTypes && cell.shiftType) return cell.shiftType;
  return "出勤";
}

export function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename, { cellStyles: true });
}

export async function readWorkbookFromFile(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellDates: true });
}
