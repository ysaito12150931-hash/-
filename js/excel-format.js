/** @typedef {import('xlsx').WorkSheet} WorkSheet */

import { getWorkerSections } from "./worker-groups.js";

const DOW_LABELS = ["日", "月", "火", "水", "木", "金", "土"];

/** 白背景に 50% 重ねた相当色（Excel は透過非対応のため近似） */
export const EXCEL_COLORS = {
  white: "FFFFFF",
  headerBg: "2F5496",
  weekendBg: "FF6B6B",
  weekendBgSoft: "FFE5E5",
  workerBg: "FFFFFF",
  workerNameBg: "F2F2F2",
  /** 50% 灰色（#808080 @ 50% on #FFF） */
  spacerGray50: "BFBFBF",
  /** 50% 黑色（#000000 @ 50% on #FFF） */
  spacerBlack50: "808080",
};

const baseBorder = {
  top: { style: "thin", color: { rgb: "B4B4B4" } },
  bottom: { style: "thin", color: { rgb: "B4B4B4" } },
  left: { style: "thin", color: { rgb: "B4B4B4" } },
  right: { style: "thin", color: { rgb: "B4B4B4" } },
};

function fill(rgb) {
  return { patternType: "solid", fgColor: { rgb } };
}

function font(rgb, bold = false) {
  return { name: "Yu Gothic UI", sz: 10, bold, color: { rgb } };
}

export const STYLES = {
  headerName: {
    font: font(EXCEL_COLORS.white, true),
    fill: fill(EXCEL_COLORS.headerBg),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  headerDay: {
    font: font(EXCEL_COLORS.white, true),
    fill: fill(EXCEL_COLORS.headerBg),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  headerDayWeekend: {
    font: font(EXCEL_COLORS.white, true),
    fill: fill(EXCEL_COLORS.weekendBg),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  headerDow: {
    font: font("333333", true),
    fill: fill("E8EEF4"),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  headerDowWeekend: {
    font: font(EXCEL_COLORS.white, true),
    fill: fill(EXCEL_COLORS.weekendBg),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  workerName: {
    font: font("111111", true),
    fill: fill(EXCEL_COLORS.workerNameBg),
    alignment: { horizontal: "left", vertical: "center" },
    border: baseBorder,
  },
  workerCell: {
    font: font("111111", false),
    fill: fill(EXCEL_COLORS.workerBg),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  workerCellWeekend: {
    font: font("111111", false),
    fill: fill(EXCEL_COLORS.weekendBgSoft),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  teamLabel: {
    font: font("1a2744", true),
    fill: fill("D9E2F3"),
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: baseBorder,
  },
  /** 同一チーム内の勤務者と勤務者の間 */
  spacerWithinTeam: {
    font: font("666666", false),
    fill: fill(EXCEL_COLORS.spacerGray50),
    border: baseBorder,
  },
  /** チームとチームの境界 */
  spacerTeamBoundary: {
    font: font("FFFFFF", false),
    fill: fill(EXCEL_COLORS.spacerBlack50),
    border: baseBorder,
  },
};

/**
 * @param {WorkSheet} ws
 * @param {number} r
 * @param {number} c
 * @param {string|number} value
 * @param {object} style
 */
export function setStyledCell(ws, r, c, value, style) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = { s: style };
  if (value === "" || value == null) {
    cell.v = "";
    cell.t = "s";
  } else if (typeof value === "number") {
    cell.v = value;
    cell.t = "n";
  } else {
    cell.v = String(value);
    cell.t = "s";
  }
  ws[addr] = cell;
}

export function getWeekendDaySet(year, month, daysInMonth) {
  const set = new Set();
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month - 1, d).getDay();
    if (dow === 0 || dow === 6) set.add(d);
  }
  return set;
}

export function getWeekdayLabel(year, month, day) {
  return DOW_LABELS[new Date(year, month - 1, day).getDay()];
}

const COL_TEAM = 0;
const COL_NAME = 1;
export const OFF_DAYS_HEADER = "月間休み日数";

function dayColumn(day) {
  return COL_NAME + day;
}

function offDaysColumn(days) {
  return dayColumn(days) + 1;
}

function lastColumnIndex(days) {
  return offDaysColumn(days);
}

function fillSpacerRow(ws, r, days, style) {
  for (let c = COL_TEAM; c <= lastColumnIndex(days); c++) {
    setStyledCell(ws, r, c, "", style);
  }
}

function writeWorkerRow(ws, r, days, worker, weekends) {
  setStyledCell(ws, r, COL_TEAM, "", STYLES.workerName);
  setStyledCell(ws, r, COL_NAME, worker.name, STYLES.workerName);
  for (let d = 1; d <= days; d++) {
    const style = weekends.has(d) ? STYLES.workerCellWeekend : STYLES.workerCell;
    const value = worker.cells?.[d - 1] ?? "";
    setStyledCell(ws, r, dayColumn(d), value, style);
  }
  const offDays = worker.monthlyOffDays;
  setStyledCell(
    ws,
    r,
    offDaysColumn(days),
    offDays == null || offDays === "" ? "" : offDays,
    STYLES.workerCell
  );
}

function getGroupLabel(section) {
  if (section.team?.name && section.subGroup?.name) {
    return `${section.team.name} / ${section.subGroup.name}`;
  }
  return section.label;
}

function applyTeamColumnMerge(ws, rStart, rEnd, label) {
  setStyledCell(ws, rStart, COL_TEAM, label, STYLES.teamLabel);
  if (rEnd > rStart) {
    if (!ws["!merges"]) ws["!merges"] = [];
    ws["!merges"].push({ s: { r: rStart, c: COL_TEAM }, e: { r: rEnd, c: COL_TEAM } });
  }
}

/**
 * 2行ヘッダー（日付・曜日）＋チームごとにまとめた勤務者行
 * @param {WorkSheet} ws
 * @param {number} year
 * @param {number} month
 * @param {number} days
 * @param {{ id?: string, name: string, teamId?: string|null, subGroupId?: string|null, isSupervisor?: boolean, cells?: string[] }[]} workers
 * @param {{ id: string, name: string }[]} teams
 * @param {{ id: string, name: string, teamId: string }[]} [subGroups]
 */
export function fillCalendarTemplateSheet(ws, year, month, days, workers, teams = [], subGroups = []) {
  const weekends = getWeekendDaySet(year, month, days);
  const sections = getWorkerSections(workers, teams, subGroups);

  setStyledCell(ws, 0, COL_TEAM, "グループ", STYLES.headerName);
  setStyledCell(ws, 0, COL_NAME, "勤務者", STYLES.headerName);
  for (let d = 1; d <= days; d++) {
    const style = weekends.has(d) ? STYLES.headerDayWeekend : STYLES.headerDay;
    setStyledCell(ws, 0, dayColumn(d), d, style);
  }
  setStyledCell(ws, 0, offDaysColumn(days), OFF_DAYS_HEADER, STYLES.headerName);

  setStyledCell(ws, 1, COL_TEAM, "", STYLES.headerName);
  setStyledCell(ws, 1, COL_NAME, "", STYLES.headerName);
  for (let d = 1; d <= days; d++) {
    const style = weekends.has(d) ? STYLES.headerDowWeekend : STYLES.headerDow;
    setStyledCell(ws, 1, dayColumn(d), getWeekdayLabel(year, month, d), style);
  }
  setStyledCell(ws, 1, offDaysColumn(days), "日数", STYLES.headerDow);

  let r = 2;
  ws["!merges"] = [];

  sections.forEach((section, sectionIndex) => {
    const groupStartRow = r;

    section.members.forEach((w, memberIndex) => {
      writeWorkerRow(ws, r, days, w, weekends);
      r++;

      if (memberIndex < section.members.length - 1) {
        fillSpacerRow(ws, r, days, STYLES.spacerWithinTeam);
        r++;
      }
    });

    applyTeamColumnMerge(ws, groupStartRow, r - 1, getGroupLabel(section));

    if (sectionIndex < sections.length - 1) {
      fillSpacerRow(ws, r, days, STYLES.spacerTeamBoundary);
      r++;
    }
  });

  ws["!ref"] = XLSX.utils.encode_range({
    s: { c: 0, r: 0 },
    e: { c: lastColumnIndex(days), r: Math.max(r - 1, 1) },
  });
  ws["!cols"] = [{ wch: 12 }, { wch: 14 }, ...Array.from({ length: days }, () => ({ wch: 6 })), { wch: 10 }];
  ws["!rows"] = [{ hpt: 22 }, { hpt: 20 }];
  ws["!views"] = [
    {
      state: "frozen",
      xSplit: 2,
      ySplit: 2,
      topLeftCell: "C3",
      activeCell: "C3",
    },
  ];
}

export function detectCalendarDataStartRow(rows) {
  if (rows.length < 2) return 1;
  const row1 = rows[1] ?? [];
  for (let c = 1; c < row1.length; c++) {
    if (/^[月火水木金土日]$/u.test(String(row1[c] ?? "").trim())) return 2;
  }
  return 1;
}

/** @param {unknown[]} header */
export function detectOffDaysColumn(header) {
  for (let c = 0; c < header.length; c++) {
    const h = String(header[c] ?? "").trim();
    if (h === OFF_DAYS_HEADER || h.includes("月間休み")) return c;
  }
  return null;
}

/** @param {unknown[]} header */
export function detectNameColumn(header) {
  for (let c = 0; c < header.length; c++) {
    const h = String(header[c] ?? "").trim();
    if (h === "勤務者" || h.includes("勤務者")) return c;
  }
  const a = String(header[0] ?? "").trim();
  if (a === "チーム" || a.includes("チーム") || a === "グループ" || a.includes("グループ")) return 1;
  return 0;
}
