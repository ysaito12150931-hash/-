/** @typedef {import('xlsx').WorkSheet} WorkSheet */

import { getWorkerSections, groupWorkerSectionsByTeam } from "./worker-groups.js";

const DOW_LABELS = ["日", "月", "火", "水", "木", "金", "土"];

/** 白背景に 50% 重ねた相当色（Excel は透過非対応のため近似） */
export const EXCEL_COLORS = {
  white: "FFFFFF",
  headerBg: "2F5496",
  weekendBg: "FF6B6B",
  weekendBgSoft: "FFE5E5",
  preferredOffBg: "86EFAC",
  preferredOffFont: "14532D",
  autoOffBg: "FEF3C7",
  autoOffFont: "92400E",
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
  headerDayWeekendShift: {
    font: font(EXCEL_COLORS.white, true),
    fill: fill(EXCEL_COLORS.weekendBg),
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: baseBorder,
  },
  headerDowWeekendShift: {
    font: font(EXCEL_COLORS.white, true),
    fill: fill(EXCEL_COLORS.weekendBg),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  workerCellWeekendShift: {
    font: font("111111", false),
    fill: fill(EXCEL_COLORS.weekendBgSoft),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  workerCellPreferredOff: {
    font: font(EXCEL_COLORS.preferredOffFont, true),
    fill: fill(EXCEL_COLORS.preferredOffBg),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  workerCellAutoOff: {
    font: font(EXCEL_COLORS.autoOffFont, false),
    fill: fill(EXCEL_COLORS.autoOffBg),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  workerCellConference: {
    font: font("047857", true),
    fill: fill("D1FAE5"),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  workerCellConferenceWeekendShift: {
    font: font("047857", true),
    fill: fill("BBF7D0"),
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
  footerLabel: {
    font: font("7F1D1D", true),
    fill: fill("FFEDD5"),
    alignment: { horizontal: "left", vertical: "center" },
    border: baseBorder,
  },
  footerCell: {
    font: font("111111", false),
    fill: fill("FFF7ED"),
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: baseBorder,
  },
  footerCellAbsent: {
    font: { name: "Yu Gothic UI", sz: 8, bold: true, color: { rgb: "991B1B" } },
    fill: fill("FECACA"),
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: baseBorder,
  },
  footerCellDeviation: {
    font: { name: "Yu Gothic UI", sz: 9, bold: true, color: { rgb: "6B21A8" } },
    fill: fill("EDE9FE"),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  },
  teamTotalLabel: {
    font: font("FFFFFF", true),
    fill: fill("64748B"),
    alignment: { horizontal: "left", vertical: "center" },
    border: baseBorder,
  },
  teamTotalCell: {
    font: font("0F172A", true),
    fill: fill("94A3B8"),
    alignment: { horizontal: "center", vertical: "center" },
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

function makeStyle(fontRgb, fillRgb, bold = false) {
  return {
    font: font(fontRgb, bold),
    fill: fill(fillRgb),
    alignment: { horizontal: "center", vertical: "center" },
    border: baseBorder,
  };
}

function resolveWeekendStyles(variant) {
  if (variant === "shift") {
    return {
      headerDay: STYLES.headerDayWeekendShift,
      headerDow: STYLES.headerDowWeekendShift,
      workerCell: STYLES.workerCellWeekendShift,
    };
  }
  return {
    headerDay: STYLES.headerDayWeekend,
    headerDow: STYLES.headerDowWeekend,
    workerCell: STYLES.workerCellWeekend,
  };
}

function normalizeCellValue(cell) {
  if (cell == null || cell === "")
    return { text: "", conference: false, preferredOff: false, off: false, attending: false };
  if (typeof cell === "object" && "text" in cell) {
    return {
      text: cell.text,
      conference: Boolean(cell.conference),
      preferredOff: Boolean(cell.preferredOff),
      off: Boolean(cell.off),
      attending: Boolean(cell.attending),
    };
  }
  return { text: cell, conference: false, preferredOff: false, off: false, attending: false };
}

function countAttendingFromWorkerCells(members, day) {
  return members.filter((w) => normalizeCellValue(w.cells?.[day - 1] ?? "").attending).length;
}

function writeWorkerRow(ws, r, days, worker, weekends, variant, getConferenceStyle) {
  setStyledCell(ws, r, COL_TEAM, "", STYLES.workerName);
  setStyledCell(ws, r, COL_NAME, worker.name, STYLES.workerName);
  const weekendStyles = resolveWeekendStyles(variant);
  for (let d = 1; d <= days; d++) {
    const cell = normalizeCellValue(worker.cells?.[d - 1] ?? "");
    let style = weekends.has(d) ? weekendStyles.workerCell : STYLES.workerCell;
    const confStyle = getConferenceStyle?.(worker, d);
    if (confStyle) {
      style = makeStyle(confStyle.color.replace("#", ""), confStyle.bg.replace("#", ""), true);
    } else if (cell.preferredOff && variant === "shift") {
      style = STYLES.workerCellPreferredOff;
    } else if (cell.off && variant === "shift") {
      style = STYLES.workerCellAutoOff;
    } else if (cell.conference) {
      style = weekends.has(d) && variant === "shift"
        ? STYLES.workerCellConferenceWeekendShift
        : STYLES.workerCellConference;
    }
    setStyledCell(ws, r, dayColumn(d), cell.text, style);
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

function writeGroupSubHeaderRow(ws, r, days, weekends, variant) {
  const weekendStyles = resolveWeekendStyles(variant);
  setStyledCell(ws, r, COL_NAME, "", STYLES.headerDow);
  for (let d = 1; d <= days; d++) {
    const style = weekends.has(d) ? weekendStyles.headerDay : STYLES.headerDay;
    setStyledCell(ws, r, dayColumn(d), d, style);
  }
  setStyledCell(ws, r, offDaysColumn(days), "日数", STYLES.headerDow);
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
 * @param {{ variant?: 'template'|'shift', getConferenceHeaderLabel?: (day:number)=>string, getConferenceStyle?: (worker:object, day:number)=>object|null, getFooterLabel?: (day:number)=>string, getSupervisorCount?: (day:number)=>number, getGroupTotalOutOfRange?: (section:object, day:number, total:number)=>boolean, getTeamTotalOutOfRange?: (team:object, day:number, total:number)=>boolean }} [options]
 */
export function fillCalendarTemplateSheet(ws, year, month, days, workers, teams = [], subGroups = [], options = {}) {
  const variant = options.variant ?? "template";
  const getConferenceHeaderLabel = options.getConferenceHeaderLabel;
  const getConferenceStyle = options.getConferenceStyle;
  const getFooterLabel = options.getFooterLabel;
  const getSupervisorCount = options.getSupervisorCount;
  const getGroupTotalOutOfRange = options.getGroupTotalOutOfRange;
  const getTeamTotalOutOfRange = options.getTeamTotalOutOfRange;
  const weekends = getWeekendDaySet(year, month, days);
  const weekendStyles = resolveWeekendStyles(variant);
  const sections = getWorkerSections(workers, teams, subGroups);
  const sectionGroups = groupWorkerSectionsByTeam(sections);

  setStyledCell(ws, 0, COL_TEAM, "グループ", STYLES.headerName);
  setStyledCell(ws, 0, COL_NAME, "勤務者", STYLES.headerName);
  for (let d = 1; d <= days; d++) {
    const style = weekends.has(d) ? weekendStyles.headerDay : STYLES.headerDay;
    const confLabel = getConferenceHeaderLabel?.(d);
    const value = confLabel ? `${d}会` : d;
    setStyledCell(ws, 0, dayColumn(d), value, style);
  }
  setStyledCell(ws, 0, offDaysColumn(days), OFF_DAYS_HEADER, STYLES.headerName);

  setStyledCell(ws, 1, COL_TEAM, "", STYLES.headerName);
  setStyledCell(ws, 1, COL_NAME, "", STYLES.headerName);
  for (let d = 1; d <= days; d++) {
    const style = weekends.has(d) ? weekendStyles.headerDow : STYLES.headerDow;
    const confLabel = getConferenceHeaderLabel?.(d);
    const value = confLabel ? `${getWeekdayLabel(year, month, d)}会` : getWeekdayLabel(year, month, d);
    setStyledCell(ws, 1, dayColumn(d), value, style);
  }
  setStyledCell(ws, 1, offDaysColumn(days), "日数", STYLES.headerDow);

  let r = 2;
  ws["!merges"] = [];

  sectionGroups.forEach((group, groupIndex) => {
    group.sections.forEach((section, sectionIndex) => {
      const groupStartRow = r;

      writeGroupSubHeaderRow(ws, r, days, weekends, variant);
      r++;

      section.members.forEach((w, memberIndex) => {
        writeWorkerRow(ws, r, days, w, weekends, variant, getConferenceStyle);
        r++;

        if (memberIndex < section.members.length - 1) {
          fillSpacerRow(ws, r, days, STYLES.spacerWithinTeam);
          r++;
        }
      });

      const membersEndRow = r - 1;
      applyTeamColumnMerge(ws, groupStartRow, membersEndRow, getGroupLabel(section));

      const hasSubSections = group.sections.some((s) => s.subGroup);
      if (variant === "shift" && (!group.team || hasSubSections)) {
        setStyledCell(ws, r, COL_TEAM, "", STYLES.workerName);
        setStyledCell(ws, r, COL_NAME, "出勤合計", STYLES.footerLabel);
        for (let d = 1; d <= days; d++) {
          const total = countAttendingFromWorkerCells(section.members, d);
          const deviant = Boolean(getGroupTotalOutOfRange?.(section, d, total));
          setStyledCell(ws, r, dayColumn(d), total, deviant ? STYLES.footerCellDeviation : STYLES.footerCell);
        }
        setStyledCell(ws, r, offDaysColumn(days), "", STYLES.footerCell);
        r++;
      }

      const isLastSection = sectionIndex === group.sections.length - 1;
      if (!isLastSection) {
        fillSpacerRow(ws, r, days, STYLES.spacerTeamBoundary);
        r++;
      }
    });

    if (variant === "shift" && group.team) {
      const hasSubSections = group.sections.some((s) => s.subGroup);
      const teamMembers = workers.filter((w) => w.teamId === group.team.id);
      const labelStyle = hasSubSections ? STYLES.teamTotalLabel : STYLES.footerLabel;
      const cellStyle = hasSubSections ? STYLES.teamTotalCell : STYLES.footerCell;
      setStyledCell(ws, r, COL_TEAM, hasSubSections ? group.team.name : "", labelStyle);
      setStyledCell(ws, r, COL_NAME, hasSubSections ? "出勤合計（メイン）" : "出勤合計", labelStyle);
      for (let d = 1; d <= days; d++) {
        const total = countAttendingFromWorkerCells(teamMembers, d);
        const deviant = Boolean(getTeamTotalOutOfRange?.(group.team, d, total));
        setStyledCell(ws, r, dayColumn(d), total, deviant ? STYLES.footerCellDeviation : cellStyle);
      }
      setStyledCell(ws, r, offDaysColumn(days), "", cellStyle);
      r++;
    }

    if (groupIndex < sectionGroups.length - 1) {
      fillSpacerRow(ws, r, days, STYLES.spacerTeamBoundary);
      r++;
    }
  });

  let footerRow = null;
  if (variant === "shift" && (getSupervisorCount || getFooterLabel)) {
    footerRow = r;
    setStyledCell(ws, r, COL_TEAM, "", STYLES.footerLabel);
    setStyledCell(ws, r, COL_NAME, getSupervisorCount ? "責任者" : "責任者不在", STYLES.footerLabel);
    for (let d = 1; d <= days; d++) {
      if (getSupervisorCount) {
        const count = Number(getSupervisorCount(d)) || 0;
        setStyledCell(ws, r, dayColumn(d), count, count === 0 ? STYLES.footerCellDeviation : STYLES.footerCell);
      } else {
        const label = getFooterLabel(d) || "";
        setStyledCell(ws, r, dayColumn(d), label, label ? STYLES.footerCellAbsent : STYLES.footerCell);
      }
    }
    setStyledCell(ws, r, offDaysColumn(days), "", STYLES.footerCell);
    r++;
  }

  ws["!ref"] = XLSX.utils.encode_range({
    s: { c: 0, r: 0 },
    e: { c: lastColumnIndex(days), r: Math.max(r - 1, 1) },
  });
  ws["!cols"] = [{ wch: 12 }, { wch: 14 }, ...Array.from({ length: days }, () => ({ wch: 6 })), { wch: 10 }];
  const rows = [{ hpt: 22 }, { hpt: 20 }];
  if (footerRow != null) rows[footerRow] = { hpt: 36 };
  ws["!rows"] = rows;
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

/** グループ先頭の日付サブヘッダー行（勤務者名が空で1日目が 1） */
export function isGroupSubHeaderRow(row, nameCol) {
  if (String(row[nameCol] ?? "").trim()) return false;
  const day1 = row[nameCol + 1];
  if (typeof day1 === "number" && day1 === 1) return true;
  return String(day1 ?? "").trim() === "1";
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
