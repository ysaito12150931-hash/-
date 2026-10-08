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

/** A=印, B=氏名, C=区分, D以降=日付（2026勤務.xlsx の並び） */
const COL_MARK = 0;
const COL_NAME = 1;
const COL_LABEL = 2;
const DAY_START = 3;
export const OFF_DAYS_HEADER = "月間休み日数";

const GROUP_PALETTE = ["CCFF33", "CCECFF", "CCFFCC", "FFFFCC", "FFCCFF", "FDE9D9"];
const GRID_FONT = "ＭＳ ゴシック";

const gridBorder = {
  top: { style: "thin", color: { rgb: "808080" } },
  bottom: { style: "thin", color: { rgb: "808080" } },
  left: { style: "thin", color: { rgb: "808080" } },
  right: { style: "thin", color: { rgb: "808080" } },
};

function dayColumn(day) {
  return DAY_START + day - 1;
}

function offDaysColumn(days) {
  return DAY_START + days;
}

function lastColumnIndex(days) {
  return offDaysColumn(days);
}

function gridFont(rgb, bold = false, sz = 10) {
  return { name: GRID_FONT, sz, bold, color: { rgb } };
}

function gridStyle(fillRgb, fontRgb = "000000", bold = false, horizontal = "center", sz = 10) {
  return {
    font: gridFont(fontRgb, bold, sz),
    fill: fill(fillRgb),
    alignment: { horizontal, vertical: "center", wrapText: false },
    border: gridBorder,
  };
}

function titleStyle(sz = 14) {
  return {
    font: gridFont("000000", true, sz),
    alignment: { horizontal: "center", vertical: "center" },
  };
}

export function compactLabel(value) {
  return String(value ?? "")
    .replace(/[\s\u3000]/g, "")
    .trim();
}

const SKIP_CALENDAR_NAMES = new Set([
  "チーム",
  "備考",
  "出勤合計",
  "出勤合計（メイン）",
  "責任者",
  "責任者不在",
  "合計",
  "総合計",
  "リーダー",
  "朝礼",
  "日付",
  "氏名",
  "勤務調整",
  "グループ",
  "勤務者",
]);

export function isSkippedCalendarName(name) {
  const n = compactLabel(name);
  if (!n) return false;
  if (SKIP_CALENDAR_NAMES.has(n)) return true;
  if (n.startsWith("出勤合計") || n.startsWith("専従")) return true;
  if (n.includes("未所属") || n.includes("責任者")) return true;
  return false;
}

/** 日付ヘッダーとして読める 1〜31。読めなければ null */
export function readDayNumber(cell) {
  if (typeof cell === "number" && cell >= 1 && cell <= 31) return Math.floor(cell);
  if (cell instanceof Date && !Number.isNaN(cell.getTime())) return cell.getDate();
  const s = String(cell ?? "").trim();
  if (!s) return null;
  if (/^\d{1,2}$/.test(s)) {
    const n = parseInt(s, 10);
    if (n >= 1 && n <= 31) return n;
  }
  const m = s.match(/^(\d{1,2})\s*日$/);
  if (m) {
    const n = parseInt(m[1], 10);
    if (n >= 1 && n <= 31) return n;
  }
  return null;
}

function blankRow(ws, r, days) {
  for (let c = COL_MARK; c <= lastColumnIndex(days); c++) {
    setStyledCell(ws, r, c, "", {
      font: gridFont("000000", false, 9),
      fill: fill("FFFFFF"),
      border: gridBorder,
    });
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

function countWorkerAttendingDays(worker) {
  return (worker.cells || []).filter((cell) => normalizeCellValue(cell).attending).length;
}

function dayHeaderStyle(weekend) {
  return gridStyle("FFFFFF", weekend ? "FF0000" : "000000", true, "center", 9);
}

function nameHeaderStyle() {
  return gridStyle("FFFFFF", "000000", true, "center", 10);
}

function writeWorkerRow(ws, r, days, worker, variant, getConferenceStyle, groupRgb) {
  const rowStyle = gridStyle(groupRgb, "000000", false, "center", 10);
  const nameStyle = gridStyle(groupRgb, "000000", true, "left", 10);
  const mark = worker.isSupervisor || worker.isGroupSupervisor ? "★" : "";
  setStyledCell(ws, r, COL_MARK, mark, gridStyle(groupRgb, "FF0000", true, "center", 10));
  setStyledCell(ws, r, COL_NAME, worker.name, nameStyle);
  setStyledCell(ws, r, COL_LABEL, "", rowStyle);
  for (let d = 1; d <= days; d++) {
    const cell = normalizeCellValue(worker.cells?.[d - 1] ?? "");
    let style = variant === "template" ? rowStyle : gridStyle("FFFFFF");
    const confStyle = getConferenceStyle?.(worker, d);
    if (variant === "shift" && confStyle) {
      style = makeStyle(confStyle.color.replace("#", ""), confStyle.bg.replace("#", ""), true);
    } else if (variant === "shift" && cell.preferredOff) {
      style = STYLES.workerCellPreferredOff;
    } else if (variant === "shift" && cell.off) {
      style = STYLES.workerCellAutoOff;
    } else if (variant === "shift" && cell.conference) {
      style = STYLES.workerCellConference;
    }
    setStyledCell(ws, r, dayColumn(d), variant === "template" ? "" : cell.text, style);
  }
  const offDays = worker.monthlyOffDays;
  const rightValue = variant === "shift" ? countWorkerAttendingDays(worker) : offDays == null || offDays === "" ? "" : offDays;
  setStyledCell(ws, r, offDaysColumn(days), rightValue, rowStyle);
}

function writeDateHeaderRow(ws, r, days, weekends, groupLabel, groupRgb, rightHeader, getConferenceHeaderLabel) {
  const head = gridStyle(groupRgb, "000000", true, "center", 9);
  setStyledCell(ws, r, COL_MARK, "", head);
  setStyledCell(ws, r, COL_NAME, "日　付", head);
  setStyledCell(ws, r, COL_LABEL, groupLabel, gridStyle(groupRgb, "000000", true, "left", 8));
  for (let d = 1; d <= days; d++) {
    const confLabel = getConferenceHeaderLabel?.(d);
    const value = confLabel ? `${d}会` : d;
    setStyledCell(ws, r, dayColumn(d), value, dayHeaderStyle(weekends.has(d)));
  }
  setStyledCell(ws, r, offDaysColumn(days), rightHeader, nameHeaderStyle());
}

function writeNameHeaderRow(ws, r, year, month, days, weekends) {
  setStyledCell(ws, r, COL_MARK, "", nameHeaderStyle());
  setStyledCell(ws, r, COL_NAME, "氏　名", nameHeaderStyle());
  setStyledCell(ws, r, COL_LABEL, "", nameHeaderStyle());
  for (let d = 1; d <= days; d++) {
    setStyledCell(ws, r, dayColumn(d), getWeekdayLabel(year, month, d), dayHeaderStyle(weekends.has(d)));
  }
  setStyledCell(ws, r, offDaysColumn(days), "日数", nameHeaderStyle());
}

function writeTotalRow(ws, r, days, label, members, groupRgb, variant, deviantAt) {
  const style = gridStyle(groupRgb, "000000", true, "center", 10);
  const nameStyle = gridStyle(groupRgb, "000000", true, "left", 10);
  setStyledCell(ws, r, COL_MARK, members.length, style);
  setStyledCell(ws, r, COL_NAME, label, nameStyle);
  setStyledCell(ws, r, COL_LABEL, "", style);
  for (let d = 1; d <= days; d++) {
    if (variant !== "shift") {
      setStyledCell(ws, r, dayColumn(d), "", style);
      continue;
    }
    const total = countAttendingFromWorkerCells(members, d);
    const deviant = Boolean(deviantAt?.(d, total));
    setStyledCell(ws, r, dayColumn(d), total, deviant ? STYLES.footerCellDeviation : style);
  }
  const right =
    variant === "shift" ? members.reduce((sum, w) => sum + countWorkerAttendingDays(w), 0) : "";
  setStyledCell(ws, r, offDaysColumn(days), right, style);
}

function getGroupLabel(section) {
  if (section.team?.name && section.subGroup?.name) {
    return `${section.team.name} / ${section.subGroup.name}`;
  }
  return section.label || "";
}

/**
 * 2026勤務.xlsx に近い表。
 * 1行目: 年・月 / 2行目: 勤務調整 / 各グループ: 日　付・氏　名（土日は赤字）・勤務者・合計
 * @param {WorkSheet} ws
 * @param {number} year
 * @param {number} month
 * @param {number} days
 * @param {{ id?: string, name: string, teamId?: string|null, subGroupId?: string|null, isSupervisor?: boolean, isGroupSupervisor?: boolean, monthlyOffDays?: number, cells?: string[] }[]} workers
 * @param {{ id: string, name: string }[]} teams
 * @param {{ id: string, name: string, teamId: string }[]} [subGroups]
 * @param {{ variant?: 'template'|'shift', getConferenceHeaderLabel?: (day:number)=>boolean, getConferenceStyle?: (worker:object, day:number)=>object|null, getFooterLabel?: (day:number)=>string, getSupervisorCount?: (day:number)=>number, getGroupTotalOutOfRange?: (section:object, day:number, total:number)=>boolean, getTeamTotalOutOfRange?: (team:object, day:number, total:number)=>boolean }} [options]
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
  const sections = getWorkerSections(workers, teams, subGroups);
  const sectionGroups = groupWorkerSectionsByTeam(sections);
  const rightHeader = variant === "shift" ? "出勤" : OFF_DAYS_HEADER;

  setStyledCell(ws, 0, COL_NAME, year, titleStyle(16));
  setStyledCell(ws, 0, COL_LABEL, "年", titleStyle(14));
  setStyledCell(ws, 0, DAY_START, month, titleStyle(16));
  setStyledCell(ws, 0, DAY_START + 1, "月", titleStyle(14));
  setStyledCell(ws, 1, COL_NAME, "勤務調整", titleStyle(12));

  let r = 2;
  let colorIndex = 0;
  ws["!merges"] = [];

  const blocks = sections.length ? sections : [{ label: "", members: [], team: null, subGroup: null }];

  blocks.forEach((section, sectionIndex) => {
    const groupRgb = GROUP_PALETTE[colorIndex % GROUP_PALETTE.length];
    colorIndex += 1;
    const groupLabel = getGroupLabel(section);

    writeDateHeaderRow(ws, r, days, weekends, groupLabel, groupRgb, rightHeader, getConferenceHeaderLabel);
    r++;
    writeNameHeaderRow(ws, r, year, month, days, weekends);
    r++;

    section.members.forEach((worker) => {
      writeWorkerRow(ws, r, days, worker, variant, getConferenceStyle, groupRgb);
      r++;
    });

    writeTotalRow(ws, r, days, "合計", section.members, groupRgb, variant, (day, total) =>
      getGroupTotalOutOfRange?.(section, day, total)
    );
    r++;

    const group = sectionGroups.find((g) => g.sections.includes(section));
    const isLastInTeam = group ? group.sections[group.sections.length - 1] === section : true;
    if (variant === "shift" && group?.team && group.sections.length > 1 && isLastInTeam) {
      const teamMembers = workers.filter((w) => w.teamId === group.team.id);
      const teamRgb = "CCFFFF";
      writeTotalRow(ws, r, days, "合計", teamMembers, teamRgb, variant, (day, total) =>
        getTeamTotalOutOfRange?.(group.team, day, total)
      );
      setStyledCell(ws, r, COL_LABEL, group.team.name, gridStyle(teamRgb, "000000", true, "left", 8));
      r++;
    }

    if (sectionIndex < blocks.length - 1) {
      blankRow(ws, r, days);
      r++;
    }
  });

  if (variant === "shift") {
    blankRow(ws, r, days);
    r++;
    writeTotalRow(ws, r, days, "総合計", workers, "FFCCFF", "shift");
    r++;

    if (getSupervisorCount || getFooterLabel) {
      const foot = gridStyle("FFCCFF", "000000", true, "center", 10);
      setStyledCell(ws, r, COL_MARK, "", foot);
      setStyledCell(ws, r, COL_NAME, getSupervisorCount ? "責任者" : "責任者不在", gridStyle("FFCCFF", "000000", true, "left", 10));
      setStyledCell(ws, r, COL_LABEL, "", foot);
      for (let d = 1; d <= days; d++) {
        if (getSupervisorCount) {
          const count = Number(getSupervisorCount(d)) || 0;
          setStyledCell(ws, r, dayColumn(d), count, count === 0 ? STYLES.footerCellDeviation : foot);
        } else {
          const label = getFooterLabel(d) || "";
          setStyledCell(ws, r, dayColumn(d), label, label ? STYLES.footerCellAbsent : foot);
        }
      }
      setStyledCell(ws, r, offDaysColumn(days), "", foot);
      r++;
    }
  }

  ws["!ref"] = XLSX.utils.encode_range({
    s: { c: 0, r: 0 },
    e: { c: lastColumnIndex(days), r: Math.max(r - 1, 1) },
  });
  ws["!cols"] = [
    { wch: 3 },
    { wch: 14 },
    { wch: 12 },
    ...Array.from({ length: days }, () => ({ wch: 4.2 })),
    { wch: 12 },
  ];
  ws["!rows"] = [{ hpt: 18 }, { hpt: 16 }];
  ws["!views"] = [
    {
      state: "frozen",
      xSplit: 3,
      ySplit: 4,
      topLeftCell: "D5",
      activeCell: "D5",
    },
  ];
}

export function detectCalendarDataStartRow(rows) {
  return detectSheetGeometry(rows).dataStartRow;
}

/**
 * 氏名列・日付行・月間休み日数列・データ開始行。
 * 旧形式（1行目が勤務者）と、日　付 / 氏　名 形式の両方を読む。
 */
export function detectSheetGeometry(rows) {
  const limit = Math.min(rows.length, 30);
  let nameCol = null;
  let nameHeaderRow = -1;
  for (let r = 0; r < limit; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const h = compactLabel(row[c]);
      if (nameHeaderRow < 0 && (h === "勤務者" || h === "氏名")) {
        nameCol = c;
        nameHeaderRow = r;
      }
    }
  }
  if (nameCol == null) {
    nameCol = detectNameColumn(rows[0] || []);
    nameHeaderRow = 0;
  }

  let dateRow = -1;
  for (let r = 0; r < limit; r++) {
    const row = rows[r] || [];
    if (row.some((cell) => compactLabel(cell) === "日付")) {
      dateRow = r;
      break;
    }
  }
  if (dateRow < 0) dateRow = nameHeaderRow >= 0 ? nameHeaderRow : 0;

  let offDaysCol = null;
  for (let r = 0; r < limit; r++) {
    const found = detectOffDaysColumn(rows[r] || []);
    if (found != null) {
      offDaysCol = found;
      break;
    }
  }

  let dataStartRow = Math.max(nameHeaderRow, 0) + 1;
  const maybeDow = rows[dataStartRow] || [];
  const nameOnDow = compactLabel(maybeDow[nameCol]);
  const dowCount = maybeDow.filter((cell) => /^[月火水木金土日]$/u.test(String(cell ?? "").trim())).length;
  if (!nameOnDow && dowCount >= 3) dataStartRow += 1;

  return { nameCol, nameHeaderRow, dateRow, offDaysCol, dataStartRow };
}

/** グループ先頭の日付サブヘッダー行（勤務者名が空で1日目が 1） */
export function isGroupSubHeaderRow(row, nameCol, dayCol) {
  const name = compactLabel(row[nameCol]);
  if (name === "日付" || name === "氏名" || name === "勤務調整") return true;
  if (name) return false;
  const dayIndex = dayCol ?? nameCol + 1;
  return readDayNumber(row[dayIndex]) === 1;
}

/** @param {unknown[]} header */
export function detectOffDaysColumn(header) {
  for (let c = 0; c < header.length; c++) {
    const h = compactLabel(header[c]);
    if (h === compactLabel(OFF_DAYS_HEADER) || h.includes("月間休み")) return c;
  }
  return null;
}

/** @param {unknown[]} header */
export function detectNameColumn(header) {
  for (let c = 0; c < header.length; c++) {
    const h = compactLabel(header[c]);
    if (h === "勤務者" || h === "氏名") return c;
  }
  const a = compactLabel(header[0]);
  if (a === "チーム" || a.includes("チーム") || a === "グループ" || a.includes("グループ")) return 1;
  return 0;
}
