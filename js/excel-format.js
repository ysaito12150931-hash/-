/** @typedef {import('xlsx').WorkSheet} WorkSheet */

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

function workerKey(w) {
  return w.id ?? w.name;
}

function hasValidTeam(worker, teams) {
  return Boolean(worker.teamId && teams.some((t) => t.id === worker.teamId));
}

function sortByName(list) {
  return [...list].sort((a, b) => a.name.localeCompare(b.name, "ja"));
}

/** 責任者を先頭、それ以外を名前順 */
function sortTeamMembers(members) {
  const supervisors = sortByName(members.filter((w) => w.isSupervisor));
  const others = sortByName(members.filter((w) => !w.isSupervisor));
  return [...supervisors, ...others];
}

/**
 * @param {{ id?: string, name: string, teamId?: string|null, isSupervisor?: boolean, cells?: string[] }[]} workers
 * @param {{ id: string, name: string }[]} teams
 */
export function groupWorkersByTeam(workers, teams) {
  const groups = [];
  const assigned = new Set();

  const mark = (list) => list.forEach((w) => assigned.add(workerKey(w)));
  const isUnassigned = (w) => !assigned.has(workerKey(w));

  const supervisorsNoTeam = sortByName(
    workers.filter((w) => w.isSupervisor && !hasValidTeam(w, teams))
  );
  if (supervisorsNoTeam.length) {
    groups.push({ team: null, members: supervisorsNoTeam });
    mark(supervisorsNoTeam);
  }

  for (const team of teams) {
    const members = sortTeamMembers(workers.filter((w) => w.teamId === team.id));
    if (members.length) {
      groups.push({ team, members });
      mark(members);
    }
  }

  const unassigned = sortByName(workers.filter((w) => isUnassigned(w)));
  if (unassigned.length) {
    groups.push({ team: null, members: unassigned });
  }

  return groups;
}

const COL_TEAM = 0;
const COL_NAME = 1;

function dayColumn(day) {
  return COL_NAME + day;
}

function lastColumnIndex(days) {
  return dayColumn(days);
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
}

function getGroupLabel(group) {
  if (group.team?.name) return group.team.name;
  if (group.members.length > 0 && group.members.every((w) => w.isSupervisor)) {
    return "責任者（チーム未所属）";
  }
  return "（未所属）";
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
 * @param {{ id?: string, name: string, teamId?: string|null, isSupervisor?: boolean, cells?: string[] }[]} workers
 * @param {{ id: string, name: string }[]} teams
 */
export function fillCalendarTemplateSheet(ws, year, month, days, workers, teams = []) {
  const weekends = getWeekendDaySet(year, month, days);
  const groups = groupWorkersByTeam(workers, teams);

  setStyledCell(ws, 0, COL_TEAM, "チーム", STYLES.headerName);
  setStyledCell(ws, 0, COL_NAME, "勤務者", STYLES.headerName);
  for (let d = 1; d <= days; d++) {
    const style = weekends.has(d) ? STYLES.headerDayWeekend : STYLES.headerDay;
    setStyledCell(ws, 0, dayColumn(d), d, style);
  }

  setStyledCell(ws, 1, COL_TEAM, "", STYLES.headerName);
  setStyledCell(ws, 1, COL_NAME, "", STYLES.headerName);
  for (let d = 1; d <= days; d++) {
    const style = weekends.has(d) ? STYLES.headerDowWeekend : STYLES.headerDow;
    setStyledCell(ws, 1, dayColumn(d), getWeekdayLabel(year, month, d), style);
  }

  let r = 2;
  ws["!merges"] = [];

  groups.forEach((group, groupIndex) => {
    const groupStartRow = r;

    group.members.forEach((w, memberIndex) => {
      writeWorkerRow(ws, r, days, w, weekends);
      r++;

      if (memberIndex < group.members.length - 1) {
        fillSpacerRow(ws, r, days, STYLES.spacerWithinTeam);
        r++;
      }
    });

    applyTeamColumnMerge(ws, groupStartRow, r - 1, getGroupLabel(group));

    if (groupIndex < groups.length - 1) {
      fillSpacerRow(ws, r, days, STYLES.spacerTeamBoundary);
      r++;
    }
  });

  ws["!ref"] = XLSX.utils.encode_range({
    s: { c: 0, r: 0 },
    e: { c: lastColumnIndex(days), r: Math.max(r - 1, 1) },
  });
  ws["!cols"] = [{ wch: 12 }, { wch: 14 }, ...Array.from({ length: days }, () => ({ wch: 6 }))];
  ws["!rows"] = [{ hpt: 22 }, { hpt: 20 }];
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
export function detectNameColumn(header) {
  for (let c = 0; c < header.length; c++) {
    const h = String(header[c] ?? "").trim();
    if (h === "勤務者" || h.includes("勤務者")) return c;
  }
  const a = String(header[0] ?? "").trim();
  if (a === "チーム" || a.includes("チーム")) return 1;
  return 0;
}
