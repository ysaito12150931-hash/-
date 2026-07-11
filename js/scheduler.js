import { getDaysInMonth } from "./store.js";

const MAX_ATTEMPTS = 200;

/**
 * @returns {{ ok: boolean, assignments?: object, messages: string[], stats?: object, year?: number, month?: number, workers?: object[] }}
 */
export function generateShift(state) {
  const {
    year,
    month,
    workers,
    teams,
    subGroups = [],
    subGroupConstraints = {},
    constraints,
    teamConstraints,
    preferences,
    maxConsecutiveWork,
    useShiftTypes,
    shiftTypes,
  } = state;
  const days = getDaysInMonth(year, month);
  const messages = [];

  if (!workers.length) {
    return { ok: false, messages: ["勤務者を1人以上登録してください。"] };
  }

  const locked = buildPreferenceLocks(workers, preferences, days);
  const { lockedOff, lockedWork, halfOff } = locked;

  for (const w of workers) {
    const pref = preferences[w.name] || {};
    const prefOffTotal = countPreferenceOffDays(pref, days);
    if (prefOffTotal > (w.monthlyOffDays ?? 0)) {
      messages.push(
        `${w.name}: Excel希望休（${formatOffDays(prefOffTotal)}日）が月間休み日数（${w.monthlyOffDays}日）を超えています。`
      );
    }
    const lockedWorkCount = countLockedWorkDays(lockedWork, w.id, days);
    const halfOffCount = countHalfOffDays(halfOff, w.id, days);
    const maxOffPossible = days - lockedWorkCount - halfOffCount * 0.5;
    if ((w.monthlyOffDays ?? 0) > maxOffPossible + 0.001) {
      messages.push(
        `${w.name}: 出勤希望・半休が多く、月間休み日数（${w.monthlyOffDays}日）を満たしにくい可能性があります。`
      );
    }

    const lockedOffCount = countLockedOffDays(lockedOff, w.id, days);
    const effectiveLockedOff = lockedOffCount + halfOffCount * 0.5;
    const target = w.monthlyOffDays ?? 0;
    if (Math.abs(target - effectiveLockedOff) < 0.001) {
      const row = {};
      for (let d = 1; d <= days; d++) {
        if (lockedOff[w.id]?.[d]) row[d] = false;
        else if (halfOff[w.id]?.[d]) row[d] = true;
        else row[d] = true;
      }
      if (violatesConsecutiveWork(row, days, maxConsecutiveWork)) {
        messages.push(
          `${w.name}: Excel希望休が月間休み日数（${formatOffDays(target)}日）と一致するため残りは全出勤になりますが、連勤上限（${maxConsecutiveWork}日）を超えます。休み希望の日を分散するか、月間休み日数・連勤上限を見直してください。`
        );
      }
    }
  }

  let best = null;
  let bestScore = -Infinity;

  for (const requireSupervisor of [true, false]) {
    for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
      const seed = attempt * 7919 + year * 100 + month;
      const grid = tryBuildSchedule({
        workers,
        days,
        lockedOff,
        lockedWork,
        halfOff,
        maxConsecutiveWork,
        constraints,
        teamConstraints,
        teams,
        subGroups,
        subGroupConstraints,
        requireSupervisor,
        seed,
      });
      if (!grid) continue;
      const score = scoreSchedule(grid, workers, days, preferences, requireSupervisor);
      if (score > bestScore) {
        bestScore = score;
        best = grid;
      }
    }
    if (best) break;
  }

  if (!best) {
    return {
      ok: false,
      messages: [
        ...messages,
        "制約を満たすシフトを生成できませんでした。人数制約・休み日数・連勤上限・Excel希望を見直してください。",
      ],
    };
  }

  applyShiftTypes(best, workers, days, useShiftTypes, shiftTypes);
  const assignments = gridToAssignments(best, workers, days, halfOff);
  const stats = buildStats(assignments, workers, days);
  const conferenceDays = computeConferenceDays(
    assignments,
    workers,
    teams,
    teamConstraints,
    subGroups,
    subGroupConstraints,
    year,
    month
  );

  return {
    ok: true,
    assignments,
    messages: messages.length ? messages : ["シフトを生成しました。"],
    stats,
    conferenceDays,
    year,
    month,
    workers,
    teams,
  };
}

function buildPreferenceLocks(workers, preferences, days) {
  const lockedOff = {};
  const lockedWork = {};
  const halfOff = {};
  for (const w of workers) {
    lockedOff[w.id] = {};
    lockedWork[w.id] = {};
    halfOff[w.id] = {};
    const pref = preferences[w.name] || {};
    for (let d = 1; d <= days; d++) {
      const v = pref[d];
      if (v === "off" || v === true) lockedOff[w.id][d] = true;
      else if (v === "work") lockedWork[w.id][d] = true;
      else if (v === "am-off") halfOff[w.id][d] = "am";
      else if (v === "pm-off") halfOff[w.id][d] = "pm";
    }
  }
  return { lockedOff, lockedWork, halfOff };
}

function countPreferenceOffDays(pref, days) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    const v = pref?.[d];
    if (v === "off" || v === true) n += 1;
    else if (v === "am-off" || v === "pm-off") n += 0.5;
  }
  return n;
}

function formatOffDays(n) {
  return Number.isInteger(n) ? String(n) : n.toFixed(1);
}

function countHalfOffDays(halfOff, workerId, days) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    if (halfOff[workerId]?.[d]) n++;
  }
  return n;
}

function countLockedOffDays(lockedOff, workerId, days) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    if (lockedOff[workerId]?.[d]) n++;
  }
  return n;
}

function countLockedWorkDays(lockedWork, workerId, days) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    if (lockedWork[workerId]?.[d]) n++;
  }
  return n;
}

function countEffectiveOffDays(grid, workerId, days, halfOff) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    if (!grid[workerId][d]) n += 1;
    else if (halfOff[workerId]?.[d]) n += 0.5;
  }
  return n;
}

function offDaysMatch(current, target) {
  return Math.abs(current - target) < 0.001;
}

function countWorkerOffs(grid, workerId, days, halfOff) {
  return countEffectiveOffDays(grid, workerId, days, halfOff);
}

function tryBuildSchedule(ctx) {
  const {
    workers,
    days,
    lockedOff,
    lockedWork,
    halfOff,
    maxConsecutiveWork,
    constraints,
    teamConstraints,
    teams,
    subGroups = [],
    subGroupConstraints = {},
    requireSupervisor,
    seed,
  } = ctx;
  const groupCtx = { constraints, teamConstraints, teams, subGroups, subGroupConstraints };
  const rng = mulberry32(seed);

  const grid = {};
  for (const w of workers) {
    grid[w.id] = {};
    for (let d = 1; d <= days; d++) {
      grid[w.id][d] = false;
    }
  }

  for (let d = 1; d <= days; d++) {
    if (!assignDay(grid, workers, d, groupCtx, lockedOff, lockedWork, halfOff, requireSupervisor, rng)) {
      return null;
    }
  }

  const targets = Object.fromEntries(workers.map((w) => [w.id, w.monthlyOffDays ?? 0]));

  for (let pass = 0; pass < days * workers.length * 4; pass++) {
    let changed = false;
    for (const w of shuffledArray(workers, rng)) {
      const cur = countEffectiveOffDays(grid, w.id, days, halfOff);
      const target = targets[w.id];
      if (cur < target - 0.001) {
        if (tryAddOff(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutiveWork, groupCtx, workers, requireSupervisor, rng)) {
          changed = true;
        }
      } else if (cur > target + 0.001) {
        if (tryAddWork(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutiveWork, groupCtx, workers, requireSupervisor, rng)) {
          changed = true;
        }
      }
    }
    if (!changed) break;
  }

  for (let iter = 0; iter < days * workers.length * 3; iter++) {
    let fixed = false;
    for (let d = 1; d <= days; d++) {
      if (validateDay(grid, workers, d, groupCtx, requireSupervisor).ok) {
        continue;
      }
      if (repairDay(grid, workers, d, days, lockedOff, lockedWork, halfOff, maxConsecutiveWork, groupCtx, requireSupervisor, rng)) {
        fixed = true;
      }
    }
    if (!fixed) break;
  }

  for (const w of workers) {
    if (!offDaysMatch(countEffectiveOffDays(grid, w.id, days, halfOff), targets[w.id])) return null;
  }

  for (let d = 1; d <= days; d++) {
    if (!validateDay(grid, workers, d, groupCtx, requireSupervisor).ok) {
      return null;
    }
  }

  for (const w of workers) {
    for (let d = 1; d <= days; d++) {
      if (lockedOff[w.id]?.[d] && grid[w.id][d]) return null;
      if (lockedWork[w.id]?.[d] && !grid[w.id][d]) return null;
      if (halfOff[w.id]?.[d] && !grid[w.id][d]) return null;
    }
  }

  for (const w of workers) {
    if (violatesConsecutiveWork(grid[w.id], days, maxConsecutiveWork)) return null;
  }

  return grid;
}

function assignDay(grid, workers, day, groupCtx, lockedOff, lockedWork, halfOff, requireSupervisor, rng) {
  for (const w of workers) {
    if (lockedOff[w.id]?.[day]) grid[w.id][day] = false;
    else grid[w.id][day] = true;
  }
  if (validateDay(grid, workers, day, groupCtx, requireSupervisor).ok) {
    return true;
  }

  for (let attempt = 0; attempt < workers.length * 4; attempt++) {
    const working = shuffledArray(
      workers.filter(
        (w) => grid[w.id][day] && !lockedOff[w.id]?.[day] && !lockedWork[w.id]?.[day] && !halfOff[w.id]?.[day]
      ),
      rng
    );
    for (const w of working) {
      grid[w.id][day] = false;
      if (validateDay(grid, workers, day, groupCtx, requireSupervisor).ok) {
        return true;
      }
      grid[w.id][day] = true;
    }

    const offPool = shuffledArray(
      workers.filter((w) => !grid[w.id][day] && !lockedOff[w.id]?.[day]),
      rng
    );
    for (const w of offPool) {
      grid[w.id][day] = true;
      if (validateDay(grid, workers, day, groupCtx, requireSupervisor).ok) {
        return true;
      }
      grid[w.id][day] = false;
    }
  }

  return validateDay(grid, workers, day, groupCtx, requireSupervisor).ok;
}

function tryAddOff(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng) {
  const candidates = shuffledRange(1, days, rng);
  for (const d of candidates) {
    if (lockedOff[w.id]?.[d]) continue;
    if (lockedWork[w.id]?.[d]) continue;
    if (halfOff[w.id]?.[d]) continue;
    if (!grid[w.id][d]) continue;
    grid[w.id][d] = false;
    if (violatesConsecutiveWork(grid[w.id], days, maxConsecutive)) {
      grid[w.id][d] = true;
      continue;
    }
    if (!validateDay(grid, workers, d, groupCtx, requireSupervisor).ok) {
      grid[w.id][d] = true;
      continue;
    }
    return true;
  }
  return false;
}

function repairDay(grid, workers, day, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, requireSupervisor, rng) {
  for (const w of shuffledArray(workers, rng)) {
    if (tryAddWork(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng)) {
      if (validateDay(grid, workers, day, groupCtx, requireSupervisor).ok) return true;
    }
    if (tryAddOff(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng)) {
      if (validateDay(grid, workers, day, groupCtx, requireSupervisor).ok) return true;
    }
  }
  return false;
}

function tryAddWork(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng) {
  void lockedWork;
  void halfOff;
  const candidates = shuffledRange(1, days, rng);
  for (const d of candidates) {
    if (lockedOff[w.id]?.[d]) continue;
    if (grid[w.id][d]) continue;
    grid[w.id][d] = true;
    if (violatesConsecutiveWork(grid[w.id], days, maxConsecutive)) {
      grid[w.id][d] = false;
      continue;
    }
    if (!validateDay(grid, workers, d, groupCtx, requireSupervisor).ok) {
      grid[w.id][d] = false;
      continue;
    }
    return true;
  }
  return false;
}

function validateDay(grid, workers, day, groupCtx, requireSupervisor) {
  const { constraints, teamConstraints, teams, subGroups = [], subGroupConstraints = {} } = groupCtx;
  const working = workers.filter((w) => grid[w.id][day]);
  const count = working.length;

  const supCount = working.filter((w) => w.isSupervisor).length;
  if (supCount < constraints.supervisorMin || supCount > constraints.supervisorMax) {
    return { ok: false };
  }

  if (requireSupervisor && supCount < 1 && workers.some((w) => w.isSupervisor) && count > 0) {
    return { ok: false };
  }

  for (const team of teams) {
    const tc = teamConstraints[team.id];
    if (!tc) continue;
    const teamCount = working.filter((w) => w.teamId === team.id).length;
    if (teamCount < tc.min || teamCount > tc.max) {
      return { ok: false };
    }
  }

  for (const sg of subGroups) {
    const sgc = subGroupConstraints[sg.id];
    if (!sgc) continue;
    const sgCount = working.filter((w) => w.subGroupId === sg.id).length;
    if (sgCount < sgc.min || sgCount > sgc.max) {
      return { ok: false };
    }
  }

  return { ok: true };
}

function violatesConsecutiveWork(row, days, maxConsecutive) {
  let streak = 0;
  for (let d = 1; d <= days; d++) {
    if (row[d]) {
      streak++;
      if (streak > maxConsecutive) return true;
    } else {
      streak = 0;
    }
  }
  return false;
}

function applyShiftTypes(grid, workers, days, useShiftTypes, shiftTypes) {
  if (!useShiftTypes || !shiftTypes?.length) return;
  for (let d = 1; d <= days; d++) {
    const working = workers.filter((w) => grid[w.id][d]);
    working.forEach((w, i) => {
      grid[w.id][`_type_${d}`] = shiftTypes[i % shiftTypes.length];
    });
  }
}

function gridToAssignments(grid, workers, days, halfOff) {
  const assignments = {};
  for (const w of workers) {
    assignments[w.id] = {};
    for (let d = 1; d <= days; d++) {
      if (!grid[w.id][d]) {
        assignments[w.id][d] = { type: "off" };
      } else if (halfOff[w.id]?.[d]) {
        assignments[w.id][d] = { type: "half-off", half: halfOff[w.id][d] };
      } else {
        assignments[w.id][d] = {
          type: "work",
          shiftType: grid[w.id][`_type_${d}`] || null,
        };
      }
    }
  }
  return assignments;
}

function scoreSchedule(grid, workers, days, preferences, requireSupervisor) {
  let score = 0;
  for (const w of workers) {
    const pref = preferences[w.name] || {};
    for (let d = 1; d <= days; d++) {
      const p = pref[d];
      if (p === "off" || p === true) {
        if (!grid[w.id][d]) score += 10;
        if (grid[w.id][d]) score -= 20;
      }
      if (p === "work") {
        if (grid[w.id][d]) score += 10;
        if (!grid[w.id][d]) score -= 20;
      }
      if (p === "am-off" || p === "pm-off") {
        if (grid[w.id][d]) score += 8;
        if (!grid[w.id][d]) score -= 16;
      }
    }
  }
  for (let d = 1; d <= days; d++) {
    const sup = workers.filter((w) => w.isSupervisor && grid[w.id][d]).length;
    if (requireSupervisor && sup >= 1) score += 2;
  }
  return score;
}

function buildStats(assignments, workers, days) {
  const daily = [];
  for (let d = 1; d <= days; d++) {
    const working = workers.filter((w) => {
      const c = assignments[w.id][d];
      return c && c.type !== "off";
    });
    daily.push({
      day: d,
      total: working.length,
      supervisors: working.filter((w) => w.isSupervisor).length,
    });
  }
  return { daily };
}

/** 月曜始まりの週キー */
function getWeekKey(year, month, day) {
  const date = new Date(year, month - 1, day);
  const mondayOffset = (date.getDay() + 6) % 7;
  const monday = new Date(year, month - 1, day - mondayOffset);
  return `${monday.getFullYear()}-${monday.getMonth() + 1}-${monday.getDate()}`;
}

/**
 * グループごとに週1日、出勤人数が最多の日をカンファレンス日とする
 * @returns {{ teams: Record<string, { day: number, count: number, weekKey: string }[]>, subGroups: Record<string, { day: number, count: number, weekKey: string }[]> }}
 */
export function normalizeConferenceDays(conferenceDays) {
  if (conferenceDays?.teams || conferenceDays?.subGroups) {
    return {
      teams: conferenceDays.teams || {},
      subGroups: conferenceDays.subGroups || {},
    };
  }
  return { teams: conferenceDays || {}, subGroups: {} };
}

function pickWeeklyConferenceDays(members, assignments, year, month) {
  const days = getDaysInMonth(year, month);
  if (!members.length) return [];

  const weekToDays = new Map();
  for (let d = 1; d <= days; d++) {
    const wk = getWeekKey(year, month, d);
    if (!weekToDays.has(wk)) weekToDays.set(wk, []);
    weekToDays.get(wk).push(d);
  }

  const picked = [];
  for (const [weekKey, daysInWeek] of weekToDays) {
    let bestDay = null;
    let bestCount = -1;
    for (const d of daysInWeek) {
      const count = members.filter((m) => {
        const c = assignments[m.id][d];
        return c && c.type !== "off";
      }).length;
      if (count > bestCount || (count === bestCount && bestDay != null && d < bestDay)) {
        bestCount = count;
        bestDay = d;
      }
    }
    if (bestDay != null && bestCount > 0) {
      picked.push({ weekKey, day: bestDay, count: bestCount });
    }
  }
  picked.sort((a, b) => a.day - b.day);
  return picked;
}

export function computeConferenceDays(
  assignments,
  workers,
  teams,
  teamConstraints,
  subGroups = [],
  subGroupConstraints = {},
  year,
  month
) {
  const result = { teams: {}, subGroups: {} };

  for (const team of teams) {
    const tc = teamConstraints[team.id];
    if (!tc?.useConferenceDay) continue;
    const members = workers.filter((w) => w.teamId === team.id);
    const picked = pickWeeklyConferenceDays(members, assignments, year, month);
    if (picked.length) result.teams[team.id] = picked;
  }

  for (const sg of subGroups) {
    const sgc = subGroupConstraints[sg.id];
    if (!sgc?.useConferenceDay) continue;
    const members = workers.filter((w) => w.subGroupId === sg.id);
    const picked = pickWeeklyConferenceDays(members, assignments, year, month);
    if (picked.length) result.subGroups[sg.id] = picked;
  }

  return result;
}

export function isConferenceDayForTeam(conferenceDays, teamId, day) {
  if (!teamId) return false;
  const cd = normalizeConferenceDays(conferenceDays);
  return cd.teams[teamId]?.some((e) => e.day === day) ?? false;
}

export function isConferenceDayForSubGroup(conferenceDays, subGroupId, day) {
  if (!subGroupId) return false;
  const cd = normalizeConferenceDays(conferenceDays);
  return cd.subGroups[subGroupId]?.some((e) => e.day === day) ?? false;
}

export function isConferenceDayForWorker(conferenceDays, worker, day) {
  return (
    isConferenceDayForTeam(conferenceDays, worker.teamId, day) ||
    isConferenceDayForSubGroup(conferenceDays, worker.subGroupId, day)
  );
}

export function getConferenceTeamsOnDay(conferenceDays, teams, day) {
  const cd = normalizeConferenceDays(conferenceDays);
  return teams.filter((t) => cd.teams[t.id]?.some((e) => e.day === day));
}

export function getConferenceSubGroupsOnDay(conferenceDays, subGroups, day) {
  const cd = normalizeConferenceDays(conferenceDays);
  return subGroups.filter((sg) => cd.subGroups[sg.id]?.some((e) => e.day === day));
}

function randomInt(min, max, rng) {
  return Math.floor(rng() * (max - min + 1)) + min;
}

function mulberry32(a) {
  return function () {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffledRange(from, to, rng) {
  const arr = [];
  for (let i = from; i <= to; i++) arr.push(i);
  shuffleInPlace(arr, rng);
  return arr;
}

function shuffledArray(arr, rng) {
  const a = [...arr];
  shuffleInPlace(a, rng);
  return a;
}

function shuffleInPlace(arr, rng) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}

export function formatCellDisplay(cell, useShiftTypes) {
  if (!cell || cell.type === "off") return "休";
  if (cell.type === "half-off") return cell.half === "am" ? "前休" : "後休";
  if (useShiftTypes && cell.shiftType) return cell.shiftType;
  return "出";
}

export function isAttendingCell(cell) {
  return Boolean(cell && cell.type !== "off");
}
