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
  if (constraints.dailyMin > constraints.dailyMax) {
    return { ok: false, messages: ["全体の出勤下限が上限を超えています。"] };
  }
  if (workers.length < constraints.dailyMin) {
    return {
      ok: false,
      messages: [`勤務者数（${workers.length}人）が1日の出勤下限（${constraints.dailyMin}人）より少ないです。`],
    };
  }

  const lockedOff = buildLockedPreferences(workers, preferences, days);
  for (const w of workers) {
    const locked = countLockedOffDays(lockedOff, w.id, days);
    if (locked > (w.monthlyOffDays ?? 0)) {
      messages.push(
        `${w.name}: Excel希望休（${locked}日）が月間休み日数（${w.monthlyOffDays}日）を超えています。`
      );
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
        maxConsecutiveWork,
        constraints,
        teamConstraints,
        teams,
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
  const assignments = gridToAssignments(best, workers, days);
  const stats = buildStats(assignments, workers, days);

  return {
    ok: true,
    assignments,
    messages: messages.length ? messages : ["シフトを生成しました。"],
    stats,
    year,
    month,
    workers,
  };
}

function buildLockedPreferences(workers, preferences, days) {
  const locked = {};
  for (const w of workers) {
    locked[w.id] = {};
    const pref = preferences[w.name] || {};
    for (let d = 1; d <= days; d++) {
      if (pref[d]) locked[w.id][d] = true;
    }
  }
  return locked;
}

function countLockedOffDays(lockedOff, workerId, days) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    if (lockedOff[workerId]?.[d]) n++;
  }
  return n;
}

function countWorkerOffs(grid, workerId, days) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    if (!grid[workerId][d]) n++;
  }
  return n;
}

function tryBuildSchedule(ctx) {
  const {
    workers,
    days,
    lockedOff,
    maxConsecutiveWork,
    constraints,
    teamConstraints,
    teams,
    requireSupervisor,
    seed,
  } = ctx;
  const rng = mulberry32(seed);

  const grid = {};
  for (const w of workers) {
    grid[w.id] = {};
    for (let d = 1; d <= days; d++) {
      grid[w.id][d] = false;
    }
  }

  for (let d = 1; d <= days; d++) {
    if (!assignDay(grid, workers, d, days, constraints, teamConstraints, teams, lockedOff, requireSupervisor, rng)) {
      return null;
    }
  }

  const targets = Object.fromEntries(workers.map((w) => [w.id, w.monthlyOffDays ?? 0]));

  for (let pass = 0; pass < days * workers.length * 4; pass++) {
    let changed = false;
    for (const w of shuffledArray(workers, rng)) {
      const cur = countWorkerOffs(grid, w.id, days);
      const target = targets[w.id];
      if (cur < target) {
        if (tryAddOff(grid, w, days, lockedOff, maxConsecutiveWork, constraints, teamConstraints, teams, workers, requireSupervisor, rng)) {
          changed = true;
        }
      } else if (cur > target) {
        if (tryAddWork(grid, w, days, lockedOff, maxConsecutiveWork, constraints, teamConstraints, teams, workers, requireSupervisor, rng)) {
          changed = true;
        }
      }
    }
    if (!changed) break;
  }

  for (let iter = 0; iter < days * workers.length * 3; iter++) {
    let fixed = false;
    for (let d = 1; d <= days; d++) {
      if (validateDay(grid, workers, d, constraints, teamConstraints, teams, requireSupervisor).ok) {
        continue;
      }
      if (repairDay(grid, workers, d, days, lockedOff, maxConsecutiveWork, constraints, teamConstraints, teams, requireSupervisor, rng)) {
        fixed = true;
      }
    }
    if (!fixed) break;
  }

  for (const w of workers) {
    if (countWorkerOffs(grid, w.id, days) !== targets[w.id]) return null;
  }

  for (let d = 1; d <= days; d++) {
    if (!validateDay(grid, workers, d, constraints, teamConstraints, teams, requireSupervisor).ok) {
      return null;
    }
  }

  for (const w of workers) {
    if (violatesConsecutiveWork(grid[w.id], days, maxConsecutiveWork)) return null;
  }

  return grid;
}

function assignDay(grid, workers, day, days, constraints, teamConstraints, teams, lockedOff, requireSupervisor, rng) {
  const lockedOffToday = workers.filter((w) => lockedOff[w.id]?.[day]);

  for (let tryN = 0; tryN < 80; tryN++) {
    let target = randomInt(constraints.dailyMin, constraints.dailyMax, rng);
    target = Math.max(target, lockedOffToday.length);
    target = Math.min(target, workers.length);

    for (const w of workers) grid[w.id][day] = false;

    for (const w of lockedOffToday) grid[w.id][day] = false;

    const pool = workers.filter((w) => !lockedOff[w.id]?.[day]);
    const need = target - workers.filter((w) => grid[w.id][day]).length;
    if (need < 0) continue;

    const picked = pickWorkersForDay(pool, need, workers, grid, day, constraints, teamConstraints, teams, requireSupervisor, rng);
    if (!picked) continue;

    for (const w of picked) grid[w.id][day] = true;

    if (validateDay(grid, workers, day, constraints, teamConstraints, teams, requireSupervisor).ok) {
      return true;
    }
  }

  return assignDayGreedy(grid, workers, day, days, constraints, teamConstraints, teams, lockedOff, requireSupervisor, rng);
}

function assignDayGreedy(grid, workers, day, days, constraints, teamConstraints, teams, lockedOff, requireSupervisor, rng) {
  for (const w of workers) {
    if (!lockedOff[w.id]?.[day]) grid[w.id][day] = false;
  }
  for (const w of workers) {
    if (lockedOff[w.id]?.[day]) grid[w.id][day] = false;
  }

  const order = shuffledArray(workers.filter((w) => !lockedOff[w.id]?.[day]), rng);
  for (const w of order) {
    if (countWorking(grid, workers, day) >= constraints.dailyMax) break;
    grid[w.id][day] = true;
    const v = validateDay(grid, workers, day, constraints, teamConstraints, teams, requireSupervisor);
    if (!v.ok) grid[w.id][day] = false;
  }

  while (countWorking(grid, workers, day) < constraints.dailyMin) {
    let flipped = false;
    for (const w of shuffledArray(workers, rng)) {
      if (lockedOff[w.id]?.[day]) continue;
      if (grid[w.id][day]) continue;
      grid[w.id][day] = true;
      if (validateDay(grid, workers, day, constraints, teamConstraints, teams, requireSupervisor).ok) {
        flipped = true;
        break;
      }
      grid[w.id][day] = false;
    }
    if (!flipped) return false;
  }

  return validateDay(grid, workers, day, constraints, teamConstraints, teams, requireSupervisor).ok;
}

function pickWorkersForDay(pool, need, allWorkers, grid, day, constraints, teamConstraints, teams, requireSupervisor, rng) {
  if (need > pool.length) return null;
  const supervisors = pool.filter((w) => w.isSupervisor);
  const chosen = [];

  if (requireSupervisor && supervisors.length > 0 && need > 0) {
    chosen.push(supervisors[Math.floor(rng() * supervisors.length)]);
  }

  const rest = shuffledArray(
    pool.filter((w) => !chosen.includes(w)),
    rng
  );
  for (const w of rest) {
    if (chosen.length >= need) break;
    chosen.push(w);
  }

  return chosen.length === need ? chosen : null;
}

function countWorking(grid, workers, day) {
  return workers.filter((w) => grid[w.id][day]).length;
}

function tryAddOff(grid, w, days, lockedOff, maxConsecutive, constraints, teamConstraints, teams, workers, requireSupervisor, rng) {
  const candidates = shuffledRange(1, days, rng);
  for (const d of candidates) {
    if (lockedOff[w.id]?.[d]) continue;
    if (!grid[w.id][d]) continue;
    grid[w.id][d] = false;
    if (violatesConsecutiveWork(grid[w.id], days, maxConsecutive)) {
      grid[w.id][d] = true;
      continue;
    }
    if (!validateDay(grid, workers, d, constraints, teamConstraints, teams, requireSupervisor).ok) {
      grid[w.id][d] = true;
      continue;
    }
    return true;
  }
  return false;
}

function repairDay(grid, workers, day, days, lockedOff, maxConsecutive, constraints, teamConstraints, teams, requireSupervisor, rng) {
  for (const w of shuffledArray(workers, rng)) {
    if (tryAddWork(grid, w, days, lockedOff, maxConsecutive, constraints, teamConstraints, teams, workers, requireSupervisor, rng)) {
      if (validateDay(grid, workers, day, constraints, teamConstraints, teams, requireSupervisor).ok) return true;
    }
    if (tryAddOff(grid, w, days, lockedOff, maxConsecutive, constraints, teamConstraints, teams, workers, requireSupervisor, rng)) {
      if (validateDay(grid, workers, day, constraints, teamConstraints, teams, requireSupervisor).ok) return true;
    }
  }
  return false;
}

function tryAddWork(grid, w, days, lockedOff, maxConsecutive, constraints, teamConstraints, teams, workers, requireSupervisor, rng) {
  const candidates = shuffledRange(1, days, rng);
  for (const d of candidates) {
    if (lockedOff[w.id]?.[d]) continue;
    if (grid[w.id][d]) continue;
    grid[w.id][d] = true;
    if (violatesConsecutiveWork(grid[w.id], days, maxConsecutive)) {
      grid[w.id][d] = false;
      continue;
    }
    if (!validateDay(grid, workers, d, constraints, teamConstraints, teams, requireSupervisor).ok) {
      grid[w.id][d] = false;
      continue;
    }
    return true;
  }
  return false;
}

function validateDay(grid, workers, day, constraints, teamConstraints, teams, requireSupervisor) {
  const working = workers.filter((w) => grid[w.id][day]);
  const count = working.length;

  if (count < constraints.dailyMin || count > constraints.dailyMax) {
    return { ok: false };
  }

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

function gridToAssignments(grid, workers, days) {
  const assignments = {};
  for (const w of workers) {
    assignments[w.id] = {};
    for (let d = 1; d <= days; d++) {
      if (!grid[w.id][d]) {
        assignments[w.id][d] = { type: "off" };
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
      if (pref[d] && !grid[w.id][d]) score += 10;
      if (pref[d] && grid[w.id][d]) score -= 20;
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
    const working = workers.filter((w) => assignments[w.id][d]?.type === "work");
    daily.push({
      day: d,
      total: working.length,
      supervisors: working.filter((w) => w.isSupervisor).length,
    });
  }
  return { daily };
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
  if (useShiftTypes && cell.shiftType) return cell.shiftType;
  return "出";
}
