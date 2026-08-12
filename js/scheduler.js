import { getDaysInMonth } from "./store.js";

const MAX_ATTEMPTS = 500;
const MAX_AUTO_OFF_CONSECUTIVE = 3;
const PREFERRED_MAX_CONSECUTIVE = 3;
const CONFERENCE_SESSIONS_PER_MONTH = 4;

export const CONFERENCE_COLORS = [
  { bg: "#dbeafe", color: "#1e40af", border: "#93c5fd" },
  { bg: "#fce7f3", color: "#9d174d", border: "#f9a8d4" },
  { bg: "#fef3c7", color: "#92400e", border: "#fcd34d" },
  { bg: "#ffedd5", color: "#9a3412", border: "#fdba74" },
  { bg: "#ede9fe", color: "#5b21b6", border: "#c4b5fd" },
  { bg: "#e0e7ff", color: "#3730a3", border: "#a5b4fc" },
];

export const SUBGROUP_CONFERENCE_COLORS = [
  { bg: "#d1fae5", color: "#047857", border: "#34d399" },
  { bg: "#bbf7d0", color: "#166534", border: "#4ade80" },
  { bg: "#dcfce7", color: "#15803d", border: "#86efac" },
  { bg: "#ecfccb", color: "#3f6212", border: "#bef264" },
];

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
    const achievable = validateOffTargetAchievable(w, days, lockedOff, halfOff, w.monthlyOffDays ?? 0);
    if (!achievable.ok) {
      messages.push(achievable.message);
      return { ok: false, messages };
    }
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
    if (isOffTargetFixedByPreferences(w, days, lockedOff, halfOff, target)) {
      const row = buildFixedWorkerRow(w, days, lockedOff, lockedWork, halfOff);
      if (violatesWorkerStreaks(row, days, lockedOff, w.id, maxConsecutiveWork)) {
        messages.push(
          `${w.name}: Excel希望休が月間休み日数（${formatOffDays(target)}日）と一致するため残りは全出勤になりますが、連勤上限（${maxConsecutiveWork}日）を超えます。休み希望の日を分散するか、月間休み日数・連勤上限を見直してください。`
        );
        return {
          ok: false,
          messages,
        };
      }
    }
  }

  let best = null;
  let bestScore = -Infinity;
  let bestAbsence = Infinity;
  let bestDeviation = Infinity;

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
      const groupCtx = { constraints, teamConstraints, teams, subGroups, subGroupConstraints };
      const absence = countSupervisorAbsenceDays(grid, workers, days, groupCtx);
      const deviation = countGroupHeadcountDeviationDays(grid, workers, days, groupCtx);
      const score = scoreSchedule(
        grid,
        workers,
        days,
        preferences,
        requireSupervisor,
        lockedOff,
        maxConsecutiveWork,
        groupCtx
      );
      if (
        absence < bestAbsence ||
        (absence === bestAbsence && deviation < bestDeviation) ||
        (absence === bestAbsence && deviation === bestDeviation && score > bestScore)
      ) {
        bestAbsence = absence;
        bestDeviation = deviation;
        bestScore = score;
        best = grid;
      }
    }
    if (best) break;
  }

  if (!best) {
    const hints = collectScheduleHints(workers, days, lockedOff, lockedWork, halfOff, constraints, maxConsecutiveWork);
    const supervisorCount = workers.filter((w) => w.isSupervisor).length;
    if (supervisorCount > constraints.supervisorMax) {
      hints.unshift(
        `全体責任者の登録${supervisorCount}名に対し、1日あたりの上限が${constraints.supervisorMax}名です。上限を${supervisorCount}名以上にするか、全体責任者登録を減らしてください。`
      );
    }
    for (const team of teams) {
      const tc = teamConstraints[team.id];
      if (!tc) continue;
      const groupSupCount = workers.filter((w) => w.isGroupSupervisor && w.teamId === team.id).length;
      if (groupSupCount > (tc.supervisorMax ?? 99)) {
        hints.unshift(
          `${team.name}: グループ責任者の登録${groupSupCount}名に対し、1日あたりの上限が${tc.supervisorMax}名です。`
        );
      }
    }
    return {
      ok: false,
      messages: [
        ...messages,
        ...hints,
        "制約を満たすシフトを生成できませんでした。人数制約・休み日数・連勤上限・Excel希望を見直してください。",
      ],
    };
  }

  applyShiftTypes(best, workers, days, useShiftTypes, shiftTypes);
  const assignments = gridToAssignments(best, workers, days, halfOff, lockedOff);
  const stats = buildStats(assignments, workers, days);
  const supervisorAbsence = buildSupervisorAbsence(
    assignments,
    workers,
    days,
    constraints,
    teams,
    teamConstraints
  );
  const absenceDays = Object.keys(supervisorAbsence).map((d) => Number(d)).sort((a, b) => a - b);
  if (absenceDays.length) {
    messages.push(
      `希望休の重複などにより責任者の下限を満たせない日が${absenceDays.length}日あります（${absenceDays.join("日、")}日）。候補のうち責任者不在が最も少ないシフトを表示しています。`
    );
  }
  const headcountDeviation = buildGroupHeadcountDeviation(
    assignments,
    workers,
    days,
    teams,
    teamConstraints,
    subGroups,
    subGroupConstraints
  );
  const deviationDays = Object.keys(headcountDeviation).map((d) => Number(d)).sort((a, b) => a - b);
  if (deviationDays.length) {
    messages.push(
      `グループの出勤人数の下限・上限を満たせない日が${deviationDays.length}日あります（${deviationDays.join("日、")}日）。候補のうち人数逸脱が最も少ないシフトを表示しています。`
    );
  }
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
    messages: messages.length ? ["シフトを生成しました。", ...messages] : ["シフトを生成しました。"],
    stats,
    supervisorAbsence,
    headcountDeviation,
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
      const v = pref[d] ?? pref[String(d)];
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
    const v = pref?.[d] ?? pref?.[String(d)];
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

function canAssignOffOnDay(w, d, lockedOff, lockedWork, halfOff) {
  if (lockedOff[w.id]?.[d]) return false;
  if (lockedWork[w.id]?.[d]) return false;
  if (halfOff[w.id]?.[d]) return false;
  return true;
}

function findLongestWorkStreak(row, days) {
  let best = { start: 1, end: 0, len: 0 };
  let start = 1;
  let streak = 0;
  for (let d = 1; d <= days + 1; d++) {
    if (d <= days && row[d]) {
      if (streak === 0) start = d;
      streak++;
    } else if (streak > best.len) {
      best = { start, end: d - 1, len: streak };
      streak = 0;
    } else {
      streak = 0;
    }
  }
  return best;
}

function findAllWorkStreaks(row, days) {
  const streaks = [];
  let start = 1;
  let streak = 0;
  for (let d = 1; d <= days + 1; d++) {
    if (d <= days && row[d]) {
      if (streak === 0) start = d;
      streak++;
    } else {
      if (streak > 0) streaks.push({ start, end: d - 1, len: streak });
      streak = 0;
    }
  }
  return streaks;
}

function idealWorkStreakMax(maxConsecutive) {
  return Math.max(1, Math.min(PREFERRED_MAX_CONSECUTIVE, maxConsecutive - 1));
}

function isAdjacentToLockedOff(d, lockedOff, workerId) {
  return Boolean(lockedOff[workerId]?.[d - 1] || lockedOff[workerId]?.[d + 1]);
}

function scoreWorkStreaks(row, days, maxConsecutive) {
  const idealMax = idealWorkStreakMax(maxConsecutive);
  let score = 0;
  for (const streak of findAllWorkStreaks(row, days)) {
    if (streak.len < maxConsecutive) {
      score += (maxConsecutive - streak.len) * 10;
    }
    if (streak.len <= idealMax) {
      score += (idealMax + 1 - streak.len) * 8;
    } else if (streak.len >= 4) {
      score -= (streak.len - 3) * 30;
    } else {
      score -= (streak.len - idealMax) * 12;
    }
  }
  return score;
}

function scoreAutoOffPlacement(row, days, lockedOff, workerId) {
  let score = 0;
  for (let d = 1; d <= days; d++) {
    if (row[d] || lockedOff[workerId]?.[d]) continue;
    if (isAdjacentToLockedOff(d, lockedOff, workerId)) score -= 20;
  }
  return score;
}

function daysAroundMid(start, end) {
  const mid = Math.floor((start + end) / 2);
  const out = [];
  for (let offset = 0; offset <= end - start; offset++) {
    if (offset % 2 === 0) {
      const d = mid + offset / 2;
      if (d >= start && d <= end) out.push(d);
      const d2 = mid - (offset / 2 + 1);
      if (d2 >= start && d2 <= end) out.push(d2);
    }
  }
  return out;
}

function offCandidateDays(row, days, w, lockedOff, lockedWork, halfOff, maxConsecutive) {
  const idealMax = idealWorkStreakMax(maxConsecutive);
  const candidates = [];
  for (let d = 1; d <= days; d++) {
    if (!canAssignOffOnDay(w, d, lockedOff, lockedWork, halfOff) || !row[d]) continue;
    const prev = row[d];
    row[d] = false;
    const autoOffOk = !violatesConsecutiveAutoOff(row, days, lockedOff, w.id);
    row[d] = prev;
    if (!autoOffOk) continue;
    let left = 0;
    let right = 0;
    for (let i = d - 1; i >= 1 && row[i]; i--) left++;
    for (let i = d + 1; i <= days && row[i]; i++) right++;
    const runLen = left + right + 1;
    let score = 0;
    if (runLen > idealMax) score = runLen * 8;
    else score = runLen - (idealMax + 1);
    if (isAdjacentToLockedOff(d, lockedOff, w.id)) score -= 60;
    candidates.push({ d, score });
  }
  candidates.sort((a, b) => b.score - a.score);
  return candidates.map((c) => c.d);
}

function workCandidateDays(row, days, w, lockedOff, maxConsecutive) {
  const idealMax = idealWorkStreakMax(maxConsecutive);
  const candidates = [];
  for (let d = 1; d <= days; d++) {
    if (lockedOff[w.id]?.[d] || row[d]) continue;
    const prev = row[d];
    row[d] = true;
    const streak = findLongestWorkStreak(row, days);
    row[d] = prev;
    let score = 0;
    if (streak.len > idealMax) score -= streak.len * 8;
    else score += idealMax + 1 - streak.len;
    if (streak.len >= 4) score -= 40;
    candidates.push({ d, score });
  }
  candidates.sort((a, b) => b.score - a.score);
  return candidates.map((c) => c.d);
}

function applyAllPreferenceLocks(grid, workers, days, lockedOff, lockedWork, halfOff) {
  for (const w of workers) {
    applyPreferenceLocksToGrid(grid, w, days, lockedOff, lockedWork, halfOff);
  }
}

function balanceOffTargets(grid, workers, days, targets, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, requireSupervisor, rng, maxPasses) {
  for (let pass = 0; pass < maxPasses; pass++) {
    let changed = false;
    for (const w of shuffledArray(workers, rng)) {
      const cur = countEffectiveOffDays(grid, w.id, days, halfOff);
      const target = targets[w.id];
      if (cur < target - 0.001) {
        if (
          tryAddOff(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng) ||
          trySwapForOff(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng)
        ) {
          changed = true;
        }
      } else if (cur > target + 0.001) {
        if (
          tryAddWork(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng) ||
          trySwapForWork(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng)
        ) {
          changed = true;
        }
      }
    }
    if (!changed) break;
  }
}

function repairAllDays(grid, workers, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, requireSupervisor, rng, maxPasses) {
  for (let pass = 0; pass < maxPasses; pass++) {
    let fixed = false;
    for (let d = 1; d <= days; d++) {
      if (validateDay(grid, workers, d, groupCtx, requireSupervisor).ok) continue;
      if (repairDay(grid, workers, d, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, requireSupervisor, rng)) {
        fixed = true;
      }
    }
    if (!fixed) break;
  }
}

function applyPreferenceLocksToGrid(grid, w, days, lockedOff, lockedWork, halfOff) {
  for (let d = 1; d <= days; d++) {
    if (lockedOff[w.id]?.[d]) grid[w.id][d] = false;
    else if (lockedWork[w.id]?.[d] || halfOff[w.id]?.[d]) grid[w.id][d] = true;
    else grid[w.id][d] = true;
  }
}

function seedWorkerOffDays(grid, w, days, targetOff, lockedOff, lockedWork, halfOff, maxConsecutive, rng) {
  applyPreferenceLocksToGrid(grid, w, days, lockedOff, lockedWork, halfOff);

  let cur = countEffectiveOffDays(grid, w.id, days, halfOff);
  let needed = targetOff - cur;
  if (needed > 0.001) {
    const available = offCandidateDays(grid[w.id], days, w, lockedOff, lockedWork, halfOff, maxConsecutive);
    const count = Math.min(Math.ceil(needed - 0.001), available.length);
    if (count > 0) {
      const shuffled = shuffledArray(available, rng);
      const step = shuffled.length / count;
      for (let i = 0; i < count; i++) {
        const d = shuffled[Math.min(shuffled.length - 1, Math.floor(i * step + step / 2))];
        grid[w.id][d] = false;
      }
    }
  }

  let guard = 0;
  while (violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) && guard < days * 2) {
    guard++;
    const streak = findLongestWorkStreak(grid[w.id], days);
    if (streak.len <= maxConsecutive) break;
    let placed = false;
    for (const d of daysAroundMid(streak.start, streak.end)) {
      if (!canAssignOffOnDay(w, d, lockedOff, lockedWork, halfOff) || !grid[w.id][d]) continue;
      grid[w.id][d] = false;
      placed = true;
      break;
    }
    if (!placed) break;
  }

  cur = countEffectiveOffDays(grid, w.id, days, halfOff);
  guard = 0;
  while (cur > targetOff + 0.001 && guard < days * 2) {
    guard++;
    const preferred = workCandidateDays(grid[w.id], days, w, lockedOff, maxConsecutive).filter(
      (d) => canAssignOffOnDay(w, d, lockedOff, lockedWork, halfOff) && !grid[w.id][d]
    );
    const candidates = preferred.length
      ? preferred
      : shuffledRange(1, days, rng).filter(
          (d) => canAssignOffOnDay(w, d, lockedOff, lockedWork, halfOff) && !grid[w.id][d]
        );
    let removed = false;
    for (const d of candidates) {
      grid[w.id][d] = true;
      if (!violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive)) {
        removed = true;
        break;
      }
      grid[w.id][d] = false;
    }
    if (!removed) break;
    cur = countEffectiveOffDays(grid, w.id, days, halfOff);
  }
}

function spreadAdditionalOffDays(grid, w, days, neededOff, lockedOff, lockedWork, halfOff, maxConsecutive, rng) {
  if (neededOff <= 0.001) return;
  const count = Math.min(Math.ceil(neededOff - 0.001), days);
  const available = offCandidateDays(grid[w.id], days, w, lockedOff, lockedWork, halfOff, maxConsecutive);
  if (!available.length) return;

  const shuffled = shuffledArray(available, rng);
  const step = shuffled.length / count;
  for (let i = 0; i < count; i++) {
    const d = shuffled[Math.min(shuffled.length - 1, Math.floor(i * step + step / 2))];
    grid[w.id][d] = false;
  }

  let guard = 0;
  while (violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) && guard < days * 2) {
    guard++;
    const streak = findLongestWorkStreak(grid[w.id], days);
    if (streak.len <= maxConsecutive) break;
    let placed = false;
    for (const d of daysAroundMid(streak.start, streak.end)) {
      if (!canAssignOffOnDay(w, d, lockedOff, lockedWork, halfOff) || !grid[w.id][d]) continue;
      grid[w.id][d] = false;
      placed = true;
      break;
    }
    if (!placed) break;
  }
}

function trySwapForOff(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng) {
  const preferred = offCandidateDays(grid[w.id], days, w, lockedOff, lockedWork, halfOff, maxConsecutive).filter((d) => grid[w.id][d]);
  const dayCandidates = preferred.length ? preferred.slice(0, 12) : shuffledRange(1, days, rng).slice(0, 12);
  for (const d of dayCandidates) {
    for (const w2 of shuffledArray(workers.filter((x) => x.id !== w.id), rng).slice(0, 8)) {
      if (grid[w2.id][d] || !canAssignOffOnDay(w2, d, lockedOff, lockedWork, halfOff)) continue;
      const beforeSnap = groupHeadcountSnapshot(grid, workers, d, groupCtx);
      grid[w.id][d] = false;
      grid[w2.id][d] = true;
      if (
        !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
        !violatesWorkerStreaks(grid[w2.id], days, lockedOff, w2.id, maxConsecutive) &&
        validateDay(grid, workers, d, groupCtx, requireSupervisor).ok &&
        !groupHeadcountWorsenedOnDay(grid, workers, d, groupCtx, beforeSnap)
      ) {
        return true;
      }
      grid[w.id][d] = true;
      grid[w2.id][d] = false;
    }
  }
  return false;
}

function trySwapForWork(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng) {
  const dayCandidates = shuffledRange(1, days, rng).slice(0, 12);
  for (const d of dayCandidates) {
    if (grid[w.id][d] || lockedOff[w.id]?.[d] || lockedWork[w.id]?.[d] || halfOff[w.id]?.[d]) continue;
    for (const w2 of shuffledArray(workers.filter((x) => x.id !== w.id), rng).slice(0, 8)) {
      if (!grid[w2.id][d] || !canAssignOffOnDay(w2, d, lockedOff, lockedWork, halfOff)) continue;
      const beforeSnap = groupHeadcountSnapshot(grid, workers, d, groupCtx);
      grid[w.id][d] = true;
      grid[w2.id][d] = false;
      if (
        !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
        !violatesWorkerStreaks(grid[w2.id], days, lockedOff, w2.id, maxConsecutive) &&
        validateDay(grid, workers, d, groupCtx, requireSupervisor).ok &&
        !groupHeadcountWorsenedOnDay(grid, workers, d, groupCtx, beforeSnap)
      ) {
        return true;
      }
      grid[w.id][d] = false;
      grid[w2.id][d] = true;
    }
  }
  return false;
}

function countLockedEffectiveOffDays(lockedOff, halfOff, workerId, days) {
  return countLockedOffDays(lockedOff, workerId, days) + countHalfOffDays(halfOff, workerId, days) * 0.5;
}

function validateOffTargetAchievable(w, days, lockedOff, halfOff, target) {
  const locked = countLockedEffectiveOffDays(lockedOff, halfOff, w.id, days);
  if (target < locked - 0.001) {
    return {
      ok: false,
      message: `${w.name}: Excel希望休（${formatOffDays(locked)}日）が月間休み日数（${formatOffDays(target)}日）を超えています。`,
    };
  }
  const extra = target - locked;
  if (Math.abs(extra - Math.round(extra)) > 0.001) {
    const lower = locked + Math.floor(extra);
    const upper = locked + Math.ceil(extra);
    const halfCount = countHalfOffDays(halfOff, w.id, days);
    return {
      ok: false,
      message:
        `${w.name}: Excelの半休${halfCount > 0 ? `（${halfCount}件）` : ""}により希望休が${formatOffDays(locked)}日です。` +
        `月間休み日数${formatOffDays(target)}日は、残りを1日単位でしか追加できないため一致しません。` +
        `半休を「休」に変えるか、月間休み日数を${formatOffDays(lower)}日または${formatOffDays(upper)}日に変更してください。`,
    };
  }
  return { ok: true };
}

function buildFixedWorkerRow(w, days, lockedOff, lockedWork, halfOff) {
  void lockedWork;
  const row = {};
  for (let d = 1; d <= days; d++) {
    if (lockedOff[w.id]?.[d]) row[d] = false;
    else if (halfOff[w.id]?.[d]) row[d] = true;
    else row[d] = true;
  }
  return row;
}

function isOffTargetFixedByPreferences(w, days, lockedOff, halfOff, targetOff) {
  const lockedOffCount = countLockedOffDays(lockedOff, w.id, days);
  const halfOffCount = countHalfOffDays(halfOff, w.id, days);
  return Math.abs(targetOff - (lockedOffCount + halfOffCount * 0.5)) < 0.001;
}

function collectScheduleHints(workers, days, lockedOff, lockedWork, halfOff, constraints, maxConsecutiveWork) {
  void maxConsecutiveWork;
  const hints = [];
  for (const w of workers) {
    const target = w.monthlyOffDays ?? 0;
    const prefOff =
      countLockedOffDays(lockedOff, w.id, days) + countHalfOffDays(halfOff, w.id, days) * 0.5;
    if (prefOff > target + 0.001) {
      hints.push(`${w.name}: Excel希望休（${formatOffDays(prefOff)}日）が月間休み日数（${target}日）を超えています。`);
      continue;
    }
    if (isOffTargetFixedByPreferences(w, days, lockedOff, halfOff, target)) {
      const row = buildFixedWorkerRow(w, days, lockedOff, lockedWork, halfOff);
      if (violatesWorkerStreaks(row, days, lockedOff, w.id, maxConsecutiveWork)) {
        hints.push(
          `${w.name}: Excel希望休が月間休み日数（${formatOffDays(target)}日）と一致するため残りは全出勤になりますが、連勤上限（${maxConsecutiveWork}日）を超えます。`
        );
      }
    }
  }
  const lockedPrefCount = workers.filter((w) => countLockedOffDays(lockedOff, w.id, days) > 0).length;
  if (lockedPrefCount > 0) {
    hints.push(
      `Excel休み希望が設定されている勤務者が${lockedPrefCount}名います。休み希望は優先されますが、月間休み日数との調整が難しい場合があります。`
    );
  }
  return hints;
}

export function countWorkerOffDaysFromAssignments(assignments, workerId, days) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    const cell = assignments[workerId]?.[d];
    if (!cell || cell.type === "off") n += 1;
    else if (cell.type === "half-off") n += 0.5;
  }
  return n;
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
  const groupCtx = {
    constraints,
    teamConstraints,
    teams,
    subGroups,
    subGroupConstraints,
    lockedOff,
    lockedWork,
    halfOff,
  };
  const rng = mulberry32(seed);

  const grid = {};
  for (const w of workers) {
    grid[w.id] = {};
    for (let d = 1; d <= days; d++) {
      grid[w.id][d] = false;
    }
  }

  const targets = Object.fromEntries(workers.map((w) => [w.id, w.monthlyOffDays ?? 0]));
  applyAllPreferenceLocks(grid, workers, days, lockedOff, lockedWork, halfOff);

  const strategy = seed % 3;
  if (strategy === 0) {
    for (let d = 1; d <= days; d++) {
      if (!assignDay(grid, workers, d, groupCtx, lockedOff, lockedWork, halfOff, requireSupervisor, rng)) {
        return null;
      }
    }
  } else if (strategy === 1) {
    for (const w of shuffledArray(workers, rng)) {
      seedWorkerOffDays(grid, w, days, targets[w.id], lockedOff, lockedWork, halfOff, maxConsecutiveWork, rng);
    }
  }

  repairAllDays(
    grid,
    workers,
    days,
    lockedOff,
    lockedWork,
    halfOff,
    maxConsecutiveWork,
    groupCtx,
    requireSupervisor,
    rng,
    days * 8
  );

  balanceOffTargets(
    grid,
    workers,
    days,
    targets,
    lockedOff,
    lockedWork,
    halfOff,
    maxConsecutiveWork,
    groupCtx,
    requireSupervisor,
    rng,
    days * workers.length * 16
  );

  fillSupervisorGaps(
    grid,
    workers,
    days,
    lockedOff,
    lockedWork,
    halfOff,
    maxConsecutiveWork,
    groupCtx,
    requireSupervisor,
    rng
  );

  balanceOffTargets(
    grid,
    workers,
    days,
    targets,
    lockedOff,
    lockedWork,
    halfOff,
    maxConsecutiveWork,
    groupCtx,
    requireSupervisor,
    rng,
    days * workers.length * 8
  );

  fillSupervisorGaps(
    grid,
    workers,
    days,
    lockedOff,
    lockedWork,
    halfOff,
    maxConsecutiveWork,
    groupCtx,
    requireSupervisor,
    rng
  );

  fillGroupHeadcountGaps(
    grid,
    workers,
    days,
    lockedOff,
    lockedWork,
    halfOff,
    maxConsecutiveWork,
    groupCtx,
    requireSupervisor,
    rng
  );

  repairAllDays(
    grid,
    workers,
    days,
    lockedOff,
    lockedWork,
    halfOff,
    maxConsecutiveWork,
    groupCtx,
    requireSupervisor,
    rng,
    days * workers.length * 3
  );

  fillGroupHeadcountGaps(
    grid,
    workers,
    days,
    lockedOff,
    lockedWork,
    halfOff,
    maxConsecutiveWork,
    groupCtx,
    requireSupervisor,
    rng
  );

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
    if (violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutiveWork)) return null;
  }

  return grid;
}

function assignDay(grid, workers, day, groupCtx, lockedOff, lockedWork, halfOff, requireSupervisor, rng) {
  for (const w of workers) {
    if (lockedOff[w.id]?.[day]) grid[w.id][day] = false;
    else if (lockedWork[w.id]?.[day] || halfOff[w.id]?.[day]) grid[w.id][day] = true;
    else grid[w.id][day] = true;
  }
  if (validateDay(grid, workers, day, groupCtx, requireSupervisor).ok) {
    fitDayGroupHeadcount(grid, workers, day, groupCtx, lockedOff, lockedWork, halfOff, requireSupervisor, rng);
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
        fitDayGroupHeadcount(grid, workers, day, groupCtx, lockedOff, lockedWork, halfOff, requireSupervisor, rng);
        return true;
      }
      grid[w.id][day] = true;
    }

    const offPool = shuffledArray(
      workers.filter(
        (w) => !grid[w.id][day] && !lockedOff[w.id]?.[day] && !lockedWork[w.id]?.[day] && !halfOff[w.id]?.[day]
      ),
      rng
    );
    for (const w of offPool) {
      grid[w.id][day] = true;
      if (validateDay(grid, workers, day, groupCtx, requireSupervisor).ok) {
        fitDayGroupHeadcount(grid, workers, day, groupCtx, lockedOff, lockedWork, halfOff, requireSupervisor, rng);
        return true;
      }
      grid[w.id][day] = false;
    }
  }

  const ok = validateDay(grid, workers, day, groupCtx, requireSupervisor).ok;
  if (ok) {
    fitDayGroupHeadcount(grid, workers, day, groupCtx, lockedOff, lockedWork, halfOff, requireSupervisor, rng);
  }
  return ok;
}

function tryAddOff(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng) {
  return (
    tryAddOffOnDays(
      grid,
      w,
      days,
      lockedOff,
      lockedWork,
      halfOff,
      maxConsecutive,
      groupCtx,
      workers,
      requireSupervisor,
      rng,
      false,
      true
    ) ||
    tryAddOffOnDays(
      grid,
      w,
      days,
      lockedOff,
      lockedWork,
      halfOff,
      maxConsecutive,
      groupCtx,
      workers,
      requireSupervisor,
      rng,
      false,
      false
    )
  );
}

function tryAddOffKeepingSupervisors(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng) {
  return (
    tryAddOffOnDays(
      grid,
      w,
      days,
      lockedOff,
      lockedWork,
      halfOff,
      maxConsecutive,
      groupCtx,
      workers,
      requireSupervisor,
      rng,
      true,
      true
    ) ||
    tryAddOffOnDays(
      grid,
      w,
      days,
      lockedOff,
      lockedWork,
      halfOff,
      maxConsecutive,
      groupCtx,
      workers,
      requireSupervisor,
      rng,
      true,
      false
    )
  );
}

function tryAddOffOnDays(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng, keepSupervisors, keepGroupBounds = false) {
  const candidates = offCandidateDays(grid[w.id], days, w, lockedOff, lockedWork, halfOff, maxConsecutive);
  const nonAdjacent = candidates.filter((d) => !isAdjacentToLockedOff(d, lockedOff, w.id));
  const shuffled = (nonAdjacent.length ? nonAdjacent : candidates.length ? candidates : shuffledRange(1, days, rng));
  for (const d of shuffled) {
    if (lockedOff[w.id]?.[d]) continue;
    if (lockedWork[w.id]?.[d]) continue;
    if (halfOff[w.id]?.[d]) continue;
    if (!grid[w.id][d]) continue;
    const beforeSnap = keepGroupBounds ? groupHeadcountSnapshot(grid, workers, d, groupCtx) : null;
    grid[w.id][d] = false;
    if (violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive)) {
      grid[w.id][d] = true;
      continue;
    }
    if (!validateDay(grid, workers, d, groupCtx, requireSupervisor).ok) {
      grid[w.id][d] = true;
      continue;
    }
    if (keepSupervisors && isSupervisorShortOnDay(grid, workers, d, groupCtx)) {
      grid[w.id][d] = true;
      continue;
    }
    if (keepGroupBounds && groupHeadcountWorsenedOnDay(grid, workers, d, groupCtx, beforeSnap)) {
      grid[w.id][d] = true;
      continue;
    }
    return true;
  }
  return false;
}

function isSupervisorShortOnDay(grid, workers, day, groupCtx) {
  const { constraints, teams = [], teamConstraints = {} } = groupCtx;
  const working = workers.filter((w) => grid[w.id][day]);
  if (working.filter((w) => w.isSupervisor).length < (constraints?.supervisorMin ?? 0)) {
    return true;
  }
  for (const team of teams) {
    const min = teamConstraints[team.id]?.supervisorMin ?? 0;
    if (min <= 0) continue;
    const n = working.filter((w) => w.isGroupSupervisor && w.teamId === team.id).length;
    if (n < min) return true;
  }
  return false;
}

function countSupervisorAbsenceDays(grid, workers, days, groupCtx) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    if (isSupervisorShortOnDay(grid, workers, d, groupCtx)) n++;
  }
  return n;
}

function fillSupervisorGaps(grid, workers, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, requireSupervisor, rng) {
  const { constraints, teams = [], teamConstraints = {} } = groupCtx;
  for (let pass = 0; pass < days * 3; pass++) {
    let changed = false;
    for (let d = 1; d <= days; d++) {
      if (!isSupervisorShortOnDay(grid, workers, d, groupCtx)) continue;
      const working = workers.filter((w) => grid[w.id][d]);
      if (working.filter((w) => w.isSupervisor).length < (constraints?.supervisorMin ?? 0)) {
        const placed = tryPlaceSupervisor(
          grid,
          workers,
          d,
          days,
          lockedOff,
          lockedWork,
          halfOff,
          maxConsecutive,
          groupCtx,
          requireSupervisor,
          rng,
          (w) => w.isSupervisor
        );
        if (placed) {
          const restored = tryAddOffKeepingSupervisors(
            grid,
            placed,
            days,
            lockedOff,
            lockedWork,
            halfOff,
            maxConsecutive,
            groupCtx,
            workers,
            requireSupervisor,
            rng
          );
          if (restored) changed = true;
          else grid[placed.id][d] = false;
        }
      }
      for (const team of teams) {
        const min = teamConstraints[team.id]?.supervisorMin ?? 0;
        if (min <= 0) continue;
        const n = workers.filter((w) => w.isGroupSupervisor && w.teamId === team.id && grid[w.id][d]).length;
        if (n >= min) continue;
        const placed = tryPlaceSupervisor(
          grid,
          workers,
          d,
          days,
          lockedOff,
          lockedWork,
          halfOff,
          maxConsecutive,
          groupCtx,
          requireSupervisor,
          rng,
          (w) => w.isGroupSupervisor && w.teamId === team.id
        );
        if (placed) {
          const restored = tryAddOffKeepingSupervisors(
            grid,
            placed,
            days,
            lockedOff,
            lockedWork,
            halfOff,
            maxConsecutive,
            groupCtx,
            workers,
            requireSupervisor,
            rng
          );
          if (restored) changed = true;
          else grid[placed.id][d] = false;
        }
      }
    }
    if (!changed) break;
  }
}

function listActiveGroupBounds(groupCtx) {
  const { teams = [], teamConstraints = {}, subGroups = [], subGroupConstraints = {} } = groupCtx;
  const bounds = [];
  for (const team of teams) {
    const tc = teamConstraints[team.id];
    if (!tc) continue;
    const min = Number(tc.min) || 0;
    const max = tc.max == null ? 99 : Number(tc.max);
    if (min <= 0 && max >= 99) continue;
    bounds.push({
      key: `team:${team.id}`,
      name: team.name,
      min,
      max,
      match: (w) => w.teamId === team.id,
    });
  }
  for (const sg of subGroups) {
    const sgc = subGroupConstraints[sg.id];
    if (!sgc) continue;
    const min = Number(sgc.min) || 0;
    const max = sgc.max == null ? 99 : Number(sgc.max);
    if (min <= 0 && max >= 99) continue;
    bounds.push({
      key: `sg:${sg.id}`,
      name: sg.name,
      min,
      max,
      match: (w) => w.subGroupId === sg.id,
    });
  }
  return bounds;
}

function groupHeadcountSnapshotFromWorking(working, groupCtx) {
  const issues = [];
  let magnitude = 0;
  for (const b of listActiveGroupBounds(groupCtx)) {
    const count = working.filter(b.match).length;
    if (count < b.min) {
      const gap = b.min - count;
      magnitude += gap;
      issues.push({ ...b, count, kind: "under", gap });
    } else if (count > b.max) {
      const gap = count - b.max;
      magnitude += gap;
      issues.push({ ...b, count, kind: "over", gap });
    }
  }
  return { issues, magnitude, deviant: issues.length > 0 };
}

function groupHeadcountSnapshot(grid, workers, day, groupCtx) {
  const working = workers.filter((w) => grid[w.id][day]);
  return groupHeadcountSnapshotFromWorking(working, groupCtx);
}

function groupHeadcountWorsenedOnDay(grid, workers, day, groupCtx, beforeSnap) {
  if (!beforeSnap) return false;
  const after = groupHeadcountSnapshot(grid, workers, day, groupCtx);
  if (!beforeSnap.deviant && after.deviant) return true;
  return after.magnitude > beforeSnap.magnitude;
}

function countGroupHeadcountDeviationDays(grid, workers, days, groupCtx) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    if (groupHeadcountSnapshot(grid, workers, d, groupCtx).deviant) n++;
  }
  return n;
}

function countGroupHeadcountDeviationMagnitude(grid, workers, days, groupCtx) {
  let n = 0;
  for (let d = 1; d <= days; d++) {
    n += groupHeadcountSnapshot(grid, workers, d, groupCtx).magnitude;
  }
  return n;
}

function isGroupHeadcountImproved(grid, workers, days, groupCtx, beforeDays, beforeMag, beforeAbs) {
  const afterAbs = countSupervisorAbsenceDays(grid, workers, days, groupCtx);
  if (afterAbs > beforeAbs) return false;
  const afterDays = countGroupHeadcountDeviationDays(grid, workers, days, groupCtx);
  if (afterDays < beforeDays) return true;
  if (afterDays > beforeDays) return false;
  return countGroupHeadcountDeviationMagnitude(grid, workers, days, groupCtx) < beforeMag;
}

function fitDayGroupHeadcount(grid, workers, day, groupCtx, lockedOff, lockedWork, halfOff, requireSupervisor, rng) {
  for (let attempt = 0; attempt < workers.length * 6; attempt++) {
    const snap = groupHeadcountSnapshot(grid, workers, day, groupCtx);
    if (!snap.deviant) return;
    let moved = false;
    for (const issue of snap.issues) {
      if (issue.kind === "over") {
        const pool = shuffledArray(
          workers.filter(
            (w) =>
              issue.match(w) &&
              grid[w.id][day] &&
              !lockedOff[w.id]?.[day] &&
              !lockedWork[w.id]?.[day] &&
              !halfOff[w.id]?.[day]
          ),
          rng
        );
        for (const w of pool) {
          grid[w.id][day] = false;
          if (
            validateDay(grid, workers, day, groupCtx, requireSupervisor).ok &&
            groupHeadcountSnapshot(grid, workers, day, groupCtx).magnitude < snap.magnitude
          ) {
            moved = true;
            break;
          }
          grid[w.id][day] = true;
        }
      } else {
        const pool = shuffledArray(
          workers.filter((w) => issue.match(w) && !grid[w.id][day] && !lockedOff[w.id]?.[day]),
          rng
        );
        for (const w of pool) {
          grid[w.id][day] = true;
          if (
            validateDay(grid, workers, day, groupCtx, requireSupervisor).ok &&
            groupHeadcountSnapshot(grid, workers, day, groupCtx).magnitude < snap.magnitude
          ) {
            moved = true;
            break;
          }
          grid[w.id][day] = false;
        }
      }
      if (moved) break;
    }
    if (!moved) return;
  }
}

function fillGroupHeadcountGaps(grid, workers, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, requireSupervisor, rng) {
  const bounds = listActiveGroupBounds(groupCtx);
  if (!bounds.length) return;

  for (let pass = 0; pass < days * 4; pass++) {
    let changed = false;

    for (const b of bounds) {
      const overDays = [];
      const underDays = [];
      for (let d = 1; d <= days; d++) {
        const count = workers.filter((w) => b.match(w) && grid[w.id][d]).length;
        if (count > b.max) overDays.push(d);
        if (count < b.min) underDays.push(d);
      }
      outer: for (const from of overDays) {
        for (const to of underDays) {
          const movers = shuffledArray(
            workers.filter(
              (w) =>
                b.match(w) &&
                grid[w.id][from] &&
                !grid[w.id][to] &&
                !lockedWork[w.id]?.[from] &&
                !halfOff[w.id]?.[from] &&
                !lockedOff[w.id]?.[to]
            ),
            rng
          );
          for (const w of movers) {
            const backup = { ...grid[w.id] };
            const beforeDays = countGroupHeadcountDeviationDays(grid, workers, days, groupCtx);
            const beforeMag = countGroupHeadcountDeviationMagnitude(grid, workers, days, groupCtx);
            const beforeAbs = countSupervisorAbsenceDays(grid, workers, days, groupCtx);
            grid[w.id][from] = false;
            grid[w.id][to] = true;
            if (
              !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
              validateDay(grid, workers, from, groupCtx, requireSupervisor).ok &&
              validateDay(grid, workers, to, groupCtx, requireSupervisor).ok &&
              isGroupHeadcountImproved(grid, workers, days, groupCtx, beforeDays, beforeMag, beforeAbs)
            ) {
              changed = true;
              break outer;
            }
            Object.assign(grid[w.id], backup);
          }
        }
      }
    }

    if (!changed) {
      for (let d = 1; d <= days; d++) {
        const snap = groupHeadcountSnapshot(grid, workers, d, groupCtx);
        for (const issue of snap.issues) {
          if (issue.kind !== "under") continue;
          const candidates = shuffledArray(
            workers.filter((w) => issue.match(w) && !grid[w.id][d] && !lockedOff[w.id]?.[d]),
            rng
          );
          for (const w of candidates) {
            const backup = { ...grid[w.id] };
            const beforeDays = countGroupHeadcountDeviationDays(grid, workers, days, groupCtx);
            const beforeMag = countGroupHeadcountDeviationMagnitude(grid, workers, days, groupCtx);
            const beforeAbs = countSupervisorAbsenceDays(grid, workers, days, groupCtx);
            grid[w.id][d] = true;
            if (
              violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) ||
              !validateDay(grid, workers, d, groupCtx, requireSupervisor).ok
            ) {
              Object.assign(grid[w.id], backup);
              continue;
            }
            const restored = tryAddOffKeepingSupervisors(
              grid,
              w,
              days,
              lockedOff,
              lockedWork,
              halfOff,
              maxConsecutive,
              groupCtx,
              workers,
              requireSupervisor,
              rng
            );
            if (
              restored &&
              isGroupHeadcountImproved(grid, workers, days, groupCtx, beforeDays, beforeMag, beforeAbs)
            ) {
              changed = true;
              break;
            }
            Object.assign(grid[w.id], backup);
          }
          if (changed) break;
        }
        if (changed) break;
      }
    }

    if (!changed) {
      for (let d = 1; d <= days; d++) {
        const snap = groupHeadcountSnapshot(grid, workers, d, groupCtx);
        for (const issue of snap.issues) {
          if (issue.kind !== "over") continue;
          const candidates = shuffledArray(
            workers.filter(
              (w) =>
                issue.match(w) &&
                grid[w.id][d] &&
                !lockedOff[w.id]?.[d] &&
                !lockedWork[w.id]?.[d] &&
                !halfOff[w.id]?.[d]
            ),
            rng
          );
          for (const w of candidates) {
            const backup = { ...grid[w.id] };
            const beforeDays = countGroupHeadcountDeviationDays(grid, workers, days, groupCtx);
            const beforeMag = countGroupHeadcountDeviationMagnitude(grid, workers, days, groupCtx);
            const beforeAbs = countSupervisorAbsenceDays(grid, workers, days, groupCtx);
            grid[w.id][d] = false;
            if (
              violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) ||
              !validateDay(grid, workers, d, groupCtx, requireSupervisor).ok ||
              isSupervisorShortOnDay(grid, workers, d, groupCtx)
            ) {
              Object.assign(grid[w.id], backup);
              continue;
            }
            const restored =
              tryAddWorkOnDays(
                grid,
                w,
                days,
                lockedOff,
                lockedWork,
                halfOff,
                maxConsecutive,
                groupCtx,
                workers,
                requireSupervisor,
                rng,
                true
              ) ||
              tryAddWorkOnDays(
                grid,
                w,
                days,
                lockedOff,
                lockedWork,
                halfOff,
                maxConsecutive,
                groupCtx,
                workers,
                requireSupervisor,
                rng,
                false
              );
            if (
              restored &&
              isGroupHeadcountImproved(grid, workers, days, groupCtx, beforeDays, beforeMag, beforeAbs)
            ) {
              changed = true;
              break;
            }
            Object.assign(grid[w.id], backup);
          }
          if (changed) break;
        }
        if (changed) break;
      }
    }

    if (!changed) break;
  }
}

function tryPlaceSupervisor(grid, workers, day, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, requireSupervisor, rng, predicate) {
  for (const w of shuffledArray(
    workers.filter(
      (w) =>
        predicate(w) &&
        !grid[w.id][day] &&
        !lockedOff[w.id]?.[day] &&
        !lockedWork[w.id]?.[day] &&
        !halfOff[w.id]?.[day]
    ),
    rng
  )) {
    grid[w.id][day] = true;
    if (
      !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
      validateDay(grid, workers, day, groupCtx, requireSupervisor).ok
    ) {
      return w;
    }
    grid[w.id][day] = false;
  }
  return null;
}

function repairDay(grid, workers, day, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, requireSupervisor, rng) {
  const hardOk = validateDay(grid, workers, day, groupCtx, requireSupervisor).ok;
  const working = workers.filter((w) => grid[w.id][day]);
  const supCount = working.filter((w) => w.isSupervisor).length;
  const { constraints, teams, teamConstraints } = groupCtx;
  const belowSupMin = supCount < (constraints.supervisorMin ?? 0);
  const belowGroupMin = teams.some((team) => {
    const tc = teamConstraints[team.id];
    const min = tc?.supervisorMin ?? 0;
    if (min <= 0) return false;
    const n = working.filter((w) => w.isGroupSupervisor && w.teamId === team.id).length;
    return n < min;
  });

  if (hardOk && (belowSupMin || belowGroupMin)) {
    if (belowSupMin) {
      tryPlaceSupervisor(
        grid,
        workers,
        day,
        days,
        lockedOff,
        lockedWork,
        halfOff,
        maxConsecutive,
        groupCtx,
        requireSupervisor,
        rng,
        (w) => w.isSupervisor
      );
    }
    for (const team of teams) {
      const tc = teamConstraints[team.id];
      const min = tc?.supervisorMin ?? 0;
      if (min <= 0) continue;
      const n = workers.filter((w) => w.isGroupSupervisor && w.teamId === team.id && grid[w.id][day]).length;
      if (n >= min) continue;
      tryPlaceSupervisor(
        grid,
        workers,
        day,
        days,
        lockedOff,
        lockedWork,
        halfOff,
        maxConsecutive,
        groupCtx,
        requireSupervisor,
        rng,
        (w) => w.isGroupSupervisor && w.teamId === team.id
      );
    }
    return true;
  }

  if (hardOk) return true;

  if (supCount > constraints.supervisorMax) {
    for (const w of shuffledArray(
      working.filter(
        (w) =>
          w.isSupervisor &&
          !lockedOff[w.id]?.[day] &&
          !lockedWork[w.id]?.[day] &&
          !halfOff[w.id]?.[day]
      ),
      rng
    )) {
      grid[w.id][day] = false;
      if (
        !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
        validateDay(grid, workers, day, groupCtx, requireSupervisor).ok
      ) {
        return true;
      }
      grid[w.id][day] = true;
    }
  }

  if (supCount < constraints.supervisorMin) {
    for (const w of shuffledArray(
      workers.filter(
        (w) =>
          w.isSupervisor &&
          !grid[w.id][day] &&
          !lockedOff[w.id]?.[day] &&
          !lockedWork[w.id]?.[day] &&
          !halfOff[w.id]?.[day]
      ),
      rng
    )) {
      grid[w.id][day] = true;
      if (
        !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
        validateDay(grid, workers, day, groupCtx, requireSupervisor).ok
      ) {
        return true;
      }
      grid[w.id][day] = false;
    }
  }

  for (const team of teams) {
    const tc = teamConstraints[team.id];
    if (!tc) continue;
    const groupSupMin = tc.supervisorMin ?? 0;
    const groupSupMax = tc.supervisorMax ?? 99;
    const groupSupWorking = workers.filter(
      (w) => w.isGroupSupervisor && w.teamId === team.id && grid[w.id][day]
    );
    const groupSupCount = groupSupWorking.length;

    if (groupSupCount > groupSupMax) {
      for (const w of shuffledArray(
        groupSupWorking.filter(
          (w) => !lockedOff[w.id]?.[day] && !lockedWork[w.id]?.[day] && !halfOff[w.id]?.[day]
        ),
        rng
      )) {
        grid[w.id][day] = false;
        if (
          !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
          validateDay(grid, workers, day, groupCtx, requireSupervisor).ok
        ) {
          return true;
        }
        grid[w.id][day] = true;
      }
    }

    if (groupSupCount < groupSupMin) {
      for (const w of shuffledArray(
        workers.filter(
          (w) =>
            w.isGroupSupervisor &&
            w.teamId === team.id &&
            !grid[w.id][day] &&
            !lockedOff[w.id]?.[day] &&
            !lockedWork[w.id]?.[day] &&
            !halfOff[w.id]?.[day]
        ),
        rng
      )) {
        grid[w.id][day] = true;
        if (
          !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
          validateDay(grid, workers, day, groupCtx, requireSupervisor).ok
        ) {
          return true;
        }
        grid[w.id][day] = false;
      }
    }
  }

  for (let attempt = 0; attempt < workers.length * 4; attempt++) {
    const workingNow = shuffledArray(
      workers.filter(
        (w) => grid[w.id][day] && !lockedOff[w.id]?.[day] && !lockedWork[w.id]?.[day] && !halfOff[w.id]?.[day]
      ),
      rng
    );
    for (const w of workingNow) {
      grid[w.id][day] = false;
      if (
        !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
        validateDay(grid, workers, day, groupCtx, requireSupervisor).ok
      ) {
        return true;
      }
      grid[w.id][day] = true;
    }

    const offPool = shuffledArray(
      workers.filter((w) => !grid[w.id][day] && !lockedOff[w.id]?.[day] && !lockedWork[w.id]?.[day] && !halfOff[w.id]?.[day]),
      rng
    );
    for (const w of offPool) {
      grid[w.id][day] = true;
      if (
        !violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive) &&
        validateDay(grid, workers, day, groupCtx, requireSupervisor).ok
      ) {
        return true;
      }
      grid[w.id][day] = false;
    }
  }

  return validateDay(grid, workers, day, groupCtx, requireSupervisor).ok;
}

function tryAddWork(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng) {
  return (
    tryAddWorkOnDays(
      grid,
      w,
      days,
      lockedOff,
      lockedWork,
      halfOff,
      maxConsecutive,
      groupCtx,
      workers,
      requireSupervisor,
      rng,
      true
    ) ||
    tryAddWorkOnDays(
      grid,
      w,
      days,
      lockedOff,
      lockedWork,
      halfOff,
      maxConsecutive,
      groupCtx,
      workers,
      requireSupervisor,
      rng,
      false
    )
  );
}

function tryAddWorkOnDays(grid, w, days, lockedOff, lockedWork, halfOff, maxConsecutive, groupCtx, workers, requireSupervisor, rng, keepGroupBounds) {
  void lockedWork;
  void halfOff;
  const preferred = workCandidateDays(grid[w.id], days, w, lockedOff, maxConsecutive);
  const candidates = preferred.length ? preferred : shuffledRange(1, days, rng);
  for (const d of candidates) {
    if (lockedOff[w.id]?.[d]) continue;
    if (grid[w.id][d]) continue;
    const beforeSnap = keepGroupBounds ? groupHeadcountSnapshot(grid, workers, d, groupCtx) : null;
    grid[w.id][d] = true;
    if (violatesWorkerStreaks(grid[w.id], days, lockedOff, w.id, maxConsecutive)) {
      grid[w.id][d] = false;
      continue;
    }
    if (!validateDay(grid, workers, d, groupCtx, requireSupervisor).ok) {
      grid[w.id][d] = false;
      continue;
    }
    if (keepGroupBounds && groupHeadcountWorsenedOnDay(grid, workers, d, groupCtx, beforeSnap)) {
      grid[w.id][d] = false;
      continue;
    }
    return true;
  }
  return false;
}

function validateDay(grid, workers, day, groupCtx, requireSupervisor) {
  const { constraints, teamConstraints, teams } = groupCtx;
  const working = workers.filter((w) => grid[w.id][day]);
  const count = working.length;

  const supCount = working.filter((w) => w.isSupervisor).length;
  if (supCount > constraints.supervisorMax) {
    return { ok: false };
  }

  const { lockedOff = {}, lockedWork = {}, halfOff = {} } = groupCtx;
  const canPlaceSupervisor = workers.some(
    (w) =>
      w.isSupervisor &&
      !grid[w.id][day] &&
      !lockedOff[w.id]?.[day] &&
      !lockedWork[w.id]?.[day] &&
      !halfOff[w.id]?.[day]
  );
  if (requireSupervisor && supCount < 1 && canPlaceSupervisor && count > 0) {
    return { ok: false };
  }

  for (const team of teams) {
    const tc = teamConstraints[team.id];
    if (!tc) continue;
    const groupSupCount = working.filter((w) => w.isGroupSupervisor && w.teamId === team.id).length;
    const groupSupMax = tc.supervisorMax ?? 99;
    if (groupSupCount > groupSupMax) {
      return { ok: false };
    }
  }

  return { ok: true };
}

function violatesConsecutiveAutoOff(row, days, lockedOff, workerId, maxAutoOff = MAX_AUTO_OFF_CONSECUTIVE) {
  let autoStreak = 0;
  for (let d = 1; d <= days; d++) {
    if (!row[d]) {
      if (lockedOff[workerId]?.[d]) {
        autoStreak = 0;
      } else {
        autoStreak++;
        if (autoStreak > maxAutoOff) return true;
      }
    } else {
      autoStreak = 0;
    }
  }
  return false;
}

function violatesWorkerStreaks(row, days, lockedOff, workerId, maxConsecutiveWork) {
  return (
    violatesConsecutiveWork(row, days, maxConsecutiveWork) ||
    violatesConsecutiveAutoOff(row, days, lockedOff, workerId)
  );
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

function gridToAssignments(grid, workers, days, halfOff, lockedOff = {}) {
  const assignments = {};
  for (const w of workers) {
    assignments[w.id] = {};
    for (let d = 1; d <= days; d++) {
      if (!grid[w.id][d]) {
        assignments[w.id][d] = { type: "off", preferredOff: Boolean(lockedOff[w.id]?.[d]) };
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

function scoreSchedule(grid, workers, days, preferences, requireSupervisor, lockedOff, maxConsecutiveWork, groupCtx = {}) {
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
    score += scoreWorkStreaks(grid[w.id], days, maxConsecutiveWork);
    score += scoreAutoOffPlacement(grid[w.id], days, lockedOff, w.id);
    if (violatesConsecutiveAutoOff(grid[w.id], days, lockedOff, w.id)) {
      score -= 100;
    }
  }
  const { constraints, teams = [], teamConstraints = {} } = groupCtx;
  for (let d = 1; d <= days; d++) {
    const working = workers.filter((w) => grid[w.id][d]);
    const sup = working.filter((w) => w.isSupervisor).length;
    if (requireSupervisor && sup >= 1) score += 2;
    if (isSupervisorShortOnDay(grid, workers, d, groupCtx)) score -= 10000;
    const headcount = groupHeadcountSnapshot(grid, workers, d, groupCtx);
    if (headcount.deviant) score -= 5000 + 40 * headcount.magnitude;
    const supMin = constraints?.supervisorMin ?? 0;
    if (sup < supMin) score -= 80 * (supMin - sup);
    for (const team of teams) {
      const tc = teamConstraints[team.id];
      if (!tc) continue;
      const groupMin = tc.supervisorMin ?? 0;
      if (groupMin <= 0) continue;
      const groupSup = working.filter((w) => w.isGroupSupervisor && w.teamId === team.id).length;
      if (groupSup < groupMin) score -= 60 * (groupMin - groupSup);
    }
  }
  return score;
}

function isWorkingAssignment(cell) {
  return Boolean(cell && cell.type !== "off");
}

/**
 * 責任者の下限を下回った日。表示テキストは常に「責任者不在」。
 * @returns {Record<number, { text: string, overall: boolean, teams: string[] }>}
 */
export function buildSupervisorAbsence(assignments, workers, days, constraints, teams = [], teamConstraints = {}) {
  const byDay = {};
  for (let d = 1; d <= days; d++) {
    const working = workers.filter((w) => isWorkingAssignment(assignments[w.id]?.[d]));
    const overall = working.filter((w) => w.isSupervisor).length < (constraints?.supervisorMin ?? 0);
    const missingTeams = [];
    for (const team of teams) {
      const tc = teamConstraints[team.id];
      const min = tc?.supervisorMin ?? 0;
      if (min <= 0) continue;
      const groupSup = working.filter((w) => w.isGroupSupervisor && w.teamId === team.id).length;
      if (groupSup < min) missingTeams.push(team.name);
    }
    if (overall || missingTeams.length) {
      byDay[d] = { text: "責任者不在", overall, teams: missingTeams };
    }
  }
  return byDay;
}

/**
 * グループ出勤人数の下限・上限から逸脱した日。
 * @returns {Record<number, { text: string, groups: string[] }>}
 */
export function buildGroupHeadcountDeviation(
  assignments,
  workers,
  days,
  teams = [],
  teamConstraints = {},
  subGroups = [],
  subGroupConstraints = {}
) {
  const groupCtx = { teams, teamConstraints, subGroups, subGroupConstraints };
  const byDay = {};
  for (let d = 1; d <= days; d++) {
    const working = workers.filter((w) => isWorkingAssignment(assignments[w.id]?.[d]));
    const snap = groupHeadcountSnapshotFromWorking(working, groupCtx);
    if (!snap.deviant) continue;
    byDay[d] = {
      text: "人数逸脱",
      groups: snap.issues.map((issue) => {
        const mark = issue.kind === "under" ? "不足" : "超過";
        return `${issue.name}${mark}(${issue.count}/${issue.min}〜${issue.max})`;
      }),
    };
  }
  return byDay;
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
      groupSupervisors: working.filter((w) => w.isGroupSupervisor).length,
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
 * グループごとに月4回（隔週優先）、参加者が多い日をカンファレンス日とする
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

function countAttendingMembers(members, assignments, day) {
  return members.filter((m) => {
    const c = assignments[m.id][day];
    return c && c.type !== "off";
  }).length;
}

function combinations(arr, k) {
  const result = [];
  function pick(start, combo) {
    if (combo.length === k) {
      result.push([...combo]);
      return;
    }
    for (let i = start; i <= arr.length - (k - combo.length); i++) {
      combo.push(arr[i]);
      pick(i + 1, combo);
      combo.pop();
    }
  }
  pick(0, []);
  return result;
}

function scoreConferencePick(picked) {
  if (!picked.length) return -Infinity;
  let score = picked.reduce((s, p) => s + p.count * 1000, 0);
  const ordered = [...picked].sort((a, b) => a.weekIndex - b.weekIndex);
  for (let i = 1; i < ordered.length; i++) {
    const gap = ordered[i].weekIndex - ordered[i - 1].weekIndex;
    if (gap === 2) score += 100;
    else if (gap === 1) score -= 50;
    else if (gap >= 3) score += 20;
  }
  return score;
}

function pickMonthlyConferenceDays(members, assignments, year, month, maxSessions = CONFERENCE_SESSIONS_PER_MONTH) {
  const days = getDaysInMonth(year, month);
  if (!members.length) return [];

  const weekToDays = new Map();
  for (let d = 1; d <= days; d++) {
    const wk = getWeekKey(year, month, d);
    if (!weekToDays.has(wk)) weekToDays.set(wk, []);
    weekToDays.get(wk).push(d);
  }

  const weekKeys = [...weekToDays.keys()].sort(
    (a, b) => weekToDays.get(a)[0] - weekToDays.get(b)[0]
  );

  const weekCandidates = [];
  for (let wi = 0; wi < weekKeys.length; wi++) {
    const weekKey = weekKeys[wi];
    const daysInWeek = weekToDays.get(weekKey);
    let bestDay = null;
    let bestCount = -1;
    for (const d of daysInWeek) {
      const count = countAttendingMembers(members, assignments, d);
      if (count > bestCount || (count === bestCount && (bestDay == null || d < bestDay))) {
        bestCount = count;
        bestDay = d;
      }
    }
    if (bestDay != null && bestCount > 0) {
      weekCandidates.push({ weekKey, weekIndex: wi, day: bestDay, count: bestCount });
    }
  }

  if (!weekCandidates.length) return [];

  const target = Math.min(maxSessions, weekCandidates.length);
  let bestPick = weekCandidates.slice(0, target);

  if (weekCandidates.length > target) {
    let bestScore = scoreConferencePick(bestPick);
    for (const combo of combinations(weekCandidates, target)) {
      const score = scoreConferencePick(combo);
      if (score > bestScore) {
        bestScore = score;
        bestPick = combo;
      }
    }
  }

  bestPick.sort((a, b) => a.day - b.day);
  return bestPick.map(({ weekKey, day, count }) => ({ weekKey, day, count }));
}

function pickWeeklyConferenceDays(members, assignments, year, month) {
  return pickMonthlyConferenceDays(members, assignments, year, month);
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

export function buildConferenceColorMap(teams, subGroups, teamConstraints, subGroupConstraints) {
  const map = new Map();
  let teamIdx = 0;
  let sgIdx = 0;
  for (const team of teams) {
    if (teamConstraints[team.id]?.useConferenceDay) {
      map.set(`team:${team.id}`, CONFERENCE_COLORS[teamIdx++ % CONFERENCE_COLORS.length]);
    }
  }
  for (const sg of subGroups) {
    if (subGroupConstraints[sg.id]?.useConferenceDay) {
      map.set(`subGroup:${sg.id}`, SUBGROUP_CONFERENCE_COLORS[sgIdx++ % SUBGROUP_CONFERENCE_COLORS.length]);
    }
  }
  return map;
}

export function getConferenceGroupStyle(colorMap, type, id) {
  return colorMap.get(`${type}:${id}`) ?? null;
}

export function getWorkerConferenceGroup(conferenceDays, worker, day) {
  const cd = normalizeConferenceDays(conferenceDays);
  if (worker.subGroupId && cd.subGroups[worker.subGroupId]?.some((e) => e.day === day)) {
    return { type: "subGroup", id: worker.subGroupId };
  }
  if (worker.teamId && cd.teams[worker.teamId]?.some((e) => e.day === day)) {
    return { type: "team", id: worker.teamId };
  }
  return null;
}

export function isConferenceDayForWorker(conferenceDays, worker, day) {
  return Boolean(getWorkerConferenceGroup(conferenceDays, worker, day));
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
  if (cell.type === "half-off") return cell.half === "am" ? "AM" : "PM";
  if (useShiftTypes && cell.shiftType) return cell.shiftType;
  return "１";
}

export function isAttendingCell(cell) {
  return Boolean(cell && cell.type !== "off");
}
