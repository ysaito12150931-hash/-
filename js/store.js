const STORAGE_KEY = "shift-app-v1";

/** 画面右下に表示。更新後にここが変わっていれば最新版です */
export const APP_VERSION = "1.3.0";

function emptyDefaults() {
  return {
    year: new Date().getFullYear(),
    month: new Date().getMonth() + 1,
    maxConsecutiveWork: 5,
    useShiftTypes: false,
    shiftTypes: ["早番", "遅番"],
    workers: [],
    teams: [],
    subGroups: [],
    constraints: {
      supervisorMin: 1,
      supervisorMax: 2,
    },
    teamConstraints: {},
    subGroupConstraints: {},
    preferences: {},
    lastResult: null,
  };
}

export const defaultState = () => {
  const teamA = { id: crypto.randomUUID(), name: "チームA" };
  const teamB = { id: crypto.randomUUID(), name: "チームB" };
  return {
    ...emptyDefaults(),
    workers: [
      { id: crypto.randomUUID(), name: "田中", teamId: teamA.id, subGroupId: null, isSupervisor: true, monthlyOffDays: 8 },
      { id: crypto.randomUUID(), name: "佐藤", teamId: teamA.id, subGroupId: null, isSupervisor: false, monthlyOffDays: 9 },
      { id: crypto.randomUUID(), name: "鈴木", teamId: teamB.id, subGroupId: null, isSupervisor: false, monthlyOffDays: 9 },
    ],
    teams: [teamA, teamB],
    teamConstraints: {
      [teamA.id]: { min: 1, max: 3, useConferenceDay: true },
      [teamB.id]: { min: 1, max: 3, useConferenceDay: false },
    },
  };
};

export function migrateLoadedState(state) {
  state.subGroups = state.subGroups || [];
  state.subGroupConstraints = state.subGroupConstraints || {};
  state.teamConstraints = state.teamConstraints || {};
  state.workers = (state.workers || []).map((w) => ({
    ...w,
    subGroupId: w.subGroupId ?? null,
  }));

  for (const w of state.workers) {
    if (w.subGroupId) {
      const sg = state.subGroups.find((g) => g.id === w.subGroupId);
      if (!sg || sg.teamId !== w.teamId) w.subGroupId = null;
    }
  }

  for (const t of state.teams || []) {
    if (!state.teamConstraints[t.id]) {
      state.teamConstraints[t.id] = { min: 0, max: 99, useConferenceDay: false };
    } else if (state.teamConstraints[t.id].useConferenceDay == null) {
      state.teamConstraints[t.id].useConferenceDay = false;
    }
  }

  for (const sg of state.subGroups) {
    if (!state.subGroupConstraints[sg.id]) {
      state.subGroupConstraints[sg.id] = { min: 0, max: 99 };
    }
  }

  return state;
}

export function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    return migrateLoadedState({ ...emptyDefaults(), ...parsed });
  } catch {
    return defaultState();
  }
}

export function saveState(state) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    return true;
  } catch (e) {
    console.error("saveState failed", e);
    return false;
  }
}

export function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}
