const STORAGE_KEY = "shift-app-v1";

export const defaultState = () => {
  const teamA = { id: crypto.randomUUID(), name: "チームA" };
  const teamB = { id: crypto.randomUUID(), name: "チームB" };
  return {
  year: new Date().getFullYear(),
  month: new Date().getMonth() + 1,
  maxConsecutiveWork: 5,
  useShiftTypes: false,
  shiftTypes: ["早番", "遅番"],
  workers: [
    { id: crypto.randomUUID(), name: "田中", teamId: teamA.id, isSupervisor: true, monthlyOffDays: 8 },
    { id: crypto.randomUUID(), name: "佐藤", teamId: teamA.id, isSupervisor: false, monthlyOffDays: 9 },
    { id: crypto.randomUUID(), name: "鈴木", teamId: teamB.id, isSupervisor: false, monthlyOffDays: 9 },
  ],
  teams: [teamA, teamB],
  constraints: {
    dailyMin: 3,
    dailyMax: 8,
    supervisorMin: 1,
    supervisorMax: 2,
  },
  teamConstraints: {
    [teamA.id]: { min: 1, max: 3 },
    [teamB.id]: { min: 1, max: 3 },
  },
  preferences: {},
  lastResult: null,
};
};

export function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    return { ...defaultState(), ...parsed };
  } catch {
    return defaultState();
  }
}

export function saveState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

export function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}
