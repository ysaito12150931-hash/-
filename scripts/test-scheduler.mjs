import { generateShift } from "../js/scheduler.js";
import { defaultState } from "../js/store.js";

const state = defaultState();
state.workers = [
  { id: "1", name: "田中", teamId: null, isSupervisor: true, isGroupSupervisor: false, monthlyOffDays: 8 },
  { id: "2", name: "佐藤", teamId: null, isSupervisor: false, isGroupSupervisor: false, monthlyOffDays: 9 },
  { id: "3", name: "鈴木", teamId: null, isSupervisor: false, isGroupSupervisor: false, monthlyOffDays: 9 },
  { id: "4", name: "高橋", teamId: null, isSupervisor: false, isGroupSupervisor: false, monthlyOffDays: 8 },
  { id: "5", name: "伊藤", teamId: null, isSupervisor: true, isGroupSupervisor: false, monthlyOffDays: 8 },
];
state.workers.forEach((w, i) => {
  w.teamId = state.teams[i % state.teams.length]?.id ?? null;
  w.isGroupSupervisor = w.isSupervisor;
});
state.teamConstraints[state.teams[0].id].supervisorMin = 0;
state.teamConstraints[state.teams[0].id].supervisorMax = 2;
state.teamConstraints[state.teams[1].id].supervisorMin = 0;
state.teamConstraints[state.teams[1].id].supervisorMax = 2;
state.constraints = { supervisorMin: 1, supervisorMax: 2 };
state.year = 2026;
state.month = 5;

const r = generateShift(state);
console.log(r.ok ? "OK" : "FAIL", r.messages?.[0]);
if (r.ok) {
  const days = 31;
  for (let d = 1; d <= 5; d++) {
    const n = state.workers.filter((w) => r.assignments[w.id][d]?.type === "work").length;
    const s = state.workers.filter(
      (w) => w.isSupervisor && r.assignments[w.id][d]?.type === "work"
    ).length;
    console.log(`day ${d}: work=${n} sup=${s}`);
  }
}
