function workerKey(w) {
  return w.id ?? w.name;
}

function hasValidTeam(worker, teams) {
  return Boolean(worker.teamId && teams.some((t) => t.id === worker.teamId));
}

function membersInOrder(workers, predicate) {
  return workers.filter(predicate);
}

/**
 * 勤務者をグループ単位のセクションに分割（表示・Excel共通）
 * - メイングループ: teams 配列の順
 * - サブグループ: subGroups 配列の順（同一メイングループ内）
 * - 各セクション内の勤務者: workers 配列の登録順
 */
export function getWorkerSections(workers, teams, subGroups = []) {
  const sections = [];
  const assigned = new Set();

  const mark = (list) => list.forEach((w) => assigned.add(workerKey(w)));

  for (const team of teams) {
    const teamMembers = membersInOrder(workers, (w) => w.teamId === team.id);
    if (!teamMembers.length) continue;

    const teamSubs = subGroups.filter((sg) => sg.teamId === team.id);
    const subAssigned = new Set();

    for (const sg of teamSubs) {
      const sgMembers = membersInOrder(workers, (w) => w.subGroupId === sg.id);
      if (!sgMembers.length) continue;
      sections.push({
        key: `team:${team.id}:sg:${sg.id}`,
        label: `${team.name} / ${sg.name}`,
        team,
        subGroup: sg,
        members: sgMembers,
      });
      sgMembers.forEach((w) => {
        assigned.add(workerKey(w));
        subAssigned.add(workerKey(w));
      });
    }

    const noSubMembers = teamMembers.filter((w) => !subAssigned.has(workerKey(w)));
    if (noSubMembers.length) {
      sections.push({
        key: `team:${team.id}`,
        label: team.name,
        team,
        subGroup: null,
        members: noSubMembers,
      });
      mark(noSubMembers);
    }
  }

  const supervisorsNoTeam = membersInOrder(
    workers,
    (w) => w.isSupervisor && !hasValidTeam(w, teams) && !assigned.has(workerKey(w))
  );
  if (supervisorsNoTeam.length) {
    sections.push({
      key: "supervisors-unassigned",
      label: "責任者（メイングループ未所属）",
      team: null,
      subGroup: null,
      members: supervisorsNoTeam,
    });
    mark(supervisorsNoTeam);
  }

  const unassigned = membersInOrder(workers, (w) => !assigned.has(workerKey(w)));
  if (unassigned.length) {
    sections.push({
      key: "unassigned",
      label: "（未所属）",
      team: null,
      subGroup: null,
      members: unassigned,
    });
  }

  return sections;
}

/** 連続する同一メイングループのセクションをまとめる */
export function groupWorkerSectionsByTeam(sections) {
  const groups = [];
  for (const section of sections) {
    const teamId = section.team?.id ?? null;
    const last = groups[groups.length - 1];
    if (teamId && last && last.teamId === teamId) {
      last.sections.push(section);
    } else {
      groups.push({
        teamId,
        team: section.team,
        sections: [section],
      });
    }
  }
  return groups;
}

export function getTeamHeadcountBounds(team, teamConstraints = {}) {
  if (!team) return null;
  const tc = teamConstraints[team.id];
  if (!tc) return null;
  return { min: tc.min ?? 0, max: tc.max ?? 99, label: team.name };
}

/** セクションの出勤合計に適用する下限・上限。サブグループがあるメインの残りメンバーにはチーム制約を使わない。 */
export function getSectionHeadcountBounds(section, teamConstraints = {}, subGroupConstraints = {}, subGroups = []) {
  if (section.subGroup) {
    const sgc = subGroupConstraints[section.subGroup.id];
    if (!sgc) return null;
    return { min: sgc.min ?? 0, max: sgc.max ?? 99, label: section.label };
  }
  if (section.team) {
    const hasSubs = subGroups.some((sg) => sg.teamId === section.team.id);
    if (hasSubs) return null;
    const tc = teamConstraints[section.team.id];
    if (!tc) return null;
    return { min: tc.min ?? 0, max: tc.max ?? 99, label: section.label };
  }
  return null;
}

export function isHeadcountOutOfRange(count, bounds) {
  if (!bounds) return false;
  const min = bounds.min ?? 0;
  const max = bounds.max ?? 99;
  if (min <= 0 && max >= 99) return false;
  return count < min || count > max;
}

/** @returns {string|null} */
export function getWorkerSectionKey(worker, teams, subGroups = []) {
  if (worker.subGroupId && subGroups.some((sg) => sg.id === worker.subGroupId)) {
    return `team:${worker.teamId}:sg:${worker.subGroupId}`;
  }
  if (worker.teamId && teams.some((t) => t.id === worker.teamId)) {
    return `team:${worker.teamId}`;
  }
  if (worker.isSupervisor && !hasValidTeam(worker, teams)) {
    return "supervisors-unassigned";
  }
  return "unassigned";
}

export function moveWorkerWithinSection(workers, workerId, delta, teams, subGroups) {
  const sections = getWorkerSections(workers, teams, subGroups);
  for (const section of sections) {
    const idx = section.members.findIndex((w) => w.id === workerId);
    if (idx === -1) continue;
    const target = idx + delta;
    if (target < 0 || target >= section.members.length) return false;

    const a = workers.findIndex((w) => w.id === workerId);
    const b = workers.findIndex((w) => w.id === section.members[target].id);
    if (a === -1 || b === -1) return false;
    [workers[a], workers[b]] = [workers[b], workers[a]];
    return true;
  }
  return false;
}

export function moveTeamOrder(teams, idx, delta) {
  const target = idx + delta;
  if (target < 0 || target >= teams.length) return false;
  [teams[idx], teams[target]] = [teams[target], teams[idx]];
  return true;
}

export function moveSubGroupOrder(subGroups, subGroupId, delta) {
  const idx = subGroups.findIndex((sg) => sg.id === subGroupId);
  if (idx === -1) return false;
  const teamId = subGroups[idx].teamId;
  const sameTeamIdxs = subGroups
    .map((sg, i) => (sg.teamId === teamId ? i : -1))
    .filter((i) => i >= 0);
  const pos = sameTeamIdxs.indexOf(idx);
  const targetPos = pos + delta;
  if (targetPos < 0 || targetPos >= sameTeamIdxs.length) return false;
  const a = sameTeamIdxs[pos];
  const b = sameTeamIdxs[targetPos];
  [subGroups[a], subGroups[b]] = [subGroups[b], subGroups[a]];
  return true;
}

export function canMoveSubGroupUp(subGroups, subGroupId) {
  const idx = subGroups.findIndex((sg) => sg.id === subGroupId);
  if (idx === -1) return false;
  const teamId = subGroups[idx].teamId;
  const first = subGroups.findIndex((sg) => sg.teamId === teamId);
  return idx > first;
}

export function canMoveSubGroupDown(subGroups, subGroupId) {
  const idx = subGroups.findIndex((sg) => sg.id === subGroupId);
  if (idx === -1) return false;
  const teamId = subGroups[idx].teamId;
  const last = subGroups.reduce((acc, sg, i) => (sg.teamId === teamId ? i : acc), -1);
  return idx < last;
}
