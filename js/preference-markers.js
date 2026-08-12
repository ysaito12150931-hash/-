/** 希望休とみなす丸印（U+3007 〇 / U+25CB ○ / U+25EF ◯ など） */
const CIRCLE_OFF_CODES = new Set([
  0x3007, // 〇 IDEOGRAPHIC NUMBER ZERO
  0x25cb, // ○ WHITE CIRCLE
  0x25ef, // ◯ LARGE CIRCLE
  0x25e6, // ◦ WHITE BULLET
  0x25e7, // ◧
  0xff10, // ０ fullwidth zero
]);

const OFF_MARKERS = new Set([
  "休",
  "休み",
  "×",
  "x",
  "X",
  "off",
  "OFF",
  "0",
  "希望休",
  "公休",
  "〇",
  "○",
  "◯",
  "◦",
  "０",
]);

const WORK_MARKERS = new Set([
  "出勤",
  "出",
  "勤",
  "勤務",
  "work",
  "WORK",
  "1",
  "１",
  "丸",
  "希望出勤",
  "◎",
]);

const AM_OFF_MARKERS = new Set([
  "午前休",
  "前休",
  "午前のみ休み",
  "午前だけ休み",
  "AM休",
  "am休",
  "am-off",
  "AM-OFF",
  "午前休み",
]);

const PM_OFF_MARKERS = new Set([
  "午後休",
  "後休",
  "午後のみ休み",
  "午後だけ休み",
  "PM休",
  "pm休",
  "pm-off",
  "PM-OFF",
  "午後休み",
]);

export function normalizeMarkerText(value) {
  if (value == null || value === "") return "";
  return String(value)
    .replace(/[\uFE0E\uFE0F]/g, "")
    .replace(/[\s\u3000]/g, "")
    .trim();
}

function isCircleOffText(s) {
  const t = normalizeMarkerText(s);
  if (!t) return false;
  if (t === "0" || t === "０") return true;
  const chars = [...t];
  if (chars.length === 1 && CIRCLE_OFF_CODES.has(chars[0].codePointAt(0))) return true;
  return /^[〇○◯◦０0]$/u.test(t);
}

export function isOffMarker(value) {
  if (value == null || value === "") return false;
  if (typeof value === "number" && value === 0) return true;
  const s = normalizeMarkerText(value);
  if (!s) return false;
  if (OFF_MARKERS.has(s) || OFF_MARKERS.has(s.toLowerCase())) return true;
  return isCircleOffText(s);
}

export function isWorkMarker(value) {
  if (value == null || value === "") return false;
  if (typeof value === "number" && value === 1) return true;
  const s = normalizeMarkerText(value);
  if (!s) return false;
  if (isCircleOffText(s)) return false;
  return WORK_MARKERS.has(s) || WORK_MARKERS.has(s.toLowerCase());
}

export function isAmOffMarker(value) {
  if (value == null || value === "") return false;
  return AM_OFF_MARKERS.has(normalizeMarkerText(value));
}

export function isPmOffMarker(value) {
  if (value == null || value === "") return false;
  return PM_OFF_MARKERS.has(normalizeMarkerText(value));
}

/** @returns {"off"|"work"|"am-off"|"pm-off"|"conflict"|null} */
export function parsePreferenceCell(value) {
  if (value == null || value === "") return null;
  if (typeof value === "boolean") return value ? "off" : null;

  const am = isAmOffMarker(value);
  const pm = isPmOffMarker(value);
  if (am && pm) return "conflict";
  if (am) return "am-off";
  if (pm) return "pm-off";

  const off = isOffMarker(value);
  const work = isWorkMarker(value);
  if (off && work) return "conflict";
  if (off) return "off";
  if (work) return "work";
  return null;
}

export function isPreferredOffValue(value) {
  if (value === true) return true;
  return parsePreferenceCell(value) === "off";
}

export function getPreferenceValue(prefMap, day) {
  if (!prefMap) return undefined;
  return prefMap[day] ?? prefMap[String(day)] ?? prefMap[Number(day)];
}

export function getWorkerPrefMap(preferences, workerName) {
  if (!preferences) return undefined;
  if (preferences[workerName]) return preferences[workerName];
  const want = String(workerName ?? "").trim();
  const key = Object.keys(preferences).find((n) => String(n).trim() === want);
  return key ? preferences[key] : undefined;
}
