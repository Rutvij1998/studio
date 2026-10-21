import { formatCount, type Book, type Reading, type ScopeCard, type Track, type TrackPoint } from "./book";
import { nextPeriodLabel, projectAt } from "./grid";

const MONTH =
  "(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)";

const ROW = new RegExp(
  `^${MONTH}\\s*[-/\\s]?\\s*(\\d{2,4})\\s*[:|]?\\s*([\\d,]+(?:\\.\\d+)?)(?:\\s*(?:calls|chats|contacts)\\s*[,|]?\\s*|\\s*[,|]\\s*|\\s+)([\\d,]+(?:\\.\\d+)?)\\s*(?:contained|containment|calls|chats)?(?:\\s*[,|]\\s*|\\s+)(\\d+(?:\\.\\d+)?)\\s*%?\\s*$`,
  "i",
);

const MONTH_LINE = new RegExp(`^${MONTH}\\s*[-/\\s]?\\s*\\d{2,4}$`, "i");
const COUNT_LINE = /^[\d,]+(?:\.\d+)?%?$/;

type Section = { name: string; body: string };

/** A claims book keeps its own slides. A series is used only when those counts are absent. */
export function isSeries(book: Book): boolean {
  const reading = book.reading;
  if (!reading || reading.tracks.length === 0) return false;
  if (book.active > 0 || book.claims > 0) return false;
  return true;
}

/** A pasted sheet often puts each cell on its own line. Join a month with the counts under it. */
function unstackCells(text: string): string {
  const lines = text.split("\n");
  const out: string[] = [];
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index] ?? "";
    const trimmed = line.trim();
    if (!MONTH_LINE.test(trimmed)) {
      out.push(line);
      continue;
    }
    const counts: string[] = [];
    let next = index + 1;
    while (next < lines.length && counts.length < 3) {
      const cell = (lines[next] ?? "").trim();
      if (!cell) {
        const after = (lines[next + 1] ?? "").trim();
        if (!after || !COUNT_LINE.test(after)) break;
        next += 1;
        continue;
      }
      if (!COUNT_LINE.test(cell)) break;
      counts.push(cell);
      next += 1;
    }
    if (counts.length >= 2) {
      out.push([trimmed, ...counts].join(" "));
      index = next - 1;
    } else {
      out.push(line);
    }
  }
  return out.join("\n");
}

/** The brief looks like a month table, and it still did not become a series. */
export function missedTable(text: string): string | null {
  if (!text.trim() || readMaterial(text)) return null;
  const lines = text.replace(/\r/g, "").split("\n");
  const months = lines.filter((line) => MONTH_LINE.test(line.trim())).length;
  const header = lines.some((line) => /\bmm_yyyy\b|\btotal calls\b|\bcontained calls\b|\btotal chats\b|\bcontained chats\b/i.test(line));
  if (months >= 3 || (header && months >= 1)) {
    return "A month table is in the brief, and those rows were not read. Nothing was filled in.";
  }
  return null;
}

/** The source already became charts, so it does not need a second slide of the same rows. */
export function seriesChartsSource(book: Book): boolean {
  if (!isSeries(book) || !book.sourceText.trim()) return false;
  return book.reading!.tracks.some((track) =>
    track.points.some((point) => book.sourceText.includes(point.label) && book.sourceText.includes(String(point.total))),
  );
}

export function readMaterial(text: string): Reading | null {
  const clean = unstackCells(text.replace(/\r/g, "")).trim();
  if (!clean) return null;
  const sections = splitSections(clean);
  const tracks = sections
    .map((section) => trackFromBody(section.name, section.body))
    .filter((track): track is Track => track != null);
  if (tracks.length === 0) return null;
  const scope = scopesFrom(clean, sections);
  const title = titleOf(clean) || tracks[0]?.name || "";
  const period = periodField(clean) || periodFromTracks(tracks);
  return {
    title,
    period,
    subject: subjectFrom(tracks, clean),
    intent: intentOf(clean, tracks),
    scope,
    ask: askOf(clean, scope.length > 0),
    tracks,
  };
}

export function readingFromTracks(tracks: Track[], titleHint: string): Reading {
  const title = tidyTitle(titleHint.replace(/\.(xlsx|xls|csv)$/i, "").replace(/[-_]+/g, " "));
  return {
    title: title || tracks[0]?.name || "",
    period: periodFromTracks(tracks),
    subject: subjectFrom(tracks, ""),
    intent:
      tracks.length > 1
        ? "Show each series in the sheet, month by month, then side by side."
        : "Show how this series moves across the months in the sheet.",
    scope: [],
    ask: "",
    tracks,
  };
}

/** Sheet numbers win. The brief still supplies the title, the scope, and the intent. */
export function mergeReadings(text: Reading | null, sheet: Reading | null): Reading | null {
  if (!sheet?.tracks.length) return text;
  if (!text) return sheet;
  return {
    title: text.title || sheet.title,
    period: text.period || sheet.period,
    subject: text.scope.length ? text.subject : sheet.subject || text.subject,
    intent: text.intent || sheet.intent,
    scope: text.scope.length ? text.scope : sheet.scope,
    ask: text.ask || sheet.ask,
    tracks: sheet.tracks,
  };
}

export function trackTotals(track: Track): { total: number; part: number; rate: number } {
  const total = track.points.reduce((sum, point) => sum + point.total, 0);
  const part = track.points.reduce((sum, point) => sum + point.part, 0);
  return { total, part, rate: total > 0 ? Math.round((part / total) * 100) : 0 };
}

export function readingLine(reading: Reading): string {
  const bits = reading.tracks.map((track) => {
    const totals = trackTotals(track);
    return `${track.name} ${totals.rate}% of ${formatCount(totals.total)} ${track.unit}`;
  });
  return bits.length ? `${bits.join(". ")}.` : reading.subject;
}

export function readingSummary(reading: Reading): string {
  const title = reading.title.trim() || "Untitled";
  return `Title: ${title}. Data: ${reading.subject} Intent: ${reading.intent}`;
}

/** Full months only. A tiny month and a short last month stay out, matching the chart note. */
export function fullPoints(track: Track): TrackPoint[] {
  const counts = track.points.map((point) => point.total).sort((left, right) => left - right);
  const mid = counts[Math.floor(counts.length / 2)] ?? 0;
  const tiny = track.points.filter((point) => mid > 0 && point.total < Math.max(100, mid * 0.02));
  const last = track.points[track.points.length - 1];
  const shortLast = Boolean(last && mid > 0 && last.total < mid * 0.55 && !tiny.some((point) => point.label === last.label));
  return track.points.filter(
    (point) => !tiny.some((item) => item.label === point.label) && !(shortLast && point.label === last?.label),
  );
}

export type MonthProgress = {
  name: string;
  label: string;
  contained: number;
  rest: number;
  rate: number;
  previousRate: number | null;
  direction: "Progress" | "Not progress";
  unit: string;
};

/** Latest full month only. Contained against the rest of that month, not a sum of months. */
export function seriesProgressList(reading: Reading): MonthProgress[] {
  const cards: MonthProgress[] = [];
  for (const track of reading.tracks) {
    const full = fullPoints(track);
    const latest = full[full.length - 1];
    if (!latest) continue;
    const previous = full.length >= 2 ? full[full.length - 2] : null;
    cards.push({
      name: track.name,
      label: latest.label,
      contained: latest.part,
      rest: Math.max(0, latest.total - latest.part),
      rate: latest.rate,
      previousRate: previous ? previous.rate : null,
      direction: !previous || latest.rate >= previous.rate ? "Progress" : "Not progress",
      unit: track.unit,
    });
  }
  return cards;
}

export type RateOutlook = {
  name: string;
  label: string;
  rate: number;
  from: string;
  to: string;
};

/** One period ahead of the sheet. The line is fit on full months only. */
export function rateOutlook(track: Track): RateOutlook | null {
  const full = fullPoints(track);
  const last = track.points[track.points.length - 1];
  if (full.length < 3 || !last) return null;
  const serials = full.map((point) => monthSerial(point.label));
  const aligned = serials.every((value) => value != null);
  const xs = aligned ? (serials as number[]) : full.map((_, index) => index);
  const lastSerial = monthSerial(last.label);
  const nextX = aligned && lastSerial != null ? lastSerial + 1 : (xs[xs.length - 1] ?? 0) + 1;
  const projected = projectAt(xs, full.map((point) => point.rate), nextX);
  if (projected == null) return null;
  return {
    name: track.name,
    label: nextPeriodLabel(last.label, "month"),
    rate: Math.round(projected),
    from: full[0]!.label,
    to: full[full.length - 1]!.label,
  };
}

export function trackNote(track: Track): string {
  const counts = track.points.map((point) => point.total).sort((left, right) => left - right);
  const mid = counts[Math.floor(counts.length / 2)] ?? 0;
  const tiny = track.points.filter((point) => mid > 0 && point.total < Math.max(100, mid * 0.02));
  const last = track.points[track.points.length - 1];
  const shortLast = Boolean(last && mid > 0 && last.total < mid * 0.55 && !tiny.some((point) => point.label === last.label));
  const full = fullPoints(track);
  const notes: string[] = [];
  if (full.length >= 2) {
    const start = full[0]!;
    const end = full[full.length - 1]!;
    if (start.rate !== end.rate) {
      notes.push(`Across full months, the rate moves from ${start.rate}% in ${start.label} to ${end.rate}% in ${end.label}.`);
    }
  }
  if (tiny.length > 0) {
    notes.push(`${tiny.map((point) => `${point.label} is ${formatCount(point.total)} ${track.unit}`).join(". ")}.`);
  }
  if (shortLast && last) {
    notes.push(`${last.label} is much smaller than the middle of the series, so read it as a short month.`);
  }
  const drop = biggestDrop({ ...track, points: full }, tiny);
  if (drop) notes.push(drop);
  notes.push("The line is that month's containment rate.");
  return notes.join(" ");
}

/** A grid whose first column is months and whose other columns are counts. */
export function trackFromRows(rows: string[][], sheetName: string): Track | null {
  if (rows.length < 4) return null;
  const headers = rows[0]!.map((cell) => cell.trim().toLowerCase());
  const rateAt = headers.findIndex((header) => header === "%" || /percent|rate/.test(header) || header.includes("%"));
  const partAt = headers.findIndex((header, index) => index !== rateAt && /contain/.test(header));
  const totalAt = headers.findIndex(
    (header, index) => index !== partAt && index !== rateAt && (/total/.test(header) || header === "calls" || header === "chats"),
  );
  const labelAt = headers.findIndex((header) => /month|date|period|yyyy|mm/.test(header));
  const headered = labelAt >= 0 && totalAt >= 0 && partAt >= 0;
  const start = headered ? 1 : 0;
  const points: TrackPoint[] = [];
  for (const row of rows.slice(start)) {
    const point = headered
      ? pointFromCells(row[labelAt] ?? "", row[totalAt] ?? "", row[partAt] ?? "", rateAt >= 0 ? row[rateAt] ?? "" : "")
      : pointFromLine(row.map((cell) => cell.trim()).filter(Boolean).join(" "));
    if (point) points.push(point);
  }
  if (points.length < 3) return null;
  const unit = unitOf(sheetName, rows.flat().join(" "), points);
  return {
    name: headingName(sheetName) || tidyTitle(sheetName) || "Series",
    unit,
    partLabel: "Contained",
    points,
  };
}

function trackFromBody(name: string, body: string): Track | null {
  const points = body
    .split(/\n+/)
    .map((line) => pointFromLine(line.trim()))
    .filter((point): point is TrackPoint => point != null);
  if (points.length < 3) return null;
  const unit = unitOf(name, body, points);
  return {
    name: name || (unit === "chats" ? "Chat" : "Voice"),
    unit,
    partLabel: "Contained",
    points,
  };
}

function pointFromLine(line: string): TrackPoint | null {
  const match = line.match(ROW);
  if (!match) return null;
  return pointFromParts(match[1] ?? "", match[2] ?? "", match[3] ?? "", match[4] ?? "", match[5] ?? "");
}

function pointFromCells(label: string, total: string, part: string, rate: string): TrackPoint | null {
  const match = label.trim().match(new RegExp(`^${MONTH}\\s*[-/\\s]?\\s*(\\d{2,4})$`, "i"));
  if (!match) return null;
  return pointFromParts(match[1] ?? "", match[2] ?? "", total, part, rate);
}

function pointFromParts(month: string, year: string, totalRaw: string, partRaw: string, rateRaw: string): TrackPoint | null {
  const total = count(totalRaw);
  const part = count(partRaw);
  if (total == null || part == null || total <= 0 || part > total * 1.01) return null;
  const stated = count(rateRaw);
  const rate = stated != null && stated <= 100 ? stated : Math.round((part / total) * 100);
  return { label: monthLabel(month, year), total, part: Math.min(part, total), rate };
}

function splitSections(text: string): Section[] {
  const sections: Section[] = [];
  let name = "";
  let body: string[] = [];
  const push = () => {
    const joined = body.join("\n").trim();
    if (joined) sections.push({ name, body: joined });
    body = [];
  };
  for (const line of text.split(/\n/)) {
    const heading = headingName(line);
    // A bare "IVA" line is the sheet name above the voice rows. It stays in that section.
    if (heading && heading !== "IVA") {
      push();
      name = heading;
      continue;
    }
    body.push(line);
  }
  push();
  return sections;
}

function headingName(line: string): string | null {
  const clean = line.trim().replace(/[:.]+$/, "");
  if (/^iva\s*\(voice\)$/i.test(clean) || /^voice$/i.test(clean)) return "Voice";
  if (/^chatbot$/i.test(clean) || /^chat$/i.test(clean)) return "Chat";
  if (/^iva$/i.test(clean)) return "IVA";
  return null;
}

function scopesFrom(text: string, sections: Section[]): ScopeCard[] {
  const cards: ScopeCard[] = [];
  for (const section of sections) {
    const included = section.body.match(/included\s*:\s*([^\n]+)/i)?.[1]?.trim();
    const excluded = section.body.match(/excluded\s*:\s*([^\n]+)/i)?.[1]?.trim();
    if (!included && !excluded) continue;
    const parts = [included, excluded ? `Out: ${excluded}` : ""].filter(Boolean);
    cards.push({ label: section.name || "Population", text: parts.join(" ") });
  }
  const defined = text.match(/containment definition\s*:\s*([^\n]+)/i)?.[1]?.trim();
  const means = text.match(/containment means\s+([^.]+)/i)?.[1]?.trim();
  if (defined || means) {
    const sentence = (defined || means || "").trim().replace(/[.]+$/, "");
    const titled = sentence.charAt(0).toUpperCase() + sentence.slice(1);
    cards.push({ label: "Containment", text: `${titled}.` });
  }
  if (!cards.some((card) => card.label === "Voice")) {
    const voice = text.match(/Voice includes[^.]*\.(?:\s*It excludes[^.]*\.)?/i)?.[0]?.trim();
    if (voice) cards.push({ label: "Voice", text: voice });
  }
  if (!cards.some((card) => card.label === "Chat")) {
    const chat = text.match(/Chat includes[^.]*\./i)?.[0]?.trim();
    if (chat) cards.push({ label: "Chat", text: chat });
  }
  return cards.slice(0, 4);
}

function titleOf(text: string): string {
  const match = text.match(/\b(?:scorecard|deck|presentation|book)\s+for\s+([^\n,.]{2,48})/i);
  const named = match?.[1]?.trim().split(/\s+(?:in|with|and|using|show)\b/i)[0]?.replace(/[.,;:]+$/, "").trim();
  if (named && named.length <= 48 && !/^(the|a|an)\b/i.test(named) && !/^(hi|hey|hello|dear)\b/i.test(named)) return named;
  if (/\bSLT\b/.test(text)) return "SLT";
  const client = text.match(/\bclient\s*:\s*([^\n,]{2,48})/i)?.[1]?.trim().replace(/[.,;:]+$/, "");
  return client || "";
}

function periodField(text: string): string {
  return text.match(/\bperiod\s*:\s*([^\n.]{2,48})/i)?.[1]?.trim().replace(/[.,;:]+$/, "") ?? "";
}

function intentOf(text: string, tracks: Track[]): string {
  if (/includ|exclud|defensible/i.test(text)) {
    return "Show what is included and excluded, then how containment moves month by month.";
  }
  if (tracks.length > 1) return "Show each series from the figures given, month by month, then side by side.";
  return "Show how this series moves across the months given.";
}

function askOf(text: string, scopeFound: boolean): string {
  const line = text.match(/\b(?:recommendation|what to do|next)\s*:\s*([^\n]+)/i)?.[1]?.trim() ?? "";
  const happy = text.match(/\bhappy to\s+([^.]+)/i)?.[1]?.trim();
  if (line && scopeFound && /includes|excludes|containment means/i.test(line)) {
    const extra = line
      .split(/(?<=\.)\s+/)
      .map((part) => part.trim())
      .filter((part) => part && !/includes|excludes|containment means|short month|11 calls/i.test(part));
    if (extra.length) return extra.join(" ");
    return happy ? `${happy.charAt(0).toUpperCase()}${happy.slice(1)}.` : "";
  }
  if (line) return line;
  return happy ? `${happy.charAt(0).toUpperCase()}${happy.slice(1)}.` : "";
}

function subjectFrom(tracks: Track[], text: string): string {
  const bits = tracks.map((track) => {
    const totals = trackTotals(track);
    return `${track.name} ${formatCount(totals.total)} ${track.unit}`;
  });
  const base = `${bits.join(" and ")}, by month.`;
  if (/did not reach an agent|no agent/i.test(text)) return `${base} Containment means no agent was reached.`;
  return base;
}

function periodFromTracks(tracks: Track[]): string {
  const labels = tracks.flatMap((track) => track.points.map((point) => point.label));
  const sorted = [...new Set(labels)].sort((left, right) => monthIndex(left) - monthIndex(right));
  const first = sorted[0];
  const last = sorted[sorted.length - 1];
  if (!first || !last) return "";
  return first === last ? first : `${first} to ${last}`;
}

function unitOf(name: string, body: string, points: TrackPoint[]): string {
  if (/\bchats?\b/i.test(body) || /chat/i.test(name)) return "chats";
  if (/\bcalls?\b/i.test(body) || /voice|call/i.test(name)) return "calls";
  return points.length ? "contacts" : "contacts";
}

function biggestDrop(track: Track, tiny: TrackPoint[]): string {
  let note = "";
  let ratio = 1;
  for (let index = 1; index < track.points.length; index += 1) {
    const previous = track.points[index - 1]!;
    const current = track.points[index]!;
    if (tiny.some((point) => point.label === previous.label || point.label === current.label)) continue;
    if (previous.total <= 0 || current.total >= previous.total * 0.5) continue;
    const nextRatio = current.total / previous.total;
    if (nextRatio < ratio) {
      ratio = nextRatio;
      note = `${current.label} is ${formatCount(current.total)} ${track.unit} after ${formatCount(previous.total)} in ${previous.label}.`;
    }
  }
  return note;
}

function monthLabel(month: string, year: string): string {
  const name = month.slice(0, 3);
  const titled = name.charAt(0).toUpperCase() + name.slice(1).toLowerCase();
  const digits = year.length === 4 ? year.slice(2) : year.padStart(2, "0");
  return `${titled}-${digits}`;
}

function monthSerial(label: string): number | null {
  const index = monthIndex(label);
  return index > 0 ? index : null;
}

function monthIndex(label: string): number {
  const match = label.match(/^([A-Za-z]{3})-(\d{2})$/);
  if (!match) return 0;
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const month = months.indexOf(match[1] ?? "");
  const year = 2000 + Number(match[2] ?? 0);
  return year * 12 + Math.max(month, 0);
}

function count(raw: string): number | null {
  const value = Number(raw.replace(/,/g, "").replace(/%$/, ""));
  if (!Number.isFinite(value) || value < 0) return null;
  return Math.round(value);
}

function tidyTitle(raw: string): string {
  const words = raw.trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return "";
  return words
    .map((word) => {
      if (word.length <= 5 && word === word.toUpperCase() && /[A-Z]/.test(word)) return word;
      if (word === word.toLowerCase()) return word.charAt(0).toUpperCase() + word.slice(1);
      return word;
    })
    .join(" ");
}
