import type { Book, Fact, FactGroup, FactRead, ProofRelation, Story } from "./book";
import { proofPhrase, spokenVolume } from "./story";

const VERB = "handled|processed|answered|completed|served|supported|resolved|received|ran";
const MONTH =
  "Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?";

const NUMBER =
  /\b(?:(about|around|roughly|approximately|approx\.?)\s+)?(?:\b(above|over|greater than|more than|below|under|less than|at least)\s+)?(\d[\d,]*(?:\.\d+)?)(?:\s*(%|percent|million|billion|thousand|bn|m|k|b))?(?![A-Za-z0-9])/gi;

type Hit = {
  start: number;
  end: number;
  value: number;
  hedge: string;
  relation: ProofRelation;
  isPercent: boolean;
};

/** A claims book keeps its own slides. Facts are used only when those counts are absent. */
export function isFacts(book: Book): boolean {
  if (!book.facts || book.facts.groups.length === 0) return false;
  if (book.active > 0 || book.claims > 0) return false;
  return true;
}

/**
 * Read a messy Likewize brief one clause at a time.
 * A percent written on a volume stays that percent. Voice and chat are never added.
 */
export function readFacts(text: string): FactRead | null {
  const clean = text.replace(/\r/g, "").trim();
  if (!clean) return null;
  const shared = soleRate(clean);
  let currentRate = "";
  let currentGroup = "";
  const groups: FactGroup[] = [];

  const clauses = clean.split(/\n+|(?<=[.!?;])\s+/);
  let cursor = 0;
  for (const raw of clauses) {
    const clause = raw.trim();
    const startAt = clean.indexOf(clause, cursor);
    const origin = startAt >= 0 ? startAt : cursor;
    cursor = origin + raw.length;
    if (!clause) continue;
    for (const hit of hitsIn(clause)) {
      const before = clause.slice(Math.max(0, hit.start - 100), hit.start);
      const after = clause.slice(hit.end, Math.min(clause.length, hit.end + 40));
      const named = groupName(lastGroup(before));
      if (named) currentGroup = named;
      const group = bucket(groups, currentGroup);
      if (!group) continue;
      if (hit.isPercent) {
        const local = lastRate(before);
        const word = local || currentRate || shared;
        const label = rateLabel(word);
        if (!label) continue;
        if (local) currentRate = local;
        addRate(group, label, hit);
        continue;
      }
      const volumeWord = volumeNear(before, after);
      if (volumeWord) {
        addVolume(group, volumeWord, hit);
        continue;
      }
      const count = countLabel(`${before} ${after}`);
      if (count) addCount(group, count, hit);
    }
  }

  const useful = groups.filter((group) => group.volume || group.rates.length > 0);
  if (useful.length === 0) return null;
  return {
    client: clientFrom(clean),
    period: periodFrom(clean),
    verb: verbFrom(clean),
    groups: useful,
  };
}

/** One unnamed series can use the results slides. Named series stay separate. */
export function storyFromFacts(read: FactRead): Story | null {
  if (read.groups.length !== 1) return null;
  const group = read.groups[0];
  if (!group || group.name) return null;
  if (!group.volume && group.rates.length === 0) return null;
  return {
    client: read.client,
    period: read.period,
    verb: read.verb,
    unit: group.volume?.unit ?? "",
    volume: group.volume?.value ?? null,
    hedge: group.volume?.hedge || undefined,
    proofs: group.rates.slice(0, 4).map((rate) => ({
      label: rate.label,
      relation: rate.relation,
      value: rate.value,
    })),
  };
}

export function factLine(read: FactRead): string {
  const parts = read.groups
    .map((group) => {
      const bits: string[] = [];
      if (group.volume && group.volume.value > 0) bits.push(volumePhrase(group.volume));
      for (const rate of group.rates) bits.push(`${rate.label} is ${ratePhrase(rate)}`);
      const body = bits.join(", ");
      if (!body) return "";
      return group.name ? `${group.name}: ${body}` : body;
    })
    .filter(Boolean);
  if (parts.length === 0) return "";
  return `${parts.join(". ")}.`;
}

/** Lines that only restate the brief. Floors stay floors. Series stay apart. */
export function factReadLines(read: FactRead): string[] {
  const lines: string[] = [];
  for (const group of read.groups) {
    const bits: string[] = [];
    if (group.volume && group.volume.value > 0) bits.push(volumePhrase(group.volume));
    for (const rate of group.rates) bits.push(`${rate.label} is ${ratePhrase(rate)}`);
    if (bits.length === 0) continue;
    lines.push(`${group.name ? `${group.name}: ${bits.join(", ")}` : bits.join(". ")}.`);
  }
  if (read.groups.length > 1) lines.push("Each series stays on its own. Calls and chats are not added together.");
  const floored = read.groups.some(
    (group) => group.rates.some((rate) => rate.relation !== "exact") || (group.volume != null && group.volume.relation !== "exact"),
  );
  if (floored) lines.push("Above, below, and at least stay floors. Nothing is filled in past the brief.");
  lines.push("These lines restate the brief. No other figure was added.");
  return lines.slice(0, 6);
}

export function factsSummary(read: FactRead): string {
  const who = read.client.trim() || "this client";
  const bits = [`Figures for ${who}. ${factLine(read)}`.trim()];
  if (read.groups.length > 1) bits.push("Each series stays on its own slide. Calls and chats are not added together.");
  const hedged = read.groups.some(
    (group) => group.rates.some((rate) => rate.relation !== "exact") || (group.volume != null && group.volume.relation !== "exact"),
  );
  if (hedged) bits.push("The bars stop at the figures in the brief. Above means at least that high.");
  return bits.join(" ");
}

function bucket(groups: FactGroup[], name: string): FactGroup | null {
  const found = groups.find((group) => group.name === name);
  if (found) return found;
  if (groups.length >= 6) return null;
  const next: FactGroup = { name, volume: null, rates: [], counts: [] };
  groups.push(next);
  return next;
}

function addRate(group: FactGroup, label: string, hit: Hit) {
  if (group.rates.length >= 4) return;
  if (group.rates.some((rate) => rate.label.toLowerCase() === label.toLowerCase())) return;
  group.rates.push({
    kind: "rate",
    label,
    value: hit.value,
    relation: hit.relation,
    unit: "%",
    hedge: hit.hedge,
  });
}

function addVolume(group: FactGroup, word: string, hit: Hit) {
  if (group.volume) return;
  const unit = unitLabel(word);
  group.volume = {
    kind: "volume",
    label: unit,
    value: hit.value,
    relation: hit.relation,
    unit,
    hedge: hit.hedge,
  };
}

function addCount(group: FactGroup, label: string, hit: Hit) {
  if (group.counts.length >= 4) return;
  if (group.counts.some((count) => count.label.toLowerCase() === label.toLowerCase())) return;
  group.counts.push({
    kind: "count",
    label,
    value: hit.value,
    relation: hit.relation,
    unit: "",
    hedge: hit.hedge,
  });
}

function hitsIn(clause: string): Hit[] {
  const hits: Hit[] = [];
  for (const match of clause.matchAll(NUMBER)) {
    const raw = match[3];
    if (!raw) continue;
    const unit = (match[4] ?? "").toLowerCase();
    const isPercent = unit === "%" || unit === "percent";
    const numeric = Number(raw.replace(/,/g, ""));
    if (!Number.isFinite(numeric) || numeric < 0) continue;
    const start = match.index ?? 0;
    const end = start + match[0].length;
    if (skipHit(clause, start, numeric, unit, isPercent)) continue;
    const value = isPercent ? numeric : Math.round(numeric * multiplier(unit));
    if (isPercent && value > 500) continue;
    if (!isPercent && value <= 0) continue;
    hits.push({
      start,
      end,
      value,
      hedge: hedgeWord(match[1] ?? ""),
      relation: relationOf(match[2]),
      isPercent,
    });
  }
  return hits;
}

function skipHit(clause: string, start: number, numeric: number, unit: string, isPercent: boolean): boolean {
  if (!isPercent && !unit && numeric >= 1900 && numeric <= 2099 && Number.isInteger(numeric)) return true;
  const before = clause.slice(Math.max(0, start - 8), start);
  if (/[A-Za-z]-$/.test(before) && numeric < 100) return true;
  return false;
}

function soleRate(text: string): string {
  const labels = new Set(
    [...text.matchAll(/\b(containment|authentication|auth|approval)\b/gi)].map((match) => rateLabel(match[1] ?? "")),
  );
  labels.delete("");
  if (labels.size !== 1) return "";
  return [...labels][0] ?? "";
}

function lastRate(text: string): string {
  const matches = [...text.matchAll(/\b(containment|authentication|auth|approval)\b/gi)];
  return matches.length ? (matches[matches.length - 1]?.[1] ?? "").toLowerCase() : "";
}

function rateLabel(word: string): string {
  const name = word.toLowerCase();
  if (name === "containment") return "Containment";
  if (name === "auth" || name === "authentication") return "Authentication";
  if (name === "approval") return "Approval";
  return "";
}

function lastGroup(text: string): string {
  const matches = [...text.matchAll(/\b(voice|chatbot|chats|chat)\b/gi)];
  return matches.length ? (matches[matches.length - 1]?.[1] ?? "").toLowerCase() : "";
}

function groupName(word: string): string {
  const name = word.toLowerCase();
  if (name === "voice") return "Voice";
  if (name === "chat" || name === "chats" || name === "chatbot") return "Chat";
  return "";
}

function volumeNear(before: string, after: string): string {
  return firstVolume(after) || lastVolume(before);
}

function firstVolume(text: string): string {
  const iva = text.match(/\biva\s+calls\b/i);
  if (iva) return "iva calls";
  const match = text.match(/\b(calls|chats|chatbot|repairs|contacts|sessions|tickets)\b/i);
  return match?.[1]?.toLowerCase() ?? "";
}

function lastVolume(text: string): string {
  if (/\biva\s+calls\b/i.test(text)) return "iva calls";
  const matches = [...text.matchAll(/\b(calls|chats|chatbot|repairs|contacts|sessions|tickets)\b/gi)];
  return matches.length ? (matches[matches.length - 1]?.[1] ?? "").toLowerCase() : "";
}

function unitLabel(word: string): string {
  const name = word.toLowerCase();
  if (name.includes("iva")) return "IVA calls";
  if (name.startsWith("call")) return "Calls";
  if (name.startsWith("chat")) return "Chats";
  if (name.startsWith("repair")) return "Repairs";
  if (name.startsWith("contact")) return "Contacts";
  if (name.startsWith("session")) return "Sessions";
  if (name.startsWith("ticket")) return "Tickets";
  return "Count";
}

function countLabel(text: string): string {
  const matches = [
    ...text.matchAll(/\b(active(?:\s+(?:subscribers|subs|book))?|claims|fulfilled|not approved|approved|approval)\b/gi),
  ];
  const last = (matches[matches.length - 1]?.[1] ?? "").toLowerCase();
  if (last.startsWith("active")) return "Active";
  if (last.startsWith("claim")) return "Claims";
  if (last.startsWith("fulfill")) return "Fulfilled";
  if (last.startsWith("not approved")) return "Not approved";
  if (last.startsWith("approv")) return "Approved";
  return "";
}

function volumePhrase(fact: Fact): string {
  const spoken = spokenVolume(fact.value);
  const floor = fact.relation === "least" ? "at least " : fact.relation === "exact" ? "" : `${fact.relation} `;
  const hedge = fact.hedge ? `${fact.hedge} ` : "";
  const unit = fact.unit ? ` ${fact.unit === "IVA calls" ? "IVA calls" : fact.unit.toLowerCase()}` : "";
  return `${hedge}${floor}${spoken}${unit}`;
}

function ratePhrase(fact: Fact): string {
  const phrase = proofPhrase({ label: fact.label, relation: fact.relation, value: fact.value });
  return fact.hedge ? `${fact.hedge} ${phrase}` : phrase;
}

function hedgeWord(raw: string): string {
  return raw.trim() ? "about" : "";
}

function relationOf(word: string | undefined): ProofRelation {
  const name = (word ?? "").toLowerCase();
  if (/above|over|greater|more than/.test(name)) return "above";
  if (/below|under|less than/.test(name)) return "below";
  if (name === "at least") return "least";
  return "exact";
}

function multiplier(unit: string): number {
  const name = unit.toLowerCase();
  if (name === "k" || name === "thousand") return 1e3;
  if (name === "m" || name === "million") return 1e6;
  if (name === "b" || name === "bn" || name === "billion") return 1e9;
  return 1;
}

function clientFrom(text: string): string {
  const deck = text.match(/\b(?:scorecard|deck|presentation|book)\s+for\s+(?:the\s+)?([^\n,.]{2,48})/i);
  const fromDeck = cleanClient(deck?.[1] ?? "");
  if (fromDeck) return fromDeck;
  const named = text.match(/\bfor\s+(?:the\s+)?([A-Z][A-Za-z0-9&'’.-]*(?:[ \t]+[A-Z][A-Za-z0-9&'’.-]*){0,3})/);
  const fromFor = cleanClient(named?.[1] ?? "");
  if (fromFor) return fromFor;
  const subject = text.match(
    new RegExp(`(?:^|[.!?]\\s+)([A-Z][A-Za-z0-9&'’.-]*(?:\\s+[A-Z][A-Za-z0-9&'’.-]*){0,2})\\s+(?:${VERB})\\b`),
  );
  const fromSubject = cleanClient(subject?.[1] ?? "");
  if (fromSubject && !/^(Voice|Chat|Chatbot|Containment|Authentication|Approval|IVA|Make|Create|Show)$/i.test(fromSubject)) {
    return fromSubject;
  }
  return "";
}

function cleanClient(raw: string): string {
  const cut = raw.trim().replace(/[.,;:]+$/, "").split(/\s+(?:in|with|and|using|show|where|during)\b/i)[0]?.trim() ?? "";
  if (!cut || /^(the|a|an|this|that|our|their|we)$/i.test(cut) || cut.length > 48) return "";
  return cut;
}

function periodFrom(text: string): string {
  const range = text.match(
    new RegExp(`\\b(${MONTH})\\.?\\s+(\\d{4})\\s+(?:to|through|thru)\\s+(${MONTH})\\.?\\s+(\\d{4})\\b`, "i"),
  );
  if (range?.[1] && range[2] && range[3] && range[4]) {
    return `${capWord(range[1])} ${range[2]} to ${capWord(range[3])} ${range[4]}`;
  }
  const loose = text.match(/\b(this year|last year|this quarter|last quarter|this month|year to date|ytd)\b/i);
  if (loose?.[1]) return titlePeriod(loose[1]);
  const stamped = text.match(/\b(q[1-4](?:\s+\d{4})?|fy\s*\d{2,4})\b/i);
  if (stamped?.[1]) {
    const raw = stamped[1].trim().toLowerCase();
    if (raw.startsWith("fy")) return raw.replace(/\s+/g, "").toUpperCase();
    return raw.replace(/^q/, "Q").replace(/\s+/, " ");
  }
  const year = text.match(/\b(?:in|during|for|over)\s+(\d{4})\b/);
  if (year?.[1] && Number(year[1]) >= 1900 && Number(year[1]) <= 2099) return year[1];
  return "";
}

function titlePeriod(raw: string): string {
  const name = raw.trim().toLowerCase();
  if (name === "ytd" || name === "year to date") return "Year to date";
  if (name === "this year") return "This year";
  if (name === "last year") return "Last year";
  if (name === "this quarter") return "This quarter";
  if (name === "last quarter") return "Last quarter";
  if (name === "this month") return "This month";
  return raw.trim();
}

function capWord(word: string): string {
  const clean = word.replace(/\./g, "");
  return clean.charAt(0).toUpperCase() + clean.slice(1).toLowerCase();
}

function verbFrom(text: string): string {
  const match = text.match(new RegExp(`\\b(${VERB})\\b`, "i"));
  return match?.[1]?.toLowerCase() ?? "";
}
