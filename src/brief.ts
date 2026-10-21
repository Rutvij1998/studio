import { largestRemainder, slideBrief, type Book, type JourneyStage, type ListeningRow, type MonthStack } from "./book";
import type { ListeningChart, MixChart, TypeSize } from "./options";

export type BriefRead = {
  client?: string;
  market?: string;
  period?: string;
  recommendation?: string;
  peerName?: string;
  active?: number;
  claims?: number;
  fulfilled?: number;
  approved?: number;
  peerActive?: number;
  peerClaims?: number;
  peerFulfilled?: number;
  /** Percent of active. Used only when a claim count was not written. */
  claimRate?: number;
  peerClaimRate?: number;
  /** Percent. Of claims when fulfilledOfClaims is set, otherwise of active. */
  fulfilledRate?: number;
  fulfilledOfClaims?: boolean;
  peerFulfilledRate?: number;
  peerFulfilledOfClaims?: boolean;
  listening: ListeningRow[] | null;
  journey: JourneyStage[] | null;
  months: MonthStack[] | null;
  askedListening: boolean;
  askedJourney: boolean;
  askedMonths: boolean;
  askedVersus: boolean;
  mix?: MixChart;
  listeningChart?: ListeningChart;
  type?: TypeSize;
};

const EMPTY: BriefRead = {
  listening: null,
  journey: null,
  months: null,
  askedListening: false,
  askedJourney: false,
  askedMonths: false,
  askedVersus: false,
};

type CountKey = "active" | "claims" | "fulfilled" | "approved" | "peerActive" | "peerClaims" | "peerFulfilled";
type WordKey = "client" | "market" | "period" | "recommendation" | "peerName";

const COUNT_KEYS: CountKey[] = ["active", "claims", "fulfilled", "approved", "peerActive", "peerClaims", "peerFulfilled"];
const WORD_KEYS: WordKey[] = ["client", "market", "period", "recommendation", "peerName"];

export function decodeBrief(text: string): BriefRead {
  const clean = text.replace(/\r/g, "").trim();
  if (!clean) return { ...EMPTY };

  const peerAt = clean.search(/\b(?:versus|vs\.?|compared (?:with|to)|against|next to|peer\b)/i);
  const primary = peerAt >= 0 ? clean.slice(0, peerAt) : clean;
  const peer = peerAt >= 0 ? clean.slice(peerAt) : "";
  const series = parseSeries(clean);
  const fulfilledOfClaims = /(\d+(?:\.\d+)?)\s*%\s*of claims?\s*(?:are\s+)?fulfilled|fulfilled\s*(?:is|are|:)?\s*(\d+(?:\.\d+)?)\s*%\s*of claims/i.test(primary);
  const peerFulfilledOfClaims = /(\d+(?:\.\d+)?)\s*%\s*of claims?\s*(?:are\s+)?fulfilled|fulfilled\s*(?:is|are|:)?\s*(\d+(?:\.\d+)?)\s*%\s*of claims/i.test(peer);

  return {
    client: clientOf(clean),
    market: field(clean, /\bmarket\s*:\s*([^\n,.]{2,32})/i),
    period: periodOf(clean),
    recommendation: recommendationOf(clean),
    peerName: peerNameOf(peer),
    active: count(primary, [new RegExp(`\\bactive(?:\\s+(?:subscribers|subs|book))?\\s*(?::|is|are|=)?\\s*${AMOUNT}\\b`, "i"), new RegExp(`\\b${AMOUNT}\\s+active(?:\\s+(?:subscribers|subs|book))?\\b`, "i")]),
    claims: count(primary, [new RegExp(`\\bclaims?\\s*(?::|is|are|=)\\s*${AMOUNT}\\b`, "i"), new RegExp(`\\b${AMOUNT}\\s+claims?\\b`, "i")]),
    fulfilled: count(primary, [new RegExp(`\\b${AMOUNT}\\s+(?:were|was)\\s+fulfilled\\b`, "i"), new RegExp(`\\bfulfilled(?:\\s+claims)?\\s*(?::|is|are|=)?\\s*${AMOUNT}\\b`, "i"), new RegExp(`\\b${AMOUNT}\\s+fulfilled\\b`, "i")]),
    approved: count(primary, [/(?<!not\s)\b(?:approved|approval)(?:\s+claims)?\s*(?::|is|are|=)\s*(\d[\d,]*(?:\.\d+)?)\s*(k|m|bn|b|million|billion|thousand)?\b/i, /(?<!not\s)\bapproved(?:\s+claims)?\s*(\d[\d,]*(?:\.\d+)?)\s*(k|m|bn|b|million|billion|thousand)?\b/i, /\b(\d[\d,]*(?:\.\d+)?)\s*(k|m|bn|b|million|billion|thousand)?\s+approved\b(?!\s+still)/i]),
    peerActive: count(peer, [new RegExp(`\\bactive(?:\\s+(?:subscribers|subs|book))?\\s*(?::|is|are|=)?\\s*${AMOUNT}\\b`, "i"), new RegExp(`\\b${AMOUNT}\\s+active(?:\\s+(?:subscribers|subs|book))?\\b`, "i")]),
    peerClaims: count(peer, [new RegExp(`\\bclaims?\\s*(?::|is|are|=)\\s*${AMOUNT}\\b`, "i"), new RegExp(`\\b${AMOUNT}\\s+claims?\\b`, "i")]),
    peerFulfilled: count(peer, [new RegExp(`\\b${AMOUNT}\\s+(?:were|was)\\s+fulfilled\\b`, "i"), new RegExp(`\\bfulfilled(?:\\s+claims)?\\s*(?::|is|are|=)?\\s*${AMOUNT}\\b`, "i"), new RegExp(`\\b${AMOUNT}\\s+fulfilled\\b`, "i")]),
    claimRate: percent(primary, [/(\d+(?:\.\d+)?)\s*%\s*(?:claim rate|of active)\b/i, /\bclaim rate\s*(?:of|:|is|are|=)?\s*(\d+(?:\.\d+)?)\s*%/i]),
    peerClaimRate: percent(peer, [/(\d+(?:\.\d+)?)\s*%\s*(?:claim rate|of active)\b/i, /\bclaim rate\s*(?:of|:|is|are|=)?\s*(\d+(?:\.\d+)?)\s*%/i]),
    fulfilledRate: percent(primary, [/(\d+(?:\.\d+)?)\s*%\s*(?:of claims?\s*(?:are\s+)?)?fulfilled/i, /fulfilled\s*(?:at|:|is|are)?\s*(\d+(?:\.\d+)?)\s*%/i]),
    fulfilledOfClaims,
    peerFulfilledRate: percent(peer, [/(\d+(?:\.\d+)?)\s*%\s*(?:of claims?\s*(?:are\s+)?)?fulfilled/i, /fulfilled\s*(?:at|:|is|are)?\s*(\d+(?:\.\d+)?)\s*%/i]),
    peerFulfilledOfClaims,
    listening: series.listening,
    journey: series.journey,
    months: series.months,
    askedListening: series.askedListening || /\b(sentiment|listening)\b/i.test(clean),
    askedJourney: series.askedJourney || /\b(journey|stages?)\b/i.test(clean),
    askedMonths: series.askedMonths || /\b(by month|months|monthly|over time|trend)\b/i.test(clean),
    askedVersus: peerAt >= 0,
    mix: mixOf(clean),
    listeningChart: listeningOf(clean),
    type: typeOf(clean),
  };
}

/** Spreadsheet counts stay. A brief fills only the holes, unless `overwriteWords` is set. */
export function applyBrief(book: Book, decoded: BriefRead, mode: { counts: "fill" | "empty"; words: "fill" | "empty" }): Book {
  const next: Book = { ...book, listening: [...book.listening], journey: [...book.journey], months: [...book.months] };
  for (const key of COUNT_KEYS) {
    const value = decoded[key];
    if (value == null) continue;
    if (mode.counts === "fill" || next[key] === 0) next[key] = value;
  }
  deriveRate(next, "claims", "active", decoded.claimRate, decoded.claims == null, mode.counts);
  deriveRate(next, "peerClaims", "peerActive", decoded.peerClaimRate, decoded.peerClaims == null, mode.counts);
  const fulfilledBase = decoded.fulfilledOfClaims ? "claims" : "active";
  const peerFulfilledBase = decoded.peerFulfilledOfClaims ? "peerClaims" : "peerActive";
  deriveRate(next, "fulfilled", fulfilledBase, decoded.fulfilledRate, decoded.fulfilled == null, mode.counts);
  deriveRate(next, "peerFulfilled", peerFulfilledBase, decoded.peerFulfilledRate, decoded.peerFulfilled == null, mode.counts);
  for (const key of WORD_KEYS) {
    const value = decoded[key]?.trim();
    if (!value) continue;
    if (mode.words === "fill" || !next[key].trim()) next[key] = value;
  }
  if (next.listening.length === 0 && decoded.listening) next.listening = decoded.listening;
  if (next.journey.length === 0 && decoded.journey) next.journey = decoded.journey;
  if (next.months.length === 0 && decoded.months) next.months = decoded.months;
  return next;
}

/** Cover copy keeps the story and leaves chart instructions in the controls. */
export function coverCopy(brief: string, max = 220): string {
  const parts = brief
    .split(/\n+|(?<=[.!?])\s+/)
    .map((part) => part.trim())
    .filter(Boolean);
  const kept = parts.filter((part) => !isControl(part));
  return slideBrief(kept.join(" ") || brief, max);
}

function isControl(part: string): boolean {
  if (/^(recommendation|what to do|next)\s*:/i.test(part)) return true;
  if (/^(market|period|client)\s*:/i.test(part) && part.length < 48) return true;
  if (/\b(donut|doughnut|columns?|meter|bar chart|type size|compact type|large type|small type|regular type)\b/i.test(part) && !/\d{2,}/.test(part)) return true;
  if (/^(show|use|switch)\b/i.test(part) && part.length < 80 && !/\d{2,}/.test(part)) return true;
  if (/^(listening|sentiment|journey|stages|months|by month|monthly|trend)\b/i.test(part)) return true;
  if (/:\s*\d/.test(part) && /\b(negative|positive|neutral|fulfilled|not approved|open)\b/i.test(part)) return true;
  return false;
}

export function briefSummary(input: {
  countsFrom: "sheet" | "words" | "kept" | "none";
  prompt: BriefRead | null;
  document: BriefRead | null;
  book: Book;
}): string {
  const reads = [input.prompt, input.document].filter((read): read is BriefRead => read != null);
  const lines: string[] = [];
  if (input.countsFrom === "sheet") lines.push("Counts come from the spreadsheet.");
  else if (input.countsFrom === "kept") lines.push("Counts already on the book stay.");
  else if (input.countsFrom === "words") lines.push("Counts come from the words you gave.");
  else lines.push("No counts were found, so the rate stays at 0 until figures are added.");

  const mix = input.prompt?.mix ?? input.document?.mix;
  const type = input.prompt?.type ?? input.document?.type;
  const listening = input.prompt?.listeningChart ?? input.document?.listeningChart;
  if (mix) lines.push(`The mix uses ${mix === "columns" ? "columns" : mix === "meter" ? "a meter" : "a donut"}.`);
  if (listening) lines.push(`Listening uses ${listening}.`);
  if (type) lines.push(`Type size is ${type === "l" ? "large" : type === "s" ? "compact" : "regular"}.`);

  const asked = (key: "askedListening" | "askedJourney" | "askedMonths" | "askedVersus") => reads.some((read) => read[key]);
  if (asked("askedListening") && input.book.listening.length === 0) {
    lines.push("Listening was asked for, but no topic counts were written, so that slide stays off.");
  }
  if (asked("askedJourney") && input.book.journey.length === 0) {
    lines.push("A journey was asked for, but no stage split was written, so that slide stays off.");
  }
  if (asked("askedMonths") && input.book.months.length === 0) {
    lines.push("Months were asked for, but no monthly counts were written, so that slide stays off.");
  }
  if (asked("askedVersus") && input.book.peerActive === 0 && input.book.peerClaims === 0) {
    lines.push("A comparison was asked for. Add the peer's active subscribers and claims.");
  }
  if (input.book.listening.length > 0 && reads.some((read) => read.listening)) {
    lines.push(`${input.book.listening.length} listening topics were read from the words.`);
  }
  if (input.book.months.length > 0 && reads.some((read) => read.months)) {
    lines.push(`${input.book.months.length} months were read from the words.`);
  }
  return lines.join(" ");
}

function deriveRate(
  book: Book,
  target: CountKey,
  base: CountKey,
  rate: number | undefined,
  missingCount: boolean,
  mode: "fill" | "empty",
) {
  if (rate == null || !missingCount) return;
  if (mode === "empty" && book[target] !== 0) return;
  if (book[base] <= 0) return;
  book[target] = Math.round((book[base] * rate) / 100);
}

const AMOUNT = "(\\d[\\d,]*(?:\\.\\d+)?)\\s*(k|m|bn|b|million|billion|thousand)?";

function parseCount(raw: string): number | null {
  return scaleCount(raw, undefined);
}

function scaleCount(raw: string, unit: string | undefined): number | null {
  const value = Number(raw.replace(/,/g, ""));
  if (!Number.isFinite(value) || value < 0) return null;
  const name = (unit ?? "").toLowerCase();
  const mult = name === "k" || name === "thousand" ? 1e3 : name === "m" || name === "million" ? 1e6 : name === "b" || name === "bn" || name === "billion" ? 1e9 : 1;
  return Math.round(value * mult);
}

function count(text: string, patterns: RegExp[]): number | undefined {
  for (const pattern of patterns) {
    const flags = pattern.flags.includes("g") ? pattern.flags : `${pattern.flags}g`;
    const global = new RegExp(pattern.source, flags);
    for (const match of text.matchAll(global)) {
      if (!match[1]) continue;
      const after = text.slice((match.index ?? 0) + match[0].length);
      if (/^\s*(?:%|percent\b)/i.test(after)) continue;
      const value = scaleCount(match[1], match[2]);
      if (value != null) return value;
    }
  }
  return undefined;
}

function percent(text: string, patterns: RegExp[]): number | undefined {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (!match) continue;
    const raw = match[1] || match[2];
    if (!raw) continue;
    const value = Number(raw);
    if (Number.isFinite(value) && value >= 0 && value <= 500) return value;
  }
  return undefined;
}

function field(text: string, pattern: RegExp): string | undefined {
  const match = text.match(pattern);
  const value = match?.[1]?.trim().replace(/[.,;:]+$/, "");
  return value || undefined;
}

/** A greeting or a request is the audience, not the client. */
export function isAudienceLine(name: string): boolean {
  const clean = name.trim();
  return /^(hi|hey|hello|dear)\b/i.test(clean) || /^(can you|could you|please|make a|make me|create a|create me|build a|show me)\b/i.test(clean);
}

function clientOf(text: string): string | undefined {
  const match = text.match(/\b(?:scorecard|deck|presentation|book)\s+for\s+(?:the\s+)?([^\n,.]{2,48})/i);
  const named = match?.[1]?.trim().split(/\s+(?:in|with|and|using|show|where|that|about)\b/i)[0]?.replace(/[.,;:]+$/, "").trim();
  if (named && named.length <= 48 && !/^(the|a|an)\b/i.test(named) && !isAudienceLine(named)) return named;
  const explicit = field(text, /\bclient\s*:\s*([^\n,]{2,48})/i);
  if (explicit && !isAudienceLine(explicit)) return explicit;
  const loose = text.match(/\bfor\s+([A-Z][A-Za-z0-9&'’.-]*(?:[ \t]+[A-Z][A-Za-z0-9&'’.-]*){0,3})/);
  const looseName = loose?.[1]?.replace(/[.,;:]+$/, "").trim() ?? "";
  if (!looseName || /^(The|A|An|This|That|Our|Their)$/.test(looseName) || isAudienceLine(looseName)) return undefined;
  return looseName.split(/\s+(?:in|with|and|using|show|where)\b/)[0]?.trim() || undefined;
}

function periodOf(text: string): string | undefined {
  const explicit = field(text, /\bperiod\s*:\s*([^\n.]{2,40})/i);
  if (explicit) return explicit;
  const loose = text.match(/\b(year to date|ytd|q[1-4]\s*\d{4}|fy\s*\d{2,4})\b/i);
  if (!loose) return undefined;
  return loose[1].toLowerCase() === "ytd" ? "Year to date" : loose[1];
}

function recommendationOf(text: string): string | undefined {
  const match = text.match(/\b(?:recommendation|what to do|next)\s*:\s*([^\n]+)/i);
  const value = match?.[1]?.trim();
  if (!value || value.length < 8) return undefined;
  return value.length > 500 ? `${value.slice(0, 500).trimEnd()}…` : value;
}

function peerNameOf(peer: string): string | undefined {
  if (!peer) return undefined;
  const versus = peer.match(/(?:versus|vs\.?|compared (?:with|to)|against|next to)\s+([A-Za-z][A-Za-z0-9&'’.-]*(?:[ \t]+[A-Za-z][A-Za-z0-9&'’.-]*){0,3})/i);
  const named = versus?.[1]?.replace(/[.,;:]+$/, "").trim();
  if (named && named !== named.toLowerCase() && !/^(the|a|an)\b/i.test(named)) return named;
  return field(peer, /\bpeer(?:\s+name)?\s*:\s*([^\n,]{2,40})/i);
}

function mixOf(text: string): MixChart | undefined {
  const pattern = /\b(donut|doughnut|pie chart|pie|ring|meter|one bar|single bar|stacked bar|columns?|bar chart)\b/gi;
  let last: MixChart | undefined;
  for (const match of text.matchAll(pattern)) {
    const word = match[1].toLowerCase();
    const at = match.index ?? 0;
    const around = text.slice(Math.max(0, at - 48), at + match[0].length + 48);
    const aboutSeries = /\b(sentiment|listening|topic|month|monthly|journey)\b/i.test(around);
    const aboutMix = /\b(mix|book|share|split|breakdown)\b/i.test(around);
    if (/donut|doughnut|pie|ring/.test(word)) last = "donut";
    else if (word === "meter" || word === "one bar" || word === "single bar" || word === "stacked bar") {
      if (!aboutSeries || aboutMix) last = "meter";
    } else if ((/column/.test(word) || word === "bar chart") && !(aboutSeries && !aboutMix)) last = "columns";
  }
  return last;
}

function listeningOf(text: string): ListeningChart | undefined {
  if (/\b(sentiment|listening|topics?)\b[^.?\n]{0,40}\bcolumns?\b|\bcolumns?\b[^.?\n]{0,40}\b(sentiment|listening|topics?)\b/i.test(text)) {
    return "columns";
  }
  if (/\b(sentiment|listening|topics?)\b[^.?\n]{0,40}\brows\b|\brows\b[^.?\n]{0,40}\b(sentiment|listening)\b/i.test(text)) return "rows";
  return undefined;
}

function typeOf(text: string): TypeSize | undefined {
  const pattern = /\b(compact|small|smaller|regular|medium|large|bigger|big)\b(?:\s+\w+){0,4}\s+(?:type size|type|text|titles?)|\b(?:type size|type|text|titles?)\b(?:\s+\w+){0,3}\s+(compact|small|smaller|regular|medium|large|bigger|big)\b/gi;
  let last: TypeSize | undefined;
  for (const match of text.matchAll(pattern)) {
    const word = (match[1] || match[2] || "").toLowerCase();
    if (/compact|small/.test(word)) last = "s";
    else if (/regular|medium/.test(word)) last = "m";
    else if (/large|big/.test(word)) last = "l";
  }
  return last;
}

function parseSeries(text: string): Pick<BriefRead, "listening" | "journey" | "months" | "askedListening" | "askedJourney" | "askedMonths"> {
  let mode: "listening" | "journey" | "months" | null = null;
  let askedListening = false;
  let askedJourney = false;
  let askedMonths = false;
  const listening: ListeningRow[] = [];
  const journey: JourneyStage[] = [];
  const months: MonthStack[] = [];
  for (const raw of text.split("\n")) {
    const line = raw.trim();
    if (!line) {
      mode = null;
      continue;
    }
    if (/^(listening|sentiment|mentions)\b/i.test(line)) {
      mode = "listening";
      askedListening = true;
      continue;
    }
    if (/^(journey|stages)\b/i.test(line)) {
      mode = "journey";
      askedJourney = true;
      continue;
    }
    if (/^(months|by month|monthly|trend)\b/i.test(line)) {
      mode = "months";
      askedMonths = true;
      continue;
    }
    if (mode === "listening" || (mode == null && looksLikeListening(line))) {
      const row = parseListening(line);
      if (row && listening.length < 8) {
        listening.push(row);
        askedListening = true;
        continue;
      }
    }
    if (mode === "journey" || (mode == null && /\bmentions?\b/i.test(line))) {
      const stage = parseJourney(line);
      if (stage && journey.length < 6) {
        journey.push(stage);
        askedJourney = true;
        continue;
      }
    }
    if (mode === "months" || (mode == null && /\bfulfilled\b/i.test(line) && /\bnot approved\b|\bopen\b/i.test(line))) {
      const month = parseMonth(line);
      if (month && months.length < 12) {
        months.push(month);
        askedMonths = true;
      }
    }
  }
  return {
    listening: listening.length ? listening : null,
    journey: journey.length ? journey : null,
    months: months.length ? months : null,
    askedListening,
    askedJourney,
    askedMonths,
  };
}

function looksLikeListening(line: string): boolean {
  return /\bneg(?:ative)?\b/i.test(line) && /\bpos(?:itive)?\b/i.test(line);
}

function parseListening(line: string): ListeningRow | null {
  const match = line.match(/^([^:|]{2,40}?)\s*[:|]\s*(.+)$/);
  if (!match) return null;
  const label = match[1].replace(/^[-*]\s*/, "").trim();
  if (!label || /^(listening|sentiment|journey|months)$/i.test(label)) return null;
  const negative = labeled(match[2], "neg(?:ative)?");
  const neutral = labeled(match[2], "neu(?:tral)?");
  const positive = labeled(match[2], "pos(?:itive)?");
  const found = [negative, neutral, positive].filter((value) => value != null).length;
  if (found < 2) return null;
  const row = { label, negative: negative ?? 0, neutral: neutral ?? 0, positive: positive ?? 0 };
  const count = row.negative + row.neutral + row.positive;
  if (count <= 0) return null;
  return { ...row, count };
}

function parseJourney(line: string): JourneyStage | null {
  const match = line.match(/^([^:|]{2,40}?)\s*[:|]\s*(.+)$/);
  if (!match) return null;
  const label = match[1].replace(/^[-*]\s*/, "").trim();
  const negative = labeled(match[2], "neg(?:ative)?");
  const neutral = labeled(match[2], "neu(?:tral)?");
  const positive = labeled(match[2], "pos(?:itive)?");
  if (negative == null || neutral == null || positive == null) return null;
  const mentions = labeled(match[2], "mentions?") ?? negative + neutral + positive;
  const parts = negative + neutral + positive === 100 ? [negative, neutral, positive] : largestRemainder([negative, neutral, positive], negative + neutral + positive);
  return { label, mentions, negative: parts[0] ?? 0, neutral: parts[1] ?? 0, positive: parts[2] ?? 0 };
}

function parseMonth(line: string): MonthStack | null {
  const match = line.match(/^([^:|]{2,24}?)\s*[:|]\s*(.+)$/);
  if (!match) return null;
  const fulfilled = labeled(match[2], "fulfilled");
  const open = labeled(match[2], "still open") ?? labeled(match[2], "open");
  const notApproved = labeled(match[2], "not approved") ?? labeled(match[2], "rejected");
  if (fulfilled == null || (open == null && notApproved == null)) return null;
  return { label: match[1].trim(), fulfilled, open: open ?? 0, notApproved: notApproved ?? 0 };
}

function labeled(text: string, word: string): number | null {
  const lead = text.match(new RegExp(`([\\d,]+(?:\\.\\d+)?)\\s*%?\\s*(?:${word})\\b`, "i"));
  const trail = text.match(new RegExp(`(?:${word})\\s*[:=]?\\s*([\\d,]+(?:\\.\\d+)?)`, "i"));
  const raw = lead?.[1] ?? trail?.[1];
  return raw ? parseCount(raw) : null;
}
