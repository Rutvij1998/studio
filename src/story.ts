import { formatCount, type Book, type Proof, type ProofRelation, type Story } from "./book";

const VERB = "handled|processed|answered|completed|served|supported|resolved|received|ran";
const UNIT_STOP = /^(for|in|where|with|and|this|last|during|above|below|over|under)$/i;
const LEADING = /^(where|the|and|a|an|with|that|our|their|its|this|was|were|is|are)$/i;

/** A claims book keeps its own slides. A story is used only when those counts are absent. */
export function isResults(book: Book): boolean {
  const story = book.story;
  if (!story) return false;
  if (book.active > 0 || book.claims > 0) return false;
  return (story.volume != null && story.volume > 0) || story.proofs.length > 0;
}

export function readStory(text: string): Story | null {
  const clean = text.replace(/\r/g, "").trim();
  if (!clean) return null;
  const volume = volumeFrom(clean);
  const proofs = proofsFrom(clean);
  if (!volume && proofs.length === 0) return null;
  return {
    client: clientFrom(clean),
    period: periodFrom(clean),
    verb: volume?.verb ?? "",
    unit: volume?.unit ?? "",
    volume: volume?.volume ?? null,
    proofs,
  };
}

export function proofPhrase(proof: Proof): string {
  const figure = trimNumber(proof.value);
  if (proof.relation === "above") return `above ${figure}%`;
  if (proof.relation === "below") return `below ${figure}%`;
  if (proof.relation === "least") return `at least ${figure}%`;
  return `${figure}%`;
}

export function storyLine(story: Story, period = story.period): string {
  const parts: string[] = [];
  if (story.volume != null && story.volume > 0 && story.unit) {
    const verb = story.verb ? ` ${story.verb}` : "";
    const when = period ? ` ${periodPhrase(period)}` : "";
    const amount = story.hedge ? `${story.hedge} ${spokenVolume(story.volume)}` : spokenVolume(story.volume);
    parts.push(`${amount} ${story.unit}${verb}${when}`);
  }
  for (const proof of story.proofs) parts.push(`${proof.label} is ${proofPhrase(proof)}`);
  if (parts.length === 0) return "";
  return `${parts.join(". ")}.`;
}

/** Lines that only restate the result. A floor is not turned into an exact rate. */
export function storyReadLines(story: Story, period = story.period): string[] {
  const lines: string[] = [];
  const line = storyLine(story, period);
  if (line) lines.push(line);
  if (story.proofs.some((proof) => proof.relation !== "exact")) {
    lines.push("Above, below, and at least stay floors. The chart stops at the figure in the brief.");
  }
  lines.push("These lines restate the brief. No other figure was added.");
  return lines.slice(0, 6);
}

export function storySummary(story: Story, period = story.period): string {
  const who = story.client.trim() || "this client";
  const lines = [`A results deck for ${who}.`];
  const line = storyLine(story, period);
  if (line) lines.push(line);
  if (story.proofs.some((proof) => proof.relation !== "exact")) {
    lines.push("The bars stop at the figures in the brief. Above means at least that high.");
  }
  return lines.join(" ");
}

export function compactVolume(value: number): string {
  return scaled(value, 1e9, "B", 1) ?? scaled(value, 1e6, "M", 1) ?? scaled(value, 1e3, "K", 0) ?? formatCount(value);
}

export function spokenVolume(value: number): string {
  return (
    scaled(value, 1e9, " billion", 1) ??
    scaled(value, 1e6, " million", 1) ??
    scaled(value, 1e3, " thousand", 0) ??
    formatCount(value)
  );
}

function volumeFrom(text: string): { verb: string; volume: number; unit: string } | null {
  const verbish = text.match(new RegExp(`\\b(${VERB})\\s+([\\d,.]+)\\s*(bn|billion|million|k|m|b)?\\b`, "i"));
  if (verbish?.[2]) {
    const volume = scale(verbish[2], verbish[3]);
    const unit = unitFrom(text.slice((verbish.index ?? 0) + verbish[0].length));
    if (volume != null && unit) return { verb: verbish[1]!.toLowerCase(), volume, unit };
  }
  const bare = text.match(/\b([\d,.]+)\s*(bn|billion|million|k|m|b)\s+/i);
  if (!bare?.[1]) return null;
  const volume = scale(bare[1], bare[2]);
  const unit = unitFrom(text.slice((bare.index ?? 0) + bare[0].length - 1));
  if (volume == null || !unit) return null;
  return { verb: "", volume, unit };
}

function unitFrom(after: string): string {
  const match = after.match(/^\s*([A-Za-z][A-Za-z0-9&'’.-]*(?:\s+[A-Za-z][A-Za-z0-9&'’.-]*){0,4})/);
  if (!match?.[1]) return "";
  const kept: string[] = [];
  for (const word of match[1].split(/\s+/)) {
    if (UNIT_STOP.test(word)) break;
    kept.push(word);
    if (kept.length === 4) break;
  }
  return tidyLabel(kept.join(" "));
}

function proofsFrom(text: string): Proof[] {
  const pattern =
    /\b(?:is|are|was|were)\s+(above|over|greater than|more than|below|under|less than|at least)\s+(\d+(?:\.\d+)?)\s*(?:%|\bpercent\b)|\b(?:is|are|was|were)\s+(\d+(?:\.\d+)?)\s*(?:%|\bpercent\b)/gi;
  const proofs: Proof[] = [];
  for (const match of text.matchAll(pattern)) {
    const value = Number(match[2] || match[3]);
    if (!Number.isFinite(value) || value < 0 || value > 500) continue;
    const before = text.slice(Math.max(0, (match.index ?? 0) - 80), match.index ?? 0);
    const label = labelBefore(before);
    if (!label || proofs.some((proof) => proof.label.toLowerCase() === label.toLowerCase())) continue;
    proofs.push({ label, relation: relationOf(match[1]), value });
    if (proofs.length === 4) break;
  }
  return proofs;
}

function relationOf(word: string | undefined): ProofRelation {
  const name = (word ?? "").toLowerCase();
  if (/above|over|greater|more than/.test(name)) return "above";
  if (/below|under|less than/.test(name)) return "below";
  if (name === "at least") return "least";
  return "exact";
}

function clientFrom(text: string): string {
  const tight = text.match(/\b(?:scorecard|deck|presentation|book)\s+for\s+([^\n,.]{2,48})/i);
  const tightName = tight?.[1]?.trim().split(/\s+(?:in|with|and|using|show)\b/i)[0]?.replace(/[.,;:]+$/, "").trim();
  if (tightName && tightName.length <= 48 && !/^(the|a|an)\b/i.test(tightName)) return tightName;
  const named = text.match(/\bfor\s+([A-Z][A-Za-z0-9&'’.-]*(?:[ \t]+[A-Z][A-Za-z0-9&'’.-]*){0,3})/);
  const name = named?.[1]?.replace(/[.,;:]+$/, "").trim() ?? "";
  if (!name || /^(the|a|an|this|that|our|their)$/i.test(name)) return "";
  return name;
}

function periodFrom(text: string): string {
  const match = text.match(
    /\b(?:in|during|for|over)\s+(this year|last year|this quarter|last quarter|this month|year to date|ytd|q[1-4](?:\s+\d{4})?|fy\s*\d{2,4}|\d{4})\b/i,
  );
  if (!match?.[1]) return "";
  const raw = match[1].trim().toLowerCase();
  if (raw === "ytd" || raw === "year to date") return "Year to date";
  if (raw === "this year") return "This year";
  if (raw === "last year") return "Last year";
  if (raw === "this quarter") return "This quarter";
  if (raw === "last quarter") return "Last quarter";
  if (raw === "this month") return "This month";
  if (/^q[1-4]/.test(raw)) return raw.replace(/^q/, "Q").replace(/\s+/, " ");
  if (raw.startsWith("fy")) return raw.replace(/\s+/g, "").toUpperCase();
  return match[1].trim();
}

function periodPhrase(period: string): string {
  if (/^(this|last)\b/i.test(period)) return period.toLowerCase();
  return `in ${period}`;
}

function labelBefore(before: string): string {
  const words = before.trim().split(/\s+/).filter(Boolean).slice(-6);
  while (words.length > 0 && /^(rate|percent)$/i.test(words[words.length - 1] ?? "")) words.pop();
  const kept: string[] = [];
  for (let index = words.length - 1; index >= 0; index -= 1) {
    const word = words[index] ?? "";
    if (UNIT_STOP.test(word) || /^(where|that|we|showing|show|year)$/i.test(word)) break;
    kept.unshift(word);
    if (kept.length === 3) break;
  }
  return tidyLabel(kept.join(" "));
}

function tidyLabel(raw: string): string {
  const words = raw.trim().split(/\s+/).filter(Boolean);
  while (words.length > 0 && LEADING.test(words[0] ?? "")) words.shift();
  while (words.length > 0 && /^(rate|percent)$/i.test(words[words.length - 1] ?? "")) words.pop();
  if (words.length === 0) return "";
  return words
    .map((word, index) => {
      if (word.length <= 5 && word === word.toUpperCase() && /[A-Z]/.test(word)) return word;
      if (index === 0) return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
      return word.toLowerCase();
    })
    .join(" ");
}

function scale(raw: string, unit: string | undefined): number | null {
  const value = Number(raw.replace(/,/g, ""));
  if (!Number.isFinite(value) || value < 0) return null;
  const name = (unit ?? "").toLowerCase();
  const mult = name === "k" ? 1e3 : name === "m" || name === "million" ? 1e6 : name === "b" || name === "bn" || name === "billion" ? 1e9 : 1;
  return Math.round(value * mult);
}

function trimNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Math.round(value * 10) / 10);
}

function scaled(value: number, unit: number, suffix: string, places: number): string | null {
  if (value < unit) return null;
  const amount = value / unit;
  const factor = 10 ** places;
  if (Math.abs(amount * factor - Math.round(amount * factor)) > 0.05) return null;
  return `${Math.round(amount * factor) / factor}${suffix}`;
}
