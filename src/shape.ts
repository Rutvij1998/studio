import { decodeBrief } from "./brief";
import { readFacts } from "./facts";
import { readOutline } from "./outline";
import { readMaterial } from "./reading";
import { readStory } from "./story";

const MONTH =
  "Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?";

export const SHAPE_GUIDE = `You rewrite a raw note into one studio prompt. Output only the prompt. No commentary, no markdown fence, no slide list.

Copy every number, percent, and month exactly. Do not add, round, total, average, or explain a change. Do not invent a cause. A greeting such as "Hi Martin" is the audience, not the client. If the note is for the SLT, the client is SLT.

Use one of these shapes.

Monthly rows, one month per line, voice and chat in separate blocks:
Presentation for SLT.
Period: Mar-25 to Oct-26.
Voice
Included: the included sentence.
Excluded: the excluded sentence.
Containment definition: the definition.
Mar-25: 11 calls, 11 contained, 100%

One result, keeping about, above, below, and at least:
Presentation for Boost Mobile. We handled 1.9 million IVA calls this year. Containment is above 70%. Authentication is above 90%.

Named series without months:
Make a deck for SLT. Voice containment is 64% on 1.5 million calls. Chat is 90% on 243 thousand chats.

Claims:
Scorecard for Asurion. 40000 active subscribers. 12400 claims. 6200 were fulfilled. Approval is 9100.

If the note has no figures, still name the client and write:
Presentation for Client.
Recommendation: the ask, in their words.

Do not number the lines. Do not write analysis.`;

/** The readers already have a deck, so a model must not rewrite it. */
export function needsShape(text: string): boolean {
  const clean = text.replace(/\r/g, "").trim();
  if (clean.length < 40) return false;
  if (readMaterial(clean)) return false;
  if (readFacts(clean)) return false;
  if (readStory(clean)) return false;
  const brief = decodeBrief(clean);
  if ((brief.active ?? 0) > 0 || (brief.claims ?? 0) > 0 || (brief.fulfilled ?? 0) > 0 || (brief.approved ?? 0) > 0) return false;
  if ((brief.listening?.length ?? 0) > 0 || (brief.journey?.length ?? 0) > 0 || (brief.months?.length ?? 0) > 0) return false;
  if (readOutline(clean)) return false;
  return true;
}

function scale(raw: string, unit: string | undefined): number | null {
  const value = Number(raw.replace(/,/g, ""));
  if (!Number.isFinite(value)) return null;
  const name = (unit ?? "").toLowerCase();
  const mult = name === "k" || name === "thousand" ? 1e3 : name === "m" || name === "million" ? 1e6 : name === "b" || name === "bn" || name === "billion" ? 1e9 : 1;
  return Math.round(value * mult);
}

function bare(text: string): string {
  return text
    .replace(new RegExp(`\\b(?:${MONTH})\\s*[-/\\s]?\\s*\\d{2,4}\\b`, "gi"), " ")
    .replace(/\b(?:19|20)\d{2}\b/g, " ");
}

function numbersOf(text: string): number[] {
  const found: number[] = [];
  const pattern = /(\d[\d,]*(?:\.\d+)?)(?:\s*(%|percent|million|billion|thousand|bn|m|k|b))?(?![A-Za-z0-9])/gi;
  for (const match of bare(text).matchAll(pattern)) {
    const value = scale(match[1] ?? "", match[2]);
    if (value != null) found.push(value);
  }
  return found;
}

function monthKeys(text: string): string[] {
  const keys: string[] = [];
  const pattern = new RegExp(`\\b(${MONTH})\\s*[-/\\s]?\\s*(\\d{2,4})\\b`, "gi");
  for (const match of text.matchAll(pattern)) {
    const month = (match[1] ?? "").slice(0, 3).toLowerCase();
    let year = Number(match[2]);
    if (year < 100) year += year < 70 ? 2000 : 1900;
    keys.push(`${month}-${year}`);
  }
  return keys.sort();
}

function counts(values: number[]): Map<number, number> {
  const bag = new Map<number, number>();
  for (const value of values) bag.set(value, (bag.get(value) ?? 0) + 1);
  return bag;
}

/** The rewrite kept every figure and every month. Extra repeats of a small rate are allowed. */
export function sameFigures(source: string, shaped: string): boolean {
  const from = numbersOf(source);
  const to = numbersOf(shaped);
  const sourceBag = counts(from);
  const shapedBag = counts(to);
  for (const [value, count] of sourceBag) {
    const next = shapedBag.get(value) ?? 0;
    if (value >= 1000) {
      if (next !== count) return false;
    } else if (next < 1) return false;
  }
  for (const value of shapedBag.keys()) {
    if (!sourceBag.has(value)) return false;
  }
  const sourceMonths = monthKeys(source);
  const shapedMonths = monthKeys(shaped);
  if (sourceMonths.join("|") !== shapedMonths.join("|")) return false;
  return true;
}

export function cleanShape(text: string): string {
  const fenced = text.trim().replace(/^```[a-z]*\s*/i, "").replace(/\s*```$/, "").trim();
  return fenced;
}

export function acceptShape(source: string, shaped: string): string | null {
  const prompt = cleanShape(shaped);
  if (prompt.length < 8 || !sameFigures(source, prompt)) return null;
  return prompt;
}

export function textFromResponse(body: unknown): string {
  if (!body || typeof body !== "object") return "";
  const record = body as {
    output_text?: unknown;
    output?: Array<{ content?: Array<{ type?: string; text?: string }> }>;
  };
  if (typeof record.output_text === "string" && record.output_text.trim()) return record.output_text.trim();
  const parts: string[] = [];
  for (const item of record.output ?? []) {
    for (const content of item.content ?? []) {
      if (content.type === "output_text" && content.text) parts.push(content.text);
    }
  }
  return parts.join("\n").trim();
}
