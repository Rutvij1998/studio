import JSZip from "jszip";
import type { SlidePiece, SourceChart, SourceDeck, SourceImage, SourceNumber, SourceSeries, SourceSlide, SourceTable } from "./book";
import { chartFit, isTimeLabels } from "./fit";
import { pictureFromModel } from "./seen";

export type PresentKind = "cover" | "chart" | "numbers" | "prose" | "picture";

const STOP = new Set([
  "is", "are", "was", "were", "above", "below", "over", "under", "of", "the", "a", "an",
  "at", "to", "and", "or", "for", "in", "on", "by", "than", "with", "from", "that", "this",
  "we", "our", "be", "been",
]);

const VERB = new Set(["handled", "handle", "handling", "reached", "saw", "made", "took", "had", "has", "have", "posted", "recorded"]);

const NUMBER_RE = /(?<![A-Za-z0-9])(?<currency>[$£€])?\s*(?<num>\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)(?:\s*(?<unit>%|percent|pct|million|billion|thousand|mn|bn|m|k))?(?![A-Za-z0-9])/gi;

const decode = (value: string) =>
  value
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'");

export function isSource(book: { sourceDeck: SourceDeck | null; active: number; claims: number }): boolean {
  return (book.sourceDeck?.slides.length ?? 0) > 0;
}

export function presentKind(slide: SourceSlide): PresentKind {
  if (slide.chart && slide.chart.series.some((series) => series.values.length > 0)) return "chart";
  if (slide.numbers.length > 0 && !slide.title.includes("\n")) return "numbers";
  if (slide.picture) return "picture";
  if (slide.index === 1 && slide.context.trim().length < 320) return "cover";
  return "prose";
}

export function chartKindLabel(kind: SourceChart["kind"]): string {
  if (kind === "columns") return "Columns";
  if (kind === "bars") return "Bars";
  if (kind === "line") return "Line";
  if (kind === "stacked") return "Stacked";
  if (kind === "funnel") return "Funnel";
  return "Donut";
}

export type PlotMode = "score" | "rank" | "matrix" | "trend" | "columns" | "stacked" | "donut" | "funnel";

/** How a chart should be drawn. A short list is a score. Two series are a matrix. A line is a trend. */
export function plotMode(chart: SourceChart): PlotMode {
  if (chart.kind === "line") return "trend";
  if (chart.kind === "donut") return "donut";
  if (chart.kind === "funnel") return "funnel";
  if (chart.kind === "stacked") return "stacked";
  const years = chart.series.length >= 2 && chart.series.every((series) => /^(19|20)\d{2}$/.test(series.name.trim()));
  if (years && chart.categories.length <= 8) return "columns";
  if (chart.series.length >= 2) return "matrix";
  if (chart.categories.length <= 5) return "score";
  if (chart.kind === "bars" || chart.categories.length > 8) return "rank";
  return "columns";
}

/** A heading only when it says more than the title. "Bars" is not a heading. */
export function chartEyebrow(chart: SourceChart): string {
  if (chart.drawn.startsWith("Month over month")) return "Month over month";
  if (chart.drawn.startsWith("Day over day")) return "Day over day";
  if (chart.drawn.startsWith("Year over year")) return "Year over year";
  if (chart.kind === "funnel") return "Funnel";
  if (chart.kind === "donut") return "Share";
  if (chart.kind === "line") return "Trend";
  const names = chart.series
    .map((series) => series.name.trim())
    .filter((name) => name && !/^(count|value|total|series)$/i.test(name));
  if (names.length >= 2) return names.slice(0, 3).join("  ·  ");
  return "";
}

/** Columns in the table that the chart series do not already draw. */
export function spareHeaders(table: SourceTable | null, series: { name: string }[]): string[] {
  if (!table) return [];
  const used = new Set(series.map((item) => item.name.trim()));
  return table.headers.slice(1).map((header) => header.trim()).filter((header) => header && !used.has(header));
}

/** The short name on the slide. Year, month, and day lines say so. */
export function chartKicker(chart: SourceChart): string {
  if (chart.drawn.startsWith("Month over month")) return "Month over month";
  if (chart.drawn.startsWith("Day over day")) return "Day over day";
  if (chart.drawn.startsWith("Year over year")) return "Year over year";
  if (chart.kind === "funnel") return "Funnel";
  if (chart.kind === "donut") return "Donut";
  if (chart.kind === "line") return "Line";
  if (chart.kind === "bars") return "Bars";
  if (chart.kind === "stacked") return "Stacked columns";
  return "Columns";
}

export function formatPoint(value: number): string {
  if (!Number.isFinite(value)) return "";
  if (Math.abs(value - Math.round(value)) < 0.05) return Math.round(value).toLocaleString("en-US");
  return value.toLocaleString("en-US", { maximumFractionDigits: 1 });
}

/** A round ceiling so the axis reads 100, 50, 0 rather than pinning the tallest bar to the top. */
export function axisMax(max: number): number {
  if (!Number.isFinite(max) || max <= 0) return 1;
  const power = 10 ** Math.floor(Math.log10(max));
  const unit = max / power;
  const nice = unit <= 1 ? 1 : unit <= 2 ? 2 : unit <= 5 ? 5 : 10;
  return nice * power;
}

/** Keep a zero baseline, unless every point sits in a narrow band and a zero axis would flatten the line. */
export function axisSpan(values: number[]): { min: number; max: number } {
  const flat = values.filter((value) => Number.isFinite(value)).map((value) => Math.max(0, value));
  const peak = Math.max(0, ...flat);
  const floor = flat.length ? Math.min(...flat) : 0;
  const max = axisMax(peak);
  if (peak > 0 && floor > peak * 0.6 && peak - floor < peak * 0.45) {
    const pad = Math.max(peak * 0.04, (peak - floor) * 0.35);
    const step = max >= 100 ? 10 : max >= 20 ? 5 : 1;
    const min = Math.floor(Math.max(0, floor - pad) / step) * step;
    return { min, max };
  }
  return { min: 0, max };
}

/** Keep the honest footnotes. Drop the sentence that only names the chart type. */
export function chartFoot(chart: SourceChart): string {
  return chart.drawn
    .split(/(?<=\.)\s+/)
    .map((sentence) => sentence.trim())
    .filter((sentence) => /tilde|total stays|first \d+|not drawn|another chart/i.test(sentence))
    .join(" ");
}

export function chartSentence(chart: SourceChart): string {
  const names = chart.series.map((series) => series.name).filter(Boolean);
  const named = names.length ? names.join(", ") : "the series";
  const cats = chart.categories.filter(Boolean);
  const head = cats.slice(0, 4).join(", ");
  const more = cats.length > 4 ? `, and ${cats.length - 4} more` : "";
  const drawn = chart.drawn ? ` ${chart.drawn}` : "";
  if (chart.kind === "donut") {
    const parts = cats.length ? cats.slice(0, 6).join(", ") : named;
    return `Donut of ${parts}.${drawn}`;
  }
  if (chart.kind === "funnel") {
    const parts = cats.length ? cats.slice(0, 6).join(", ") : named;
    return `Funnel of ${parts}.${drawn}`;
  }
  const phrase = chart.kind === "stacked" ? "Stacked columns" : chartKindLabel(chart.kind);
  const across = cats.length ? ` across ${head}${more}` : "";
  return `${phrase} of ${named}${across}.${drawn}`;
}

export function sourceSummary(deck: SourceDeck): string {
  return deck.intent;
}

/** One line for the open slide: where the file leads, then the charts and figures. */
export function slideRead(slide: SourceSlide): string {
  if (slide.chart) {
    const beside = slide.table ? " The table stays beside the chart." : "";
    return `${chartSentence(slide.chart)}${beside}`;
  }
  if (presentKind(slide) === "picture" || slide.images.length > 0) {
    const count = slide.images.length;
    const base = count
      ? `${count} ${count === 1 ? "picture stays" : "pictures stay"} on this slide.`
      : "This slide is a picture. No title, figure, or chart was stored as text.";
    return slide.pace ? `${base} ${slide.pace}` : base;
  }
  if (slide.table && !slide.chart) {
    const headers = slide.table.headers.map((header) => header.trim()).filter(Boolean);
    const line = `Table of ${headers.join(", ") || "the cells"}. The cells stay as written.`;
    return slide.pace ? `${slide.pace} ${line}` : line;
  }
  if ((slide.pieces ?? []).some((piece) => piece.kind === "card" || piece.kind === "list") && slide.images.length === 0) {
    const kept = "The points on this slide, kept as written.";
    return slide.pace ? `${slide.pace} ${kept}` : kept;
  }
  if (slide.numbers.length > 0) {
    const line = `${slide.numbers.map((number) => (number.label ? `${number.label}: ${number.display}` : number.display)).join(". ")}.`;
    return slide.pace ? `${slide.pace} ${line}` : line;
  }
  if (!slide.title.trim() && !slide.context.trim()) return "This slide had no readable text.";
  const words = "The words on this slide, kept as written.";
  return slide.pace ? `${slide.pace} ${words}` : words;
}

export function railLabel(slide: SourceSlide): string {
  const title = slide.title.trim();
  if (title) return title.length > 32 ? `${title.slice(0, 32).trimEnd()}…` : title;
  const kind = presentKind(slide);
  if (kind === "picture") return "Picture";
  if (kind === "chart" && slide.chart) return chartKindLabel(slide.chart.kind);
  if (kind === "numbers") return "Figures";
  return `Slide ${slide.index}`;
}

export function formLabel(slide: SourceSlide): string {
  const kind = presentKind(slide);
  if (kind === "cover") return "Opens";
  if (slide.table && !slide.chart) return "Table";
  if (kind === "chart" && slide.chart) return slide.show === "table" ? "Table" : chartKindLabel(slide.chart.kind);
  if (kind === "picture") return "Picture";
  if ((slide.pieces ?? []).some((piece) => piece.kind === "card" || piece.kind === "list")) return "Context";
  if (kind === "numbers") return "Figures";
  return "Context";
}

function intentOf(slides: SourceSlide[], title: string): string {
  const charts = slides.filter((slide) => slide.chart).length;
  const figures = slides.reduce((total, slide) => total + slide.numbers.length, 0);
  const pictures = slides.filter((slide) => slide.images.length > 0 || presentKind(slide) === "picture").length;
  const bits = [
    `${slides.length} ${slides.length === 1 ? "slide" : "slides"}`,
    `${charts} ${charts === 1 ? "chart" : "charts"}`,
    `${figures} ${figures === 1 ? "figure" : "figures"}`,
  ];
  if (pictures) bits.push(`${pictures} ${pictures === 1 ? "picture" : "pictures"}`);
  const lead = title || slides.find((slide) => slide.title.trim())?.title.trim() || "Untitled";
  const next = slides
    .map((slide) => slide.title.trim())
    .filter((name) => name && name !== lead)
    .slice(0, 3);
  const tail = next.length ? ` Then ${next.join(", ")}.` : "";
  const split = slides.filter((slide) => slide.pace?.includes("Split into")).length;
  const opened = split ? ` ${split} ${split === 1 ? "slide was" : "slides were"} split so each one stays readable.` : "";
  return `Read ${bits.join(", ")}. Every slide is kept. Pictures stay on the slide. Charts are redrawn in the studio colors. A table of numbers is drawn as a chart, and the table stays. Lead: ${lead}.${tail}${opened}`;
}

export type SlideRole = "cover" | "chart" | "figures" | "pictures" | "bullets" | "prose";

const CONTINUED = "Continued.";

function contentLines(slide: SourceSlide): string[] {
  return slide.context.split(/\n/).map((line) => line.trim()).filter((line) => line && line !== CONTINUED);
}

const BULLET_MARK = /^(?:[•●▪‣]|\d+[.)])\s+/;

function markedBullets(lines: string[]): string[] {
  return lines.filter((line) => /^[•●▪‣]\s+/.test(line)).map((line) => line.replace(BULLET_MARK, "").trim()).filter(Boolean);
}

/** A list stays a list. A long paragraph does not get split into fake bullets. */
export function bulletLines(slide: SourceSlide): string[] {
  if (slide.chart || slide.table) return [];
  const lines = contentLines(slide);
  const marked = markedBullets(lines);
  if (marked.length >= 2) return marked;
  const plain = lines.map((line) => line.replace(BULLET_MARK, "").trim()).filter(Boolean);
  if (plain.length < 2 || plain.length > 12) return [];
  if (plain.some((line) => line.length > 110)) return [];
  return plain;
}

/** The sentence above a real list. A list with no lead returns nothing. */
export function leadCopy(slide: SourceSlide): string {
  const lines = contentLines(slide);
  if (markedBullets(lines).length >= 2) return lines.filter((line) => !/^[•●▪‣]\s+/.test(line)).join("\n");
  if (bulletLines(slide).length >= 2) return "";
  return lines.join("\n");
}

export function isContinued(slide: SourceSlide): boolean {
  return slide.context.split("\n")[0]?.trim() === CONTINUED;
}

/** The title slide is dark. Every later slide stays on white. */
export function slideRole(slide: SourceSlide): SlideRole {
  const kind = presentKind(slide);
  if (slide.index === 1 && !slide.chart && slide.images.length === 0 && !slide.table) return "cover";
  if (kind === "chart") return "chart";
  if (kind === "numbers") return "figures";
  if (bulletLines(slide).length >= 2) return "bullets";
  if (kind === "picture" || slide.images.length > 0) return "pictures";
  return "prose";
}

const DATE_TOKEN = /\d{1,2}[/-]\d{1,2}[/-]\d{2,4}/;

/** A trailing date on a title slide is the date line, not part of the title. */
export function coverParts(slide: SourceSlide): { title: string; date: string; line: string } {
  const titleMatch = slide.title.trim().match(new RegExp(`^(.*?)\\s+(${DATE_TOKEN.source})$`));
  const title = (titleMatch ? titleMatch[1] : slide.title).trim() || "Untitled";
  const context = slide.context.trim();
  const contextDate = DATE_TOKEN.test(context) && context.length < 24 ? context : "";
  return {
    title,
    date: titleMatch?.[2] ?? contextDate,
    line: contextDate ? "" : context,
  };
}

function fileTitle(fileName: string): string {
  return fileName.replace(/\.[^.]+$/, "").replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 80);
}

function attr(tag: string, name: string): string {
  return tag.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1] ?? "";
}

function resolvePath(fromFile: string, target: string): string {
  if (target.startsWith("/")) return target.replace(/^\/+/, "");
  const parts = fromFile.split("/").slice(0, -1);
  for (const bit of target.split("/")) {
    if (!bit || bit === ".") continue;
    if (bit === "..") parts.pop();
    else parts.push(bit);
  }
  return parts.join("/");
}

function bulletParagraph(block: string): boolean {
  if (/<a:buNone\b/.test(block)) return false;
  return /<a:bu(?:Char|AutoNum|Blip)\b/.test(block);
}

function textsOf(xml: string, breaks = true, markBullets = false): string {
  const paragraphs = [...xml.matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)];
  const lineOf = (block: string) => {
    const pieces = block.split(/<a:br\b[^>]*\/>/);
    const lines = pieces.map((piece) => {
      const runs = [...piece.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map((match) => decode(match[1] ?? ""));
      let text = "";
      for (const run of runs) {
        if (!run) continue;
        const glue = text && /[\p{L}\p{N}]$/u.test(text) && /^[\p{L}\p{N}]/u.test(run) && !(/\d$/.test(text) && /^\d/.test(run));
        if (glue) text += " ";
        text += run;
      }
      text = text.replace(/[ \t]+/g, " ").trim();
      if (!text) return "";
      return markBullets && bulletParagraph(block) && !/^[•●▪‣]/.test(text) ? `• ${text}` : text;
    }).filter(Boolean);
    return lines.join(breaks ? "\n" : " ");
  };
  if (!paragraphs.length) return lineOf(xml);
  const lines = paragraphs.map((paragraph) => lineOf(paragraph[0])).filter(Boolean);
  return lines.join(breaks ? "\n" : " ");
}

export type CopyBlock = { kind: "prose"; text: string } | { kind: "list"; points: string[] };

/** Keep a lead sentence, then a real list. A line break or a bullet mark starts the next point. */
export function copyBlocks(text: string): CopyBlock[] {
  const source = text.split(/\n/).flatMap((line) => {
    const trimmed = line.trim();
    const marks = trimmed.match(/[•●▪‣]/g)?.length ?? 0;
    if (marks < 2) return [trimmed];
    const parts = trimmed.split(/\s*[•●▪‣]\s*/).map((part) => part.trim()).filter(Boolean);
    if (/^[•●▪‣]/.test(trimmed)) return parts.map((part) => `• ${part}`);
    const [head, ...rest] = parts;
    return [head ?? "", ...rest.map((part) => `• ${part}`)].filter(Boolean);
  }).map((line) => line.trim()).filter((line) => line && line !== CONTINUED);
  const blocks: CopyBlock[] = [];
  let prose: string[] = [];
  let points: string[] = [];
  let bullets: string[] = [];
  const flushProse = () => {
    if (!prose.length) return;
    blocks.push({ kind: "prose", text: prose.join("\n") });
    prose = [];
  };
  const flushPoints = () => {
    if (points.length >= 3) {
      flushProse();
      blocks.push({ kind: "list", points });
    } else if (points.length) prose.push(...points);
    points = [];
  };
  const flushBullets = () => {
    if (!bullets.length) return;
    const items = bullets;
    bullets = [];
    flushPoints();
    if (items.length >= 2) {
      flushProse();
      blocks.push({ kind: "list", points: items });
    } else prose.push(...items);
  };
  for (const line of source) {
    if (/^[•●▪‣]/.test(line)) {
      flushPoints();
      bullets.push(line.replace(/^[•●▪‣]\s*/, "").trim());
      continue;
    }
    flushBullets();
    if (line.length <= 72 && /[.!?]$/.test(line)) {
      points.push(line);
      continue;
    }
    flushPoints();
    prose.push(line);
  }
  flushBullets();
  flushPoints();
  flushProse();
  return blocks;
}

/** The words under a figure, without repeating the number already drawn as the hero. */
export function figureBody(slide: SourceSlide): string {
  const shown = new Set(slide.numbers.map((number) => number.display.trim()).filter(Boolean));
  return slide.context.split("\n").filter((line) => !shown.has(line.trim())).join("\n");
}

type Block = { role: "title" | "body"; text: string };

function blocksOf(xml: string): Block[] {
  const blocks: Block[] = [];
  for (const match of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
    const shape = match[0];
    const role = attr(shape.match(/<p:ph\b[^>]*>/)?.[0] ?? "", "type");
    if (role === "ftr" || role === "dt" || role === "sldNum" || role === "hdr") continue;
    const text = textsOf(shape, true, true);
    if (!text) continue;
    const title = role === "title" || role === "ctrTitle";
    blocks.push({ role: title ? "title" : "body", text });
  }
  return blocks;
}

function wordsOf(text: string): string[] {
  return text.replace(/[.,:;|/]+/g, " ").split(/\s+/).map((word) => word.trim()).filter(Boolean);
}

function trimStops(words: string[]): string[] {
  const next = [...words];
  while (next.length && STOP.has(next[next.length - 1]!.toLowerCase())) next.pop();
  while (next.length && STOP.has(next[0]!.toLowerCase())) next.shift();
  return next;
}

function isWeak(label: string): boolean {
  const words = wordsOf(label);
  return words.length === 0 || words.every((word) => STOP.has(word.toLowerCase()) || VERB.has(word.toLowerCase()));
}

function sameClause(text: string, fromEnd: boolean): string {
  const parts = text.split(/[.!?\n]/).map((part) => part.trim()).filter((part) => part.length > 0);
  if (!parts.length) return "";
  return fromEnd ? parts[parts.length - 1]! : parts[0]!;
}

function cleanLabel(label: string): string {
  return label.replace(/[$£€]?\d[\d,]*(?:\.\d+)?%?/g, " ").replace(/\s+/g, " ").trim();
}

function figuresIn(text: string, title: string): SourceNumber[] {
  const found: SourceNumber[] = [];
  for (const match of text.matchAll(NUMBER_RE)) {
    const raw = match[0] ?? "";
    const num = match.groups?.num ?? "";
    const unit = (match.groups?.unit ?? "").toLowerCase();
    const currency = match.groups?.currency ?? "";
    const start = match.index ?? 0;
    const end = start + raw.length;
    const value = Number(num.replace(/,/g, ""));
    if (!Number.isFinite(value)) continue;
    if (!unit && !currency && value >= 1900 && value <= 2099 && Number.isInteger(value)) continue;
    if (!unit && !currency && value < 10 && Number.isInteger(value) && /^[.)]/.test(text.slice(end, end + 2))) continue;
    if (!unit && !currency && /[/-]/.test(text[end] ?? "") && /\d/.test(text[end + 1] ?? "")) continue;
    if (!unit && !currency && /[/-]/.test(text[start - 1] ?? "") && /\d/.test(text[start - 2] ?? "")) continue;
    if (!unit && !currency && text[end] === "-" && /[A-Za-z]/.test(text[end + 1] ?? "")) continue;
    if (!unit && !currency && Number.isInteger(value) && value < 100 && /step\s*$/i.test(text.slice(Math.max(0, start - 8), start))) continue;
    const tail = text.slice(end, end + 20);
    if (!unit && !currency && /^\s*(?:\+|h\b|hrs?\b|hours?\b|mins?\b|minutes?\b|days?\b)/i.test(tail)) continue;
    if (!unit && !currency && /^\s*(?:to|–|-)\s*\d/.test(tail) && /\b(?:days?|hours?)\b/i.test(text.slice(start, Math.min(text.length, end + 28)))) continue;
    if (!unit && !currency && Number.isInteger(value) && value <= 366 && /\b(?:day|days)\b/i.test(text.slice(Math.max(0, start - 14), Math.min(text.length, end + 10)))) continue;
    const before = trimStops(wordsOf(sameClause(text.slice(0, start), true)).slice(-4));
    const after = trimStops(wordsOf(sameClause(text.slice(end), false)).slice(0, 4));
    let label = cleanLabel(before.join(" "));
    if (!label || VERB.has(label.toLowerCase()) || isWeak(label)) label = cleanLabel(after.join(" "));
    if (isWeak(label)) label = title.trim();
    label = label.replace(/\s+/g, " ").trim().slice(0, 48);
    const scaled = unit === "k" || unit === "thousand"
      ? value * 1_000
      : unit === "m" || unit === "mn" || unit === "million"
        ? value * 1_000_000
        : unit === "bn" || unit === "billion"
          ? value * 1_000_000_000
          : value;
    const suffix = unit === "%" ? "%" : unit === "m" ? "M" : unit === "k" ? "K" : unit ? ` ${unit}` : "";
    const display = `${currency}${num}${suffix}`.replace(/\s+/g, " ").trim();
    found.push({ label, value: scaled, display: display || raw.trim() });
  }
  return found;
}

function cachePoints(block: string): { idx: number; text: string }[] {
  return [...block.matchAll(/<c:pt\b[^>]*\bidx="(\d+)"[^>]*>[\s\S]*?<c:v>([^<]*)<\/c:v>/g)].map((match) => ({
    idx: Number(match[1]),
    text: decode(match[2] ?? "").trim(),
  }));
}

function seriesName(ser: string, index: number, total: number): string {
  const tx = ser.match(/<c:tx>([\s\S]*?)<\/c:tx>/)?.[1] ?? "";
  const cached = cachePoints(tx).map((point) => point.text).filter(Boolean);
  const direct = [...tx.matchAll(/<c:v>([^<]*)<\/c:v>/g)].map((match) => decode(match[1] ?? "").trim()).filter(Boolean);
  const name = (cached[0] || direct[0] || "").trim();
  if (name) return name.slice(0, 48);
  return total > 1 ? `Series ${index + 1}` : "Value";
}

function chartOf(xml: string): SourceChart | null {
  let kind: SourceChart["kind"] | null = null;
  let drawn = "";
  if (/<c:(?:doughnutChart|ofPieChart)\b/.test(xml)) kind = "donut";
  else if (/<c:pieChart\b/.test(xml)) {
    kind = "donut";
    drawn = "Pie, drawn as a donut.";
  } else if (/<c:areaChart\b/.test(xml)) {
    kind = "line";
    drawn = "Area, drawn as a line.";
  } else if (/<c:lineChart\b/.test(xml)) kind = "line";
  else if (/<c:barChart\b/.test(xml)) {
    const dir = xml.match(/<c:barDir\b[^>]*\bval="([^"]+)"/)?.[1] ?? "col";
    const grouping = xml.match(/<c:grouping\b[^>]*\bval="([^"]+)"/)?.[1] ?? "";
    if (grouping === "stacked" || grouping === "percentStacked") kind = "stacked";
    else kind = dir === "bar" ? "bars" : "columns";
  }
  if (!kind) return null;
  const sers = [...xml.matchAll(/<c:ser\b[^>]*>([\s\S]*?)<\/c:ser>/g)].map((match) => match[1] ?? "");
  if (!sers.length) return null;
  const parsed = sers.slice(0, 6).map((ser, index) => {
    const cat = ser.match(/<c:cat>([\s\S]*?)<\/c:cat>/)?.[1] ?? "";
    const val = ser.match(/<c:val>([\s\S]*?)<\/c:val>/)?.[1] ?? "";
    return {
      name: seriesName(ser, index, sers.length),
      cats: cachePoints(cat),
      vals: cachePoints(val),
    };
  });
  const last = Math.max(
    0,
    ...parsed.flatMap((series) => [...series.cats, ...series.vals].map((point) => point.idx)),
  );
  const count = Math.min(24, last + 1);
  if (!count) return null;
  const categories = Array.from({ length: count }, (_, index) => {
    for (const series of parsed) {
      const found = series.cats.find((point) => point.idx === index);
      if (found?.text) return found.text;
    }
    return `${index + 1}`;
  });
  const series: SourceSeries[] = parsed.map((item) => ({
    name: item.name,
    values: Array.from({ length: count }, (_, index) => {
      const found = item.vals.find((point) => point.idx === index);
      const value = found ? Number(found.text.replace(/,/g, "")) : 0;
      return Number.isFinite(value) ? value : 0;
    }),
  })).filter((item) => item.values.some((value) => value !== 0));
  if (!series.length) return null;
  return { kind, drawn, categories, series };
}

function chartTitle(xml: string): string {
  const block = xml.match(/<c:title>([\s\S]*?)<\/c:title>/)?.[1] ?? "";
  return [...block.matchAll(/<c:v>([^<]*)<\/c:v>/g)].map((match) => decode(match[1] ?? "")).join(" ").replace(/\s+/g, " ").trim();
}

function dropEchoes(numbers: SourceNumber[], chart: SourceChart | null): SourceNumber[] {
  if (!chart) return numbers;
  const points = new Set<string>();
  chart.categories.forEach((category, index) => {
    for (const series of chart.series) {
      const value = series.values[index];
      if (value == null || !Number.isFinite(value)) continue;
      points.add(`${category.trim().toLowerCase()}|${value}`);
    }
  });
  const seen = new Set<string>();
  const kept: SourceNumber[] = [];
  for (const number of numbers) {
    const key = `${number.label.trim().toLowerCase()}|${number.value}`;
    if (points.has(key) || seen.has(`${number.display}|${number.label}`)) continue;
    seen.add(`${number.display}|${number.label}`);
    kept.push(number);
  }
  return kept.slice(0, 12);
}

type CellNum = { value: number; approx: boolean };

function parseCell(text: string): CellNum | null {
  const trimmed = text.trim();
  const approx = /^[~≈]/.test(trimmed);
  const body = trimmed.replace(/^[~≈]\s*/, "");
  if (!/^[$£€]?\d{1,3}(?:,\d{3})*(?:\.\d+)?%?$/.test(body) && !/^[$£€]?\d+(?:\.\d+)?%?$/.test(body)) return null;
  const value = Number(body.replace(/[$£€%,]/g, ""));
  if (!Number.isFinite(value)) return null;
  return { value, approx };
}

function tablesOf(xml: string): string[][][] {
  const tables: string[][][] = [];
  for (const table of xml.matchAll(/<a:tbl>([\s\S]*?)<\/a:tbl>/g)) {
    const rows = [...table[1].matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((row) =>
      [...row[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map((cell) => textsOf(cell[0], false).replace(/\s+/g, " ").trim()),
    );
    if (rows.some((row) => row.some(Boolean))) tables.push(rows);
  }
  return tables;
}

function modelTable(raw: string[][]): SourceTable | null {
  if (raw.length < 2) return null;
  const width = Math.max(...raw.map((row) => row.length));
  const rows = raw.map((row) => Array.from({ length: width }, (_, index) => (row[index] ?? "").trim()));
  const numericCount = (row: string[]) => row.filter((cell) => parseCell(cell)).length;
  const firstLabel = !parseCell((rows[0] ?? [])[0] ?? "");
  const belowNumeric = rows.slice(1).filter((row) => row.slice(1).some((cell) => parseCell(cell))).length;
  const yearHeaders = (rows[0] ?? []).slice(1).filter((cell) => {
    const parsed = parseCell(cell);
    return Boolean(parsed && parsed.value >= 1900 && parsed.value <= 2099 && Number.isInteger(parsed.value));
  }).length;
  const textHeader = firstLabel
    && belowNumeric === 0
    && rows.slice(1).length >= 2
    && (rows[0] ?? []).filter((cell) => cell.trim()).length >= 2
    && (rows[0] ?? []).every((cell) => !cell.trim() || (!parseCell(cell) && cell.trim().length <= 42));
  const headered = textHeader || (firstLabel && belowNumeric >= 1 && (numericCount(rows[0] ?? []) === 0 || yearHeaders >= 1));
  const headers = headered
    ? (rows[0] ?? []).map((cell, index) => cell || (index === 0 ? "Item" : `Value ${index}`))
    : (rows[0] ?? []).map((_, index) => (index === 0 ? "Item" : `Value ${index}`));
  const body = (headered ? rows.slice(1) : rows).filter((row) => row.some(Boolean));
  if (body.length < 2) return null;
  return { headers, rows: body };
}

/** A written table whose cells are words, not a chart. */
export function tableHasNumbers(table: SourceTable): boolean {
  return table.rows.some((row) => row.slice(1).some((cell) => parseCell(cell)));
}

function pickTable(tables: string[][][]): SourceTable | null {
  const models = tables.map(modelTable).filter((table): table is SourceTable => Boolean(table));
  if (!models.length) return null;
  const score = (table: SourceTable) => table.rows.reduce((total, row) => total + row.filter((cell) => parseCell(cell)).length, 0);
  return [...models].sort((left, right) => score(right) - score(left))[0] ?? null;
}

const DAY_LABEL = /^(?:mon|tue|wed|thu|fri|sat|sun)\b|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|(?:19|20)\d{2}-\d{2}-\d{2}/i;
const MONTH_LABEL = /^(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b/i;

function shareOf(labels: string[], pattern: RegExp): number {
  if (!labels.length) return 0;
  return labels.filter((label) => pattern.test(label.trim())).length / labels.length;
}

function chartFromTable(table: SourceTable, title: string): SourceChart | null {
  const labels = table.rows.map((row, index) => (row[0] ?? "").trim() || `Row ${index + 1}`);
  const numericCols = table.headers
    .map((_, index) => index)
    .filter((index) => index > 0 && table.rows.filter((row) => parseCell(row[index] ?? "")).length >= Math.min(2, table.rows.length));
  if (!numericCols.length) return null;
  const withoutTotal = numericCols.filter((index) => !/^(total|overall|sum)$/i.test(table.headers[index] ?? ""));
  const used = withoutTotal.length ? withoutTotal : numericCols;
  const droppedTotal = withoutTotal.length > 0 && withoutTotal.length < numericCols.length;
  const dataRows = table.rows
    .map((row, index) => ({ row, label: labels[index] ?? `Row ${index + 1}` }))
    .filter((item) => used.some((index) => parseCell(item.row[index] ?? "")));
  if (dataRows.length < 2) return null;
  const capped = dataRows.slice(0, 16);
  const series = used.map((index) => ({
    name: (table.headers[index] || "Value").slice(0, 48),
    values: capped.map((item) => parseCell(item.row[index] ?? "")?.value ?? 0),
  }));
  const live = series.filter((item) => item.values.some((value) => value !== 0));
  if (!live.length) return null;
  const kept = live.length === series.length || live.length === 0 ? series : series;
  const categories = capped.map((item) => item.label);
  const approx = capped.some((item) => used.some((index) => parseCell(item.row[index] ?? "")?.approx));
  const fitSeries = kept.map((item) => ({ name: item.name, values: item.values }));
  let kind: SourceChart["kind"];
  let sentence: string;
  const yearNames = kept.length >= 2 && kept.every((item) => /^(19|20)\d{2}$/.test(item.name.trim()));
  const one = kept.length === 1 ? kept[0]! : null;
  const decreasing = Boolean(
    one && one.values.length >= 3 && one.values.every((value, index) => index === 0 || value <= (one.values[index - 1] ?? 0)) && (one.values[0] ?? 0) > (one.values[one.values.length - 1] ?? 0),
  );
  const funnelWord = /funnel|conversion/i.test(title) || /funnel|stage/i.test(kept.map((item) => item.name).join(" "));
  if (yearNames) {
    kind = "columns";
    sentence = "Year over year. Each column is a year from the table.";
  } else if (isTimeLabels(categories)) {
    kind = "line";
    if (shareOf(categories, DAY_LABEL) >= 0.6) sentence = "Day over day. The line follows the dates in the table.";
    else if (shareOf(categories, MONTH_LABEL) >= 0.6) sentence = "Month over month. The line follows the months in the table.";
    else sentence = "Drawn as a line. The points run across the periods in the table.";
  } else if (one && decreasing && funnelWord && one.values.every((value) => value >= 0)) {
    kind = "funnel";
    sentence = "Drawn as a funnel. Each step is a row in the table, top to bottom.";
  } else {
    const fitted = chartFit(categories, fitSeries);
    kind = fitted.kind;
    sentence = fitted.sentence;
  }
  if (droppedTotal) sentence += " Total stays in the table.";
  if (approx) sentence += " A tilde in the table is approximate.";
  if (dataRows.length > capped.length) sentence += ` First ${capped.length} of ${dataRows.length} rows are drawn.`;
  return { kind, drawn: sentence, categories, series: kept };
}

/** The cell as written, so a tilde is not dropped when the bar is labeled. */
export function rowDisplay(table: SourceTable | null, category: string, series: SourceSeries[], fallback: string): string {
  if (!table) return fallback;
  const row = table.rows.find((item) => (item[0] ?? "").trim() === category);
  if (!row) return fallback;
  const bits = series.map((item) => {
    const column = table.headers.findIndex((header) => header.trim() === item.name);
    return column >= 0 ? (row[column] ?? "").trim() : "";
  }).filter(Boolean);
  return bits.length ? bits.join(" · ") : fallback;
}

function imageMimeOf(name: string): SourceImage["mime"] | null {
  const lower = name.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".gif")) return "image/gif";
  return null;
}

function asDataUrl(bytes: Uint8Array, mime: SourceImage["mime"]): string {
  let binary = "";
  for (let cursor = 0; cursor < bytes.length; cursor += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(cursor, cursor + 0x8000));
  }
  return `data:${mime};base64,${btoa(binary)}`;
}

/** Every png, jpeg, and gif on the slide. A chart saved as a picture is included. */
async function imagesOn(zip: JSZip, rels: string, xml: string): Promise<SourceImage[]> {
  if (!rels) return [];
  const embeds = new Set([...xml.matchAll(/\br:embed="([^"]+)"/g)].map((match) => match[1] ?? "").filter(Boolean));
  const images: SourceImage[] = [];
  for (const tag of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    if (images.length >= 8) break;
    const id = attr(tag[0], "Id");
    const target = attr(tag[0], "Target");
    const type = attr(tag[0], "Type");
    if (!id || !target || !embeds.has(id) || !/\/image$/i.test(type)) continue;
    const mime = imageMimeOf(target);
    if (!mime) continue;
    const path = resolvePath("ppt/slides/slide.xml", target);
    const file = zip.file(path);
    if (!file) continue;
    const bytes = await file.async("uint8array");
    if (bytes.length < 40 || bytes.length > 8_000_000) continue;
    images.push({ mime, dataUrl: asDataUrl(bytes, mime) });
  }
  return images;
}

function normLine(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

/** A line repeated on most slides, even when it is longer than a classic footer. */
function runningFooter(slides: SourceSlide[]): string {
  const counts = new Map<string, number>();
  for (const slide of slides) {
    const seen = new Set<string>();
    const lines = slide.context.split("\n");
    for (const piece of slide.pieces ?? []) {
      if (piece.kind === "prose") lines.push(...piece.text.split("\n"));
      else if (piece.kind === "card") lines.push(piece.title, ...piece.detail.split("\n"));
    }
    for (const line of lines) {
      const key = normLine(line);
      if (key.length < 28 || key.length > 180 || seen.has(key)) continue;
      seen.add(key);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
  }
  const top = [...counts.entries()].sort((left, right) => right[1] - left[1])[0];
  if (!top || top[1] < Math.max(3, Math.ceil(slides.length * 0.5))) return "";
  return top[0];
}

function withoutLine(text: string, line: string): string {
  if (!line) return text;
  return text.split("\n").map((item) => item.trim()).filter((item) => item && normLine(item) !== line).join("\n");
}

function stripLine(piece: SlidePiece, line: string): SlidePiece | null {
  if (!line) return piece;
  if (piece.kind === "prose") {
    const text = withoutLine(piece.text, line);
    return text ? { kind: "prose", text } : null;
  }
  if (piece.kind === "card") return normLine(piece.title) === line ? null : piece;
  const points = piece.points.filter((point) => normLine(point) !== line);
  if (!points.length) return null;
  return { ...piece, title: normLine(piece.title) === line ? "" : piece.title, points };
}

const STAT_LINE = /^((?:[$£€]\s*)?\d[\d,]*(?:\.\d+)?(?:%|\+|(?:\s*[MKmk]))?)\s+(.+)$/;

function asList(text: string): SlidePiece | null {
  const lines = text.split("\n").map((line) => line.trim()).filter(Boolean);
  const points = lines.filter((line) => /^[•●▪‣]/.test(line)).map((line) => line.replace(/^[•●▪‣]\s*/, "").trim()).filter(Boolean);
  if (points.length < 2) return null;
  return { kind: "list", title: lines.filter((line) => !/^[•●▪‣]/.test(line)).join("\n"), points };
}

function asCard(text: string): SlidePiece | null {
  const lines = text.split("\n").map((line) => line.trim()).filter(Boolean);
  if (!lines.length || lines.length > 4) return null;
  if (lines.some((line) => /^[•●▪‣]/.test(line) || line.length > 120)) return null;
  const head = lines[0] ?? "";
  if (head.length > 52 || /[.!?]$/.test(head) || (lines.length === 1 && isKicker(head))) return null;
  if (lines.length === 1 && head.length > 44) return null;
  let title = head;
  let rest = lines.slice(1);
  const second = rest[0] ?? "";
  if (rest.length >= 2 && second.length <= 32 && !/[.!?]$/.test(second) && !STAT_LINE.test(head)) {
    title = `${head}\n${second}`;
    rest = rest.slice(1);
  }
  const stat = head.match(STAT_LINE);
  return {
    kind: "card",
    title,
    detail: rest.join("\n"),
    display: stat ? (stat[1] ?? "").replace(/\s+/g, "") : "",
    label: stat ? (stat[2] ?? "").trim() : "",
  };
}

function coalesce(pieces: SlidePiece[]): SlidePiece[] {
  const out: SlidePiece[] = [];
  let pending: string[] = [];
  const flush = () => {
    if (pending.length >= 3) out.push({ kind: "list", title: "", points: [...pending] });
    else pending.forEach((text) => out.push({ kind: "prose", text }));
    pending = [];
  };
  for (const piece of pieces) {
    const sentence = piece.kind === "prose" && !piece.text.includes("\n") && piece.text.length <= 90 && /[.!?]$/.test(piece.text);
    if (sentence && piece.kind === "prose") {
      pending.push(piece.text);
      continue;
    }
    flush();
    out.push(piece);
  }
  flush();
  return out;
}

/** Turn each source shape into a card, a list, or a sentence. A bare page number is dropped. */
export function piecesOf(groups: string[]): SlidePiece[] {
  const pieces: SlidePiece[] = [];
  for (const group of groups) {
    const text = group.trim();
    if (!text || /^\d{1,2}$/.test(text)) continue;
    const list = asList(text);
    if (list) {
      pieces.push(list);
      continue;
    }
    const card = asCard(text);
    if (card) {
      pieces.push(card);
      continue;
    }
    pieces.push(...expandProse(text));
  }
  return stackExamples(nameLists(coalesce(joinHeads(pieces))));
}

/** A paragraph that is already a lead plus a list becomes those pieces. One paragraph stays one paragraph. */
function expandProse(text: string): SlidePiece[] {
  const blocks = copyBlocks(text);
  if (blocks.length === 1 && blocks[0]?.kind === "prose") return [{ kind: "prose", text }];
  return blocks.map((block) =>
    block.kind === "list"
      ? { kind: "list" as const, title: "", points: block.points }
      : { kind: "prose" as const, text: block.text },
  );
}

/** "The actual goal" written on the line above its points is the list title. */
function nameLists(pieces: SlidePiece[]): SlidePiece[] {
  const out: SlidePiece[] = [];
  for (let index = 0; index < pieces.length; index += 1) {
    const piece = pieces[index];
    const next = pieces[index + 1];
    if (
      piece?.kind === "prose" &&
      next?.kind === "list" &&
      !next.title.trim() &&
      !piece.text.includes("\n") &&
      piece.text.length <= 60 &&
      !/[.!?]$/.test(piece.text.trim())
    ) {
      out.push({ ...next, title: piece.text.trim() });
      index += 1;
      continue;
    }
    if (piece) out.push(piece);
  }
  return out;
}

function exampleLine(piece: SlidePiece): string {
  if (piece.kind === "prose") {
    const text = piece.text.trim();
    if (!text.includes("\n") && text.length >= 18 && text.length <= 140 && !/[.!?]$/.test(text)) return text;
  }
  if (piece.kind === "card" && !piece.detail && !piece.display && !piece.title.includes("\n")) {
    const text = piece.title.trim();
    if (text.length >= 28 && text.length <= 140 && !laneHead(text) && !/[.!?]$/.test(text)) return text;
  }
  return "";
}

/** Reason-code lines were separate shapes. They read as one list, and short labels stay chips. */
function stackExamples(pieces: SlidePiece[]): SlidePiece[] {
  const out: SlidePiece[] = [];
  let run: SlidePiece[] = [];
  const flush = () => {
    const lines = run.map((piece) => exampleLine(piece)).filter(Boolean);
    const long = lines.filter((line) => line.length >= 28);
    if (lines.length >= 3 && long.length >= 2) out.push({ kind: "list", title: "", points: lines });
    else out.push(...run);
    run = [];
  };
  for (const piece of pieces) {
    if (exampleLine(piece)) {
      run.push(piece);
      continue;
    }
    flush();
    out.push(piece);
  }
  flush();
  return out;
}

type CardPiece = Extract<SlidePiece, { kind: "card" }>;
type ListPiece = Extract<SlidePiece, { kind: "list" }>;

function laneHead(title: string): boolean {
  const text = title.trim();
  return /^[A-Z]{3,12}$/.test(text) || text.includes("·");
}

function sentencePiece(piece: SlidePiece): piece is Extract<SlidePiece, { kind: "prose" }> {
  return piece.kind === "prose" && !piece.text.includes("\n") && piece.text.length <= 220 && /[.!?]["”]?$/.test(piece.text.trim());
}

function detailProse(piece: SlidePiece): piece is Extract<SlidePiece, { kind: "prose" }> {
  if (piece.kind !== "prose") return false;
  const lines = piece.text.split("\n").map((line) => line.trim()).filter(Boolean);
  return lines.length >= 2 && lines.length <= 4 && lines.every((line) => line.length <= 140 && /[.!?]["”]?$/.test(line));
}

function pureStat(text: string): string {
  const match = text.trim().match(/^((?:[$£€]\s*)?\d[\d,]*(?:\.\d+)?(?:%|\+|(?:\s*[MKmk]))?)$/);
  return match ? (match[1] ?? "").replace(/\s+/g, "") : "";
}

/** A heading and the sentence under it were separate shapes. Keep them as one card. */
function foldBundle(bundle: SlidePiece[]): CardPiece {
  const cards = bundle.filter((piece): piece is CardPiece => piece.kind === "card");
  const first = cards[0]!;
  const stat = first.display ? "" : pureStat(first.title);
  let title = first.title;
  let label = first.label;
  const display = first.display || stat;
  const second = cards[1];
  const role = Boolean(second && laneHead(first.title) && !second.title.includes("\n") && second.title.length <= 22 && !/\d/.test(second.title) && !/[,;]/.test(second.title));
  const usedSecond = Boolean(second && ((stat && !second.display && !second.title.includes("\n")) || role));
  if (stat && second && !second.display && !second.title.includes("\n")) {
    label = second.title.trim();
    title = label;
  } else if (role && second) title = `${first.title}\n${second.title.trim()}`;
  const details: string[] = [];
  let skipped = false;
  for (const piece of bundle) {
    if (piece === first) continue;
    if (piece.kind === "card" && usedSecond && !skipped) {
      skipped = true;
      continue;
    }
    if (piece.kind === "card") details.push(piece.title);
    else if (piece.kind === "prose") details.push(piece.text.trim());
  }
  return { kind: "card", title, detail: details.join("\n"), display, label: display ? label : "" };
}

function joinHeads(pieces: SlidePiece[]): SlidePiece[] {
  const out: SlidePiece[] = [];
  let index = 0;
  while (index < pieces.length) {
    const start = pieces[index];
    if (!start || start.kind !== "card" || start.detail) {
      if (start) out.push(start);
      index += 1;
      continue;
    }
    const bundle: SlidePiece[] = [start];
    let cursor = index + 1;
    let sawSentence = false;
    while (cursor < pieces.length && bundle.length < 8) {
      const next = pieces[cursor];
      if (!next || next.kind === "list" || (next.kind === "card" && next.detail)) break;
      if (next.kind === "card") {
        if (laneHead(next.title) || (sawSentence && !laneHead(start.title))) break;
        const ownsNext = laneHead(start.title) || (bundle.length === 1 && Boolean(pureStat(start.title)));
        if (!ownsNext) break;
        bundle.push(next);
        cursor += 1;
        continue;
      }
      if (!sentencePiece(next) && !detailProse(next)) break;
      if (bundle[bundle.length - 1]?.kind === "prose") break;
      bundle.push(next);
      sawSentence = true;
      cursor += 1;
      if (next.text.includes("\n")) break;
    }
    if (!sawSentence && !(laneHead(start.title) && bundle.length > 1)) {
      const follow = pieces[index + 1];
      if (follow?.kind === "list" && !follow.title.trim()) {
        out.push({ ...follow, title: start.title });
        index += 2;
        continue;
      }
      out.push(start);
      index += 1;
      continue;
    }
    out.push(foldBundle(bundle));
    index = cursor;
  }
  return out;
}

export type PieceBand =
  | { kind: "prose"; text: string; height: number }
  | { kind: "lists"; items: ListPiece[]; height: number }
  | { kind: "cards"; items: CardPiece[]; height: number; tone: "stat" | "tiles" | "chips" };

function cardTone(item: CardPiece): "stat" | "tiles" | "chips" {
  if (item.display) return "stat";
  if (!item.detail) return "chips";
  return "tiles";
}

function wrapLines(text: string, width: number): number {
  if (!text.trim()) return 0;
  return text.split("\n").reduce((total, line) => total + Math.max(1, Math.ceil(line.length / width)), 0);
}

export type PieceFit = "read" | "roomy";

type PieceType = {
  title: number;
  lead: number;
  bullet: number;
  stat: number;
  statLabel: number;
  statBody: number;
  cardTitle: number;
  cardBody: number;
  chip: number;
  column: number;
};

/** Reading size is the floor. Roomy is one step up, and only when that step still fits. */
export const PIECE_TYPE: Record<PieceFit, PieceType> = {
  read: { title: 44, lead: 22, bullet: 20, stat: 56, statLabel: 18, statBody: 16, cardTitle: 20, cardBody: 16, chip: 18, column: 22 },
  roomy: { title: 52, lead: 26, bullet: 24, stat: 72, statLabel: 20, statBody: 18, cardTitle: 22, cardBody: 18, chip: 20, column: 26 },
};

const CONTENT = 1168;
const BAND_GAP = 16;
const GRID_GAP = 14;
const FIT_LIMIT = 700;

function charsFor(font: number, width: number): number {
  return Math.max(8, Math.floor(width / (font * 0.56)));
}

function textLines(text: string, font: number, width: number): number {
  return wrapLines(text, charsFor(font, width));
}

function lineStack(lines: number, font: number, leading: number, gap: number): number {
  if (lines <= 0) return 0;
  return lines * Math.ceil(font * leading) + Math.max(0, lines - 1) * gap;
}

function columnWidth(across: number): number {
  if (across <= 1) return CONTENT;
  return (CONTENT - GRID_GAP * (across - 1)) / across;
}

function proseParts(text: string, type: PieceType): Array<{ kind: "prose"; text: string; height: number } | { kind: "list"; points: string[]; height: number }> {
  return copyBlocks(text).map((block) => {
    if (block.kind === "list") {
      const lines = block.points.reduce((count, point) => count + textLines(point, type.bullet, CONTENT - 40), 0);
      return { kind: "list" as const, points: block.points, height: lineStack(lines, type.bullet, 1.4, 8) };
    }
    return { kind: "prose" as const, text: block.text, height: lineStack(textLines(block.text, type.lead, CONTENT), type.lead, 1.4, 4) };
  });
}

function proseHeight(text: string, type: PieceType): number {
  const parts = proseParts(text, type);
  return parts.reduce((total, part, index) => total + part.height + (index ? 10 : 0), 0);
}

function listHeight(items: ListPiece[], type: PieceType): number {
  const width = columnWidth(items.length) - 40;
  const timed = items.filter((item) => isTimeTitle(item.title)).length >= 2;
  return Math.max(...items.map((item) => {
    const titleLines = item.title ? textLines(item.title, type.column, columnWidth(items.length)) : 0;
    const points = item.points.reduce((count, point) => count + textLines(point, type.bullet, width), 0);
    return lineStack(titleLines, type.column, 1.25, 2) + (titleLines ? 10 : 0) + lineStack(points, type.bullet, 1.4, 8);
  })) + (timed ? 40 : 0);
}

function cardHeight(item: CardPiece, type: PieceType, across: number): number {
  const inner = Math.max(80, columnWidth(across) - 36);
  if (!item.detail && !item.display) return 28 + Math.ceil(type.chip * 1.35);
  if (item.display) {
    const detail = textLines(item.detail, type.statBody, inner);
    return 42 + type.stat + lineStack(1, type.statLabel, 1.3, 0) + 8 + lineStack(detail, type.statBody, 1.4, 2);
  }
  const title = textLines(item.title, type.cardTitle, inner);
  const detail = textLines(item.detail, type.cardBody, inner);
  return 36 + lineStack(title, type.cardTitle, 1.3, 2) + 6 + lineStack(detail, type.cardBody, 1.4, 2);
}

/** How many cards share a row. Four tiles are a square. Two fill the row. Chips stay four across. */
export function gridAcross(tone: "stat" | "tiles" | "chips", count: number): number {
  if (count <= 1) return 1;
  if (tone === "chips") return Math.min(4, count);
  if (count === 4 || count === 2) return 2;
  return Math.min(3, count);
}

function cardsHeight(items: CardPiece[], tone: "stat" | "tiles" | "chips", type: PieceType): number {
  const across = gridAcross(tone, items.length);
  let height = 0;
  for (let row = 0; row < items.length; row += across) {
    const slice = items.slice(row, row + across);
    const wide = slice.length;
    const rowH = Math.max(...slice.map((item) => cardHeight(item, type, wide)));
    height += rowH + (height ? GRID_GAP : 0);
  }
  return height;
}

function bandHeight(band: PieceBand, type: PieceType): number {
  if (band.kind === "prose") return proseHeight(band.text, type);
  if (band.kind === "lists") return listHeight(band.items, type);
  return cardsHeight(band.items, band.tone, type);
}

/** The last short paragraph is drawn as a bar, so the fit math keeps that padding. */
function closingLift(band: PieceBand, index: number, count: number): number {
  return index === count - 1 && index > 0 && band.kind === "prose" && band.text.length <= 240 ? 16 : 0;
}

/** Group cards and lists the way the slide will draw them, with a height that must fit 720. */
export function pieceBands(pieces: SlidePiece[]): PieceBand[] {
  const bands: PieceBand[] = [];
  for (let index = 0; index < pieces.length;) {
    const piece = pieces[index];
    if (!piece) break;
    if (piece.kind === "list") {
      let end = index + 1;
      while (end < pieces.length && end - index < 3 && pieces[end]?.kind === "list") end += 1;
      const items = pieces.slice(index, end).filter((item): item is ListPiece => item.kind === "list");
      bands.push({ kind: "lists", items, height: listHeight(items, PIECE_TYPE.read) });
      index = end;
      continue;
    }
    if (piece.kind === "card") {
      const tone = cardTone(piece);
      let end = index + 1;
      while (end < pieces.length) {
        const next = pieces[end];
        if (!next || next.kind !== "card" || cardTone(next) !== tone) break;
        end += 1;
      }
      const items = pieces.slice(index, end).filter((item): item is CardPiece => item.kind === "card");
      bands.push({ kind: "cards", items, height: cardsHeight(items, tone, PIECE_TYPE.read), tone });
      index = end;
      continue;
    }
    bands.push({ kind: "prose", text: piece.text, height: proseHeight(piece.text, PIECE_TYPE.read) });
    index += 1;
  }
  return joinLoneStat(bands);
}

/** One stat sitting between tiles joins that grid, so four diagnosis cards stay a square. */
function joinLoneStat(bands: PieceBand[]): PieceBand[] {
  const out: PieceBand[] = [];
  for (let index = 0; index < bands.length; index += 1) {
    const band = bands[index];
    const prev = out[out.length - 1];
    const next = bands[index + 1];
    if (
      band?.kind === "cards" &&
      band.tone === "stat" &&
      band.items.length === 1 &&
      prev?.kind === "cards" &&
      prev.tone === "tiles" &&
      next?.kind === "cards" &&
      next.tone === "tiles"
    ) {
      const items = [...prev.items, ...band.items, ...next.items];
      out[out.length - 1] = { kind: "cards", tone: "tiles", items, height: cardsHeight(items, "tiles", PIECE_TYPE.read) };
      index += 1;
      continue;
    }
    if (band) out.push(band);
  }
  return out;
}

/** Today on the left and Tomorrow on the right, when both headings are on the slide. */
export function compareSplit(pieces: SlidePiece[]): { before: SlidePiece[]; left: SlidePiece[]; right: SlidePiece[] } | null {
  const leftAt = pieces.findIndex((piece) => piece.kind === "card" && /^today\b/i.test(piece.title.trim()));
  const rightAt = pieces.findIndex((piece) => piece.kind === "card" && /^tomorrow\b/i.test(piece.title.trim()));
  if (leftAt < 0 || rightAt <= leftAt || leftAt > 1) return null;
  return { before: pieces.slice(0, leftAt), left: pieces.slice(leftAt, rightAt), right: pieces.slice(rightAt) };
}

export type CardFace = "" | "green" | "amber" | "red" | "ink";

/** The color is the word they already wrote. Nothing is renamed. */
export function cardFace(title: string): CardFace {
  const head = title.trim().split("\n")[0] ?? "";
  if (/^green\b/i.test(head)) return "green";
  if (/^amber\b/i.test(head)) return "amber";
  if (/^red\b/i.test(head)) return "red";
  if (/^today\b/i.test(head)) return "ink";
  return "";
}

/** Day and week headings are a sequence, in the order they were written. */
export function isTimeTitle(title: string): boolean {
  return /^(?:days?|weeks?)\b/i.test(title.trim());
}

/** One hero figure beside the cards and list that explain it. A row of figures stays a row. */
export function heroAside(pieces: SlidePiece[]): { stat: CardPiece; rest: SlidePiece[] } | null {
  const first = pieces[0];
  if (!first || first.kind !== "card" || !first.display) return null;
  const rest = pieces.slice(1);
  if (rest.length < 1 || rest.length > 4) return null;
  if (rest.some((piece) => piece.kind === "card" && Boolean(piece.display))) return null;
  return { stat: first, rest };
}

/** Pixel height of a card slide at one type step, including the title and padding. */
function stackHeight(pieces: SlidePiece[], type: PieceType): number {
  const bands = pieceBands(pieces);
  let height = 0;
  bands.forEach((band, index) => {
    if (index) height += BAND_GAP;
    const lift = closingLift(band, index, bands.length);
    if (band.kind === "cards") {
      height += band.items.reduce((total, item, itemIndex) => total + cardHeight(item, type, 2) + (itemIndex ? GRID_GAP : 0), 0) + lift;
    } else height += bandHeight(band, type) + lift;
  });
  return height;
}

export function occupy(pieces: SlidePiece[], fit: PieceFit, options: { title: string; continued: boolean; eyebrow?: boolean }): number {
  const type = PIECE_TYPE[fit];
  let height = 40 + 36;
  let bands = 0;
  const add = (block: number) => {
    if (bands) height += BAND_GAP;
    height += block;
    bands += 1;
  };
  if (options.continued) add(22);
  if (options.eyebrow) add(22);
  if (options.title.trim()) add(lineStack(textLines(options.title, type.title, CONTENT), type.title, 1.12, 2));
  const compare = compareSplit(pieces);
  if (compare) {
    const before = pieceBands(compare.before);
    before.forEach((band, index) => add(bandHeight(band, type) + closingLift(band, index, before.length)));
    add(Math.max(stackHeight(compare.left, type), stackHeight(compare.right, type)));
    return height + 12;
  }
  const aside = heroAside(pieces);
  if (aside) {
    const restBands = pieceBands(aside.rest);
    let rest = 0;
    restBands.forEach((band, index) => {
      if (index) rest += BAND_GAP;
      rest += bandHeight(band, type) + closingLift(band, index, restBands.length);
    });
    add(Math.max(cardHeight(aside.stat, type, 2), rest));
    return height + 12;
  }
  const main = pieceBands(pieces);
  main.forEach((band, index) => add(bandHeight(band, type) + closingLift(band, index, main.length)));
  return height + 12;
}

/** Roomy when the larger type still fits. Otherwise the reading size. */
export function pieceFit(slide: Pick<SourceSlide, "pieces" | "title" | "context"> & { eyebrow?: string }): PieceFit {
  const options = { title: slide.title, continued: isContinued(slide as SourceSlide), eyebrow: Boolean(slide.eyebrow) };
  return occupy(slide.pieces ?? [], "roomy", options) <= FIT_LIMIT ? "roomy" : "read";
}

export type Density = "light" | "right" | "full" | "over";

const LIGHT_LIMIT = 460;

function ideaWeight(pieces: SlidePiece[]): number {
  return pieces.reduce((total, piece) => {
    if (piece.kind === "list") return total + piece.points.length + (piece.title ? 1 : 0);
    if (piece.kind === "card") return total + 1;
    return total + Math.max(1, piece.text.split("\n").filter(Boolean).length);
  }, 0);
}

/** How full a card slide is at the reading size and one step up. */
export function pieceDensity(pieces: SlidePiece[], title: string, continued = false): Density {
  const options = { title, continued };
  if (occupy(pieces, "read", options) > FIT_LIMIT) return "over";
  const roomy = occupy(pieces, "roomy", options);
  if (roomy > FIT_LIMIT) return "full";
  if (ideaWeight(pieces) <= 2 && roomy <= LIGHT_LIMIT) return "light";
  return "right";
}

function lightLine(pieces: SlidePiece[]): string {
  const only = pieces.length === 1 ? pieces[0] : undefined;
  const onePoint = only?.kind === "list" && only.points.length === 1 && !only.title;
  const oneLine = only?.kind === "prose" && !only.text.includes("\n");
  return onePoint || oneLine ? "One point. Left large." : "Left large.";
}

function loadPhrase(pieces: SlidePiece[]): string {
  let points = 0;
  let cards = 0;
  let paragraphs = 0;
  for (const piece of pieces) {
    if (piece.kind === "list") points += piece.points.length;
    else if (piece.kind === "card") cards += 1;
    else paragraphs += 1;
  }
  const bits: string[] = [];
  if (points) bits.push(`${points} ${points === 1 ? "point" : "points"}`);
  if (cards) bits.push(`${cards} ${cards === 1 ? "card" : "cards"}`);
  if (paragraphs) bits.push(`${paragraphs} ${paragraphs === 1 ? "paragraph" : "paragraphs"}`);
  return bits.join(" and ");
}

function splitLine(pieces: SlidePiece[], pages: number): string {
  const load = loadPhrase(pieces);
  const noun = `${pages} ${pages === 1 ? "slide" : "slides"}`;
  return load ? `${load}. Split into ${noun}. Nothing was rewritten.` : `Split into ${noun}. Nothing was rewritten.`;
}

export function pieceFrame(fit: PieceFit) {
  const type = PIECE_TYPE[fit];
  return {
    type,
    pt: (px: number) => Math.round(px * 0.75),
    inch: (px: number) => px / 96,
    proseHeight: (text: string) => proseHeight(text, type),
    proseParts: (text: string) => proseParts(text, type),
    listHeight: (items: ListPiece[]) => listHeight(items, type),
    cardHeight: (item: CardPiece, across: number) => cardHeight(item, type, across),
  };
}

function paginate(pieces: SlidePiece[], title: string): SlidePiece[][] {
  const pages: SlidePiece[][] = [];
  let page: SlidePiece[] = [];
  const fits = (items: SlidePiece[], continued: boolean) => occupy(items, "read", { title, continued }) <= FIT_LIMIT;

  const divide = (piece: SlidePiece, continued: boolean): SlidePiece[] => {
    if (piece.kind === "list" && piece.points.length > 1) {
      let count = 1;
      while (count < piece.points.length && fits([{ ...piece, points: piece.points.slice(0, count + 1) }], continued)) count += 1;
      if (count >= piece.points.length) return [piece];
      const rest: ListPiece = { ...piece, points: piece.points.slice(count) };
      return [{ ...piece, points: piece.points.slice(0, count) }, ...divide(rest, true)];
    }
    if (piece.kind === "prose") {
      const lines = piece.text.split("\n").map((line) => line.trim()).filter(Boolean);
      if (lines.length < 2) return [piece];
      let count = 1;
      while (count < lines.length && fits([{ kind: "prose", text: lines.slice(0, count + 1).join("\n") }], continued)) count += 1;
      if (count >= lines.length) return [piece];
      const rest: SlidePiece = { kind: "prose", text: lines.slice(count).join("\n") };
      return [{ kind: "prose", text: lines.slice(0, count).join("\n") }, ...divide(rest, true)];
    }
    return [piece];
  };

  const place = (piece: SlidePiece, allowSplit: boolean) => {
    if (page.length && !fits([...page, piece], pages.length > 0)) {
      pages.push(page);
      page = [];
    }
    if (!page.length && allowSplit && !fits([piece], pages.length > 0)) {
      const parts = divide(piece, pages.length > 0);
      if (parts.length > 1) {
        // A split list stays one column. Side by side would hide the overflow.
        for (const part of parts) {
          if (page.length) {
            pages.push(page);
            page = [];
          }
          page.push(part);
        }
        return;
      }
    }
    page.push(piece);
  };

  for (const piece of pieces) place(piece, true);
  if (page.length) pages.push(page);
  return pages.length ? pages : [[]];
}

/** A bullet wall stored as one paragraph is still a list. Fitting slides keep their original pieces. */
function unpack(pieces: SlidePiece[]): SlidePiece[] {
  const out: SlidePiece[] = [];
  for (const piece of pieces) {
    if (piece.kind !== "prose") {
      out.push(piece);
      continue;
    }
    const blocks = copyBlocks(piece.text);
    if (!blocks.some((block) => block.kind === "list")) {
      out.push(piece);
      continue;
    }
    for (const block of blocks) {
      if (block.kind === "list") out.push({ kind: "list", title: "", points: block.points });
      else if (block.text.trim()) out.push({ kind: "prose", text: block.text });
    }
  }
  return out;
}

/** A slide that cannot fit on one canvas continues on the next one. */
function spreadPieces(slides: SourceSlide[]): SourceSlide[] {
  const out: SourceSlide[] = [];
  for (const slide of slides) {
    const pieces = slide.pieces ?? [];
    if (!pieces.length || slide.chart || slide.table || slide.images.length > 0) {
      out.push(slide);
      continue;
    }
    const units = unpack(pieces);
    const pages = paginate(units, slide.title);
    if (pages.length <= 1) {
      out.push(slide);
      continue;
    }
    pages.forEach((page, index) => {
      out.push({
        ...slide,
        context: index === 0 ? slide.context : CONTINUED,
        pieces: page,
        numbers: index === 0 ? slide.numbers : [],
        images: index === 0 ? slide.images : [],
        pace: index === 0 ? splitLine(units, pages.length) : "Continued. Nothing was rewritten.",
      });
    });
  }
  return reindex(out);
}

export async function readSourceDeck(buffer: ArrayBuffer, fileName = ""): Promise<{ deck: SourceDeck; text: string; note: string }> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    throw new Error("Could not read that PowerPoint. Save it as .pptx.");
  }
  const names = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((left, right) => Number(left.match(/(\d+)\.xml$/i)?.[1] ?? 0) - Number(right.match(/(\d+)\.xml$/i)?.[1] ?? 0));
  const slides: SourceSlide[] = [];
  const markedPicture: boolean[] = [];
  for (const name of names) {
    const xml = await zip.file(name)!.async("string");
    const rels = await zip.file(name.replace("ppt/slides/", "ppt/slides/_rels/").replace(/\.xml$/i, ".xml.rels"))?.async("string");
    const charts: SourceChart[] = [];
    let named = "";
    if (rels) {
      for (const tag of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        const id = attr(tag[0], "Id");
        const target = attr(tag[0], "Target");
        const type = attr(tag[0], "Type");
        if (!id || !target || !/\/chart$/i.test(type)) continue;
        const path = resolvePath(name, target);
        const chartXml = await zip.file(path)?.async("string");
        if (!chartXml) continue;
        if (!named) named = chartTitle(chartXml);
        const chart = chartOf(chartXml);
        if (chart) charts.push(chart);
      }
    }
    const blocks = blocksOf(xml);
    let title = (blocks.find((block) => block.role === "title")?.text ?? "").split("\n")[0]?.trim() ?? "";
    let body = blocks.filter((block) => block.role === "body");
    if (!title) {
      const subject = subjectFrom(body);
      if (subject) {
        title = subject.title;
        body = body.filter((_, index) => !subject.drop.includes(index));
      } else {
        const first = body[0];
        const lines = first?.text.split("\n").map((line) => line.trim()).filter(Boolean) ?? [];
        const head = lines[0] ?? "";
        if (head && head.length <= 80) {
          title = head;
          const restOfFirst = lines.slice(1).join("\n");
          body = [
            ...(restOfFirst ? [{ role: "body" as const, text: restOfFirst }] : []),
            ...body.slice(1),
          ];
        }
      }
    }
    if (!title && named) title = named;
    const table = pickTable(tablesOf(xml));
    let chart: SourceChart | null = charts[0] ?? null;
    if (!chart && table) chart = chartFromTable(table, title);
    if (!title && chart) {
      const seriesName = chart.series.map((series) => series.name).find((item) => item && item !== "Value");
      if (seriesName) title = seriesName;
    }
    const groups = body.map((block) => block.text.trim()).filter(Boolean);
    const pieces = piecesOf(groups);
    const context = groups.filter((group) => !/^\d{1,2}$/.test(group)).join("\n").replace(/\n{3,}/g, "\n\n").trim();
    const images = await imagesOn(zip, rels ?? "", xml);
    markedPicture.push(/<p:pic\b/.test(xml));
    slides.push({
      index: slides.length + 1,
      title: title.slice(0, 140),
      context,
      numbers: [],
      chart,
      extraCharts: Math.max(0, charts.length - 1),
      picture: false,
      images,
      table,
      pieces,
    });
  }
  const footer = runningFooter(slides);
  slides.forEach((slide, index) => {
    if (footer) {
      slide.context = withoutLine(slide.context, footer);
      slide.pieces = (slide.pieces ?? []).flatMap((piece) => {
        const next = stripLine(piece, footer);
        return next ? [next] : [];
      });
    }
    const titleFigures = slide.title.includes("\n") ? [] : figuresIn(slide.title, slide.title);
    slide.numbers = dropEchoes(
      [...titleFigures, ...figuresIn(slide.context, slide.title)].filter((number) => !shownOnCard(slide.pieces ?? [], number)),
      slide.chart,
    );
    const textLen = `${slide.title} ${slide.context}`.trim().length;
    slide.picture = (slide.images.length > 0 || markedPicture[index] === true) && textLen < 12 && !slide.chart && slide.numbers.length === 0;
  });
  const shaped = markPace(spreadPieces(spreadPictures(headlinePass(titlePass(slides)))));
  const title = shaped.find((slide) => slide.title.trim())?.title.trim() || fileTitle(fileName);
  const deck: SourceDeck = { title: title.slice(0, 80), intent: "", slides: shaped };
  deck.intent = shaped.length ? intentOf(shaped, deck.title) : "";
  const text = shaped
    .map((slide) => [slide.title, slide.context].filter(Boolean).join("\n"))
    .filter(Boolean)
    .join("\n\n");
  const charts = shaped.filter((slide) => slide.chart).length;
  const figures = shaped.reduce((total, slide) => total + slide.numbers.length, 0);
  const note = shaped.length
    ? `Read ${shaped.length} ${shaped.length === 1 ? "slide" : "slides"} from ${fileName || "the deck"}. ${charts} ${charts === 1 ? "chart" : "charts"}, ${figures} ${figures === 1 ? "figure" : "figures"}.`
    : `${fileName || "That deck"} had no slides.`;
  return { deck, text, note };
}

function shownOnCard(pieces: SlidePiece[], number: SourceNumber): boolean {
  const display = number.display.replace(/\s+/g, "");
  if (!display) return false;
  return pieces.some((piece) => piece.kind === "card" && piece.display.replace(/\s+/g, "") === display);
}

function isWordmark(line: string): boolean {
  const text = line.trim();
  return text.length >= 3 && text.length <= 22 && !/\s/.test(text) && !/\d/.test(text) && /^[A-Z][A-Za-z'’&.-]*$/.test(text);
}

function isKicker(line: string): boolean {
  const text = line.trim();
  if (text.length < 12 || text.length > 90) return false;
  if (!/[·|]/.test(text) && text.length < 28) return false;
  const letters = text.replace(/[^A-Za-z]/g, "");
  if (letters.length < 8) return false;
  return letters.replace(/[^A-Z]/g, "").length / letters.length >= 0.8;
}

/** A brand word and an all-caps kicker do not beat the subject written under them. */
function subjectFrom(blocks: Block[]): { title: string; drop: number[] } | null {
  const heads = blocks.slice(0, 5).map((block) => block.text.split("\n").map((line) => line.trim()).filter(Boolean));
  for (let index = 0; index < heads.length; index += 1) {
    const lines = heads[index] ?? [];
    const stack = lines.length >= 2 && lines.length <= 3 && lines.every((line) => line.length <= 52 && /[.!?]$/.test(line));
    if (!stack) continue;
    const earlier = heads.slice(0, index);
    const chrome = earlier.every((item) => item.length === 1 && (isWordmark(item[0] ?? "") || isKicker(item[0] ?? "")));
    if (earlier.length && !chrome) return null;
    const drop = [index];
    earlier.forEach((item, place) => {
      if (item.length === 1 && isWordmark(item[0] ?? "")) drop.push(place);
    });
    return { title: lines.join("\n"), drop };
  }
  return null;
}

function personName(text: string): string {
  const line = text.trim();
  if (!line || line.length > 42 || /[\d.!?:,]/.test(line)) return "";
  const words = line.split(/\s+/);
  if (words.length < 2 || words.length > 3) return "";
  if (!words.every((word) => /^[A-Z][a-zA-Z'’-]+$/.test(word))) return "";
  return line;
}

function dateOnly(text: string): boolean {
  return DATE_TOKEN.test(text) && text.trim().length < 24;
}

function reindex(slides: SourceSlide[]): SourceSlide[] {
  return slides.map((slide, index) => ({ ...slide, index: index + 1 }));
}

/** A name on its own, with nothing under it. A label such as "Data Science" is a card, not a presenter. */
function openingName(pieces: SlidePiece[]): string {
  for (const piece of pieces) {
    if (piece.kind === "prose" && personName(piece.text)) return piece.text.trim();
    if (piece.kind === "card" && !piece.detail && !piece.display && !piece.title.includes("\n") && personName(piece.title)) return piece.title.trim();
  }
  return "";
}

function sectionLabel(title: string): boolean {
  const text = title.trim();
  if (!text || text.includes("\n") || text.length > 56 || /[.!?]$/.test(text)) return false;
  const letters = text.replace(/[^A-Za-z]/g, "");
  if (letters.length < 4) return false;
  return letters.replace(/[^A-Z]/g, "").length / letters.length >= 0.8;
}

function headlineLine(piece: SlidePiece): string {
  if (piece.kind === "prose" && !piece.text.includes("\n")) {
    const text = piece.text.trim();
    if (text.length >= 24 && text.length <= 140 && /[a-z]/.test(text) && !isKicker(text)) return text;
  }
  if (piece.kind === "card" && !piece.detail && !piece.display && !piece.title.includes("\n")) {
    const text = piece.title.trim();
    if (text.length >= 24 && text.length <= 90 && /[a-z]/.test(text) && !laneHead(text)) return text;
  }
  return "";
}

/** "80% in 24h, 95% in 48h" stays two figures. The words around them stay. */
function liftRates(pieces: SlidePiece[]): SlidePiece[] {
  const out: SlidePiece[] = [];
  const rate = /^(.*?):\s*(\d+%)\s+in\s+(\d+\s*h),\s*(\d+%)\s+in\s+(\d+\s*h)\.?\s*(.*)$/i;
  for (const piece of pieces) {
    if (piece.kind !== "prose" || !piece.text.split("\n").some((line) => rate.test(line.trim()))) {
      out.push(piece);
      continue;
    }
    const pending: string[] = [];
    const flush = () => {
      if (!pending.length) return;
      out.push({ kind: "prose", text: pending.join("\n") });
      pending.length = 0;
    };
    for (const line of piece.text.split("\n")) {
      const match = line.trim().match(rate);
      if (!match) {
        if (line.trim()) pending.push(line.trim());
        continue;
      }
      flush();
      const head = (match[1] ?? "").trim();
      const within = (hour: string) => `${head} in ${hour.replace(/\s+/g, "")}`;
      out.push({ kind: "card", title: head, detail: "", display: match[2] ?? "", label: within(match[3] ?? "") });
      out.push({ kind: "card", title: head, detail: "", display: match[4] ?? "", label: within(match[5] ?? "") });
      const tail = (match[6] ?? "").trim();
      if (tail) out.push({ kind: "prose", text: tail });
    }
    flush();
  }
  return out;
}

function dropFirstLine(context: string, line: string): string {
  const lines = context.split("\n");
  const at = lines.findIndex((item) => item.trim() === line.trim());
  if (at < 0) return context;
  lines.splice(at, 1);
  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

/** An all-caps section name stays the eyebrow. The sentence under it becomes the title. */
function headlinePass(slides: SourceSlide[]): SourceSlide[] {
  return slides.map((slide) => {
    if (slide.index === 1 || slide.chart) return slide;
    const pieces = slide.pieces ?? [];
    const first = pieces[0];
    if (first?.kind === "prose" && !first.text.includes("\n") && isKicker(first.text.trim())) {
      const line = first.text.trim();
      const next = slide.title.includes("\n") ? liftRates(pieces.slice(1)) : pieces.slice(1);
      return { ...slide, eyebrow: line, pieces: next, context: dropFirstLine(slide.context, line) };
    }
    if (!first || !sectionLabel(slide.title)) return slide;
    const line = headlineLine(first);
    if (!line) return slide;
    return { ...slide, eyebrow: slide.title.trim(), title: line, pieces: pieces.slice(1), context: dropFirstLine(slide.context, line) };
  });
}

/** Slide 1 is the title and the presenter. Any other opening words move to the next slide. */
function titlePass(slides: SourceSlide[]): SourceSlide[] {
  const first = slides[0];
  if (!first) return slides;
  const lines = first.context.split(/\n/).map((line) => line.trim()).filter(Boolean);
  const presenter = openingName(first.pieces ?? []);
  const restLines = lines.filter((line) => line !== presenter && !dateOnly(line));
  const followPieces = (first.pieces ?? []).filter((piece) => {
    if (!presenter) return true;
    if (piece.kind === "prose") return piece.text.trim() !== presenter;
    if (piece.kind === "card") return !(piece.title.trim() === presenter && !piece.detail && !piece.display);
    return true;
  });
  const title = first.title.trim() || "Untitled";
  const opening: SourceSlide = {
    index: 1,
    title,
    context: presenter,
    numbers: [],
    chart: null,
    extraCharts: 0,
    picture: false,
    images: [],
    table: null,
    pieces: [],
  };
  const hasRest = restLines.length > 0 || first.images.length > 0 || Boolean(first.chart) || Boolean(first.table) || first.numbers.length > 0;
  const follow = hasRest ? [{ ...first, title, context: restLines.join("\n"), pieces: followPieces }] : [];
  return reindex([opening, ...follow, ...slides.slice(1)]);
}

/** Pixel size of a PNG, JPEG, or GIF data URL. Returns null when the header is unread. */
export function imageSize(dataUrl: string): { width: number; height: number } | null {
  const payload = dataUrl.slice(dataUrl.indexOf(",") + 1);
  const chars = Math.min(payload.length, 4096);
  const slice = payload.slice(0, chars - (chars % 4));
  let binary = "";
  try {
    binary = atob(slice);
  } catch {
    return null;
  }
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes.length >= 24) {
    const width = (bytes[16]! << 24) | (bytes[17]! << 16) | (bytes[18]! << 8) | bytes[19]!;
    const height = (bytes[20]! << 24) | (bytes[21]! << 16) | (bytes[22]! << 8) | bytes[23]!;
    return width > 0 && height > 0 ? { width, height } : null;
  }
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes.length >= 10) {
    const width = bytes[6]! | (bytes[7]! << 8);
    const height = bytes[8]! | (bytes[9]! << 8);
    return width > 0 && height > 0 ? { width, height } : null;
  }
  if (bytes[0] === 0xff && bytes[1] === 0xd8) {
    let cursor = 2;
    while (cursor + 9 < bytes.length) {
      if (bytes[cursor] !== 0xff) break;
      const marker = bytes[cursor + 1] ?? 0;
      if (marker === 0xc0 || marker === 0xc1 || marker === 0xc2) {
        const height = ((bytes[cursor + 5] ?? 0) << 8) | (bytes[cursor + 6] ?? 0);
        const width = ((bytes[cursor + 7] ?? 0) << 8) | (bytes[cursor + 8] ?? 0);
        return width > 0 && height > 0 ? { width, height } : null;
      }
      const size = ((bytes[cursor + 2] ?? 0) << 8) | (bytes[cursor + 3] ?? 0);
      if (size < 2) break;
      cursor += 2 + size;
    }
  }
  return null;
}

export type PictureLayout = "solo" | "pair" | "stack" | "row" | "grid";

/** How a handful of pictures share one slide. Wide strips stack. Tall shots sit in a row. */
export function pictureLayout(images: Array<{ dataUrl: string }>): PictureLayout {
  const shown = images.slice(0, 4);
  if (shown.length <= 1) return "solo";
  const ratios = shown.map((image) => {
    const size = imageSize(image.dataUrl);
    return size && size.height > 0 ? size.width / size.height : 1.4;
  });
  const strip = ratios.every((ratio) => ratio >= 2.4);
  const tall = ratios.every((ratio) => ratio <= 1.15);
  if (shown.length === 2) return strip ? "stack" : "pair";
  if (shown.length === 3) {
    if (strip) return "stack";
    if (tall) return "row";
    return "grid";
  }
  return "grid";
}

/** Four pictures fit on one white slide. More than that continues on the next slide. */
function spreadPictures(slides: SourceSlide[]): SourceSlide[] {
  const out: SourceSlide[] = [];
  for (const slide of slides) {
    if (slide.images.length <= 4) {
      out.push(slide);
      continue;
    }
    const mid = Math.ceil(slide.images.length / 2);
    const lines = slide.context.split(/\n/).map((line) => line.trim()).filter(Boolean);
    const cut = lines.length >= 2 ? Math.ceil(lines.length / 2) : lines.length;
    const firstLines = lines.slice(0, cut);
    const secondLines = lines.slice(cut);
    const count = slide.images.length;
    out.push({
      ...slide,
      images: slide.images.slice(0, mid),
      context: firstLines.join("\n"),
      pace: `${count} pictures. Split into 2 slides. The pictures stay.`,
    });
    out.push({
      ...slide,
      images: slide.images.slice(mid, mid + 4),
      context: secondLines.length ? [CONTINUED, ...secondLines].join("\n") : `${CONTINUED}\nThese pictures follow the ones on the previous slide.`,
      chart: null,
      table: null,
      numbers: [],
      extraCharts: 0,
      pace: "Continued. The pictures stay.",
    });
  }
  return reindex(out);
}

/** Name a light slide and a slide that only fits at the reading size. Splits already carry their own note. */
function markPace(slides: SourceSlide[]): SourceSlide[] {
  return slides.map((slide) => {
    if (slide.pace) return slide;
    if (slide.index === 1 && !slide.chart && slide.images.length === 0 && !slide.table) return slide;
    if (slide.chart) return slide;
    const pieces = slide.pieces ?? [];
    if (pieces.length && slide.images.length === 0) {
      const band = pieceDensity(pieces, slide.title, isContinued(slide));
      if (band === "light") return { ...slide, pace: lightLine(pieces) };
      if (band === "full") return { ...slide, pace: "Fits at the reading size." };
      if (band === "over") return { ...slide, pace: "Too full for one slide. The words stay as written." };
      return slide;
    }
    const words = `${slide.title} ${slide.context}`.replace(/\s+/g, " ").trim();
    if (slide.images.length === 1 && words.length < 80) return { ...slide, pace: "Left large." };
    return slide;
  });
}

export type SlideShow = "chart" | "table" | "both";

export type SlideEdit = {
  slides: SourceSlide[];
  /** 0-based slide to show after the edit. */
  index: number;
  note: string;
};

/** Both when the file has a chart and a table. A chosen show wins. */
export function slideShow(slide: Pick<SourceSlide, "show" | "chart" | "table">): SlideShow {
  if (slide.show) return slide.show;
  if (slide.chart && slide.table) return "both";
  if (slide.table && !slide.chart) return "table";
  return "chart";
}

/** A native chart can flip. A picture of a chart cannot. */
export function canFlip(slide: SourceSlide): boolean {
  return Boolean(slide.chart && slide.chart.series.some((series) => series.values.length > 0));
}

/** The written table, or the chart's own cells when the file stored no table. */
export function tableFromSlide(slide: SourceSlide): SourceTable | null {
  if (slide.table && slide.table.rows.length > 0) return slide.table;
  const chart = slide.chart;
  if (!chart || chart.categories.length === 0 || chart.series.length === 0) return null;
  return {
    headers: ["", ...chart.series.map((series) => series.name.trim() || "Value")],
    rows: chart.categories.map((category, index) => [
      category,
      ...chart.series.map((series) => formatPoint(series.values[index] ?? 0)),
    ]),
  };
}

export function replaceSlides(deck: SourceDeck, slides: SourceSlide[]): SourceDeck {
  const next = reindex(slides);
  return { ...deck, slides: next, intent: next.length ? intentOf(next, deck.title) : "" };
}

function locate(slides: SourceSlide[], index: number): number {
  const found = slides.findIndex((slide) => slide.index === index);
  return found >= 0 ? found : Math.min(Math.max(slides.length - 1, 0), Math.max(0, index - 1));
}

function clonePiece(piece: SlidePiece): SlidePiece {
  if (piece.kind === "list") return { ...piece, points: [...piece.points] };
  return { ...piece };
}

function cloneSlide(slide: SourceSlide): SourceSlide {
  return {
    ...slide,
    numbers: slide.numbers.map((item) => ({ ...item })),
    chart: slide.chart
      ? {
          ...slide.chart,
          categories: [...slide.chart.categories],
          series: slide.chart.series.map((series) => ({ ...series, values: [...series.values] })),
        }
      : null,
    images: slide.images.map((image) => ({ ...image })),
    table: slide.table ? { headers: [...slide.table.headers], rows: slide.table.rows.map((row) => [...row]) } : null,
    pieces: slide.pieces?.map(clonePiece),
    pace: undefined,
  };
}

function bodyPieces(slide: SourceSlide): SlidePiece[] {
  if ((slide.pieces ?? []).length) return (slide.pieces ?? []).map(clonePiece);
  const text = slide.context.trim();
  return text ? [{ kind: "prose", text }] : [];
}

function reshape(slide: SourceSlide): SourceSlide[] {
  return markPace(spreadPieces(spreadPictures([slide])));
}

function kept(slides: SourceSlide[], index: number, note: string): SlideEdit {
  return { slides, index: Math.max(0, Math.min(index, slides.length - 1)), note };
}

/** The title slide stays first. Other slides trade places. */
export function moveSlide(slides: SourceSlide[], index: number, direction: -1 | 1): SlideEdit {
  const at = locate(slides, index);
  const to = at + direction;
  if (at <= 0 || to <= 0 || to >= slides.length) return kept(slides, at, "The title slide stays first.");
  const next = slides.slice();
  const [item] = next.splice(at, 1);
  next.splice(to, 0, item!);
  return { slides: next, index: to, note: direction < 0 ? "Moved up." : "Moved down." };
}

/** A copy is inserted after the slide. The copy is not a second title slide. */
export function duplicateSlide(slides: SourceSlide[], index: number): SlideEdit {
  const at = locate(slides, index);
  const source = slides[at];
  if (!source) return kept(slides, 0, "That slide is gone.");
  const copy = cloneSlide(source);
  const next = slides.slice();
  next.splice(at + 1, 0, copy);
  return { slides: next, index: at + 1, note: "Duplicated." };
}

/** The title slide stays. Any other slide can leave the deck. */
export function deleteSlide(slides: SourceSlide[], index: number): SlideEdit {
  const at = locate(slides, index);
  if (at <= 0 || slides.length < 2) return kept(slides, at, "The title slide stays.");
  const next = slides.filter((_, position) => position !== at);
  return { slides: next, index: Math.min(at, next.length - 1), note: "Deleted." };
}

/** The first title stays. Two charts stay apart. A full result continues. */
export function mergeSlides(slides: SourceSlide[], index: number): SlideEdit {
  const at = locate(slides, index);
  const left = slides[at];
  const right = slides[at + 1];
  if (!left || !right) return kept(slides, at, "There is no next slide.");
  if (at <= 0) return kept(slides, at, "The title slide stays the title.");
  if (left.chart && right.chart) return kept(slides, at, "Two charts stay on their own slides.");
  const rightTitle = right.title.trim();
  const keepTitle = rightTitle && rightTitle !== left.title.trim() ? rightTitle : "";
  const pieces = [
    ...bodyPieces(left),
    ...(keepTitle ? [{ kind: "prose" as const, text: keepTitle }] : []),
    ...bodyPieces(right),
  ];
  const context = [left.context.trim(), keepTitle, right.context.trim()].filter(Boolean).join("\n");
  const chart = left.chart ?? right.chart;
  const table = (left.chart ? left.table : right.table) ?? left.table ?? right.table;
  const images = [...left.images, ...right.images];
  const merged = cloneSlide({
    ...left,
    context,
    pieces,
    images,
    chart,
    table,
    numbers: [...left.numbers, ...right.numbers],
    extraCharts: left.extraCharts + right.extraCharts,
    picture: images.length > 0 && `${left.title} ${context}`.trim().length < 12 && !chart && left.numbers.length + right.numbers.length === 0,
    show: left.show ?? right.show,
    pace: undefined,
  });
  const shaped = reshape(merged);
  const next = [...slides.slice(0, at), ...shaped, ...slides.slice(at + 2)];
  const note = shaped.length > 1
    ? "Merged. It was too full, so it continues. Nothing was rewritten."
    : "Merged. Nothing was rewritten.";
  return { slides: next, index: at, note };
}

/** Pictures are added as themselves. A fifth picture continues on the next slide. */
export function addSlidePictures(slides: SourceSlide[], index: number, images: SourceImage[]): SlideEdit {
  const at = locate(slides, index);
  const slide = slides[at];
  if (!slide || images.length === 0) return kept(slides, at, "Use a PNG or JPEG.");
  const pieces = slide.pieces ?? [];
  const structured = pieces.some((piece) => piece.kind === "card" || piece.kind === "list");
  const crowded = structured && pieceDensity(pieces, slide.title, isContinued(slide)) !== "light";
  if (crowded && slide.images.length === 0) {
    const follower = reshape({
      ...cloneSlide(slide),
      context: "Continued.",
      pieces: [],
      numbers: [],
      chart: null,
      table: null,
      images,
      extraCharts: 0,
      picture: images.length > 0,
      show: undefined,
      pace: "Continued. The pictures stay.",
    });
    const next = [...slides.slice(0, at + 1), ...follower, ...slides.slice(at + 1)];
    return { slides: next, index: at + 1, note: "The picture continues on the next slide. The words stay." };
  }
  const merged = { ...cloneSlide(slide), images: [...slide.images, ...images], picture: false, pace: undefined };
  const shaped = reshape(merged);
  const next = [...slides.slice(0, at), ...shaped, ...slides.slice(at + 1)];
  const note = shaped.length > 1
    ? "Too many pictures for one slide. They continue on the next. The pictures stay."
    : images.length === 1 ? "Picture added." : `${images.length} pictures added.`;
  return { slides: next, index: at, note };
}

export function removeSlidePicture(slides: SourceSlide[], index: number, imageIndex: number): SlideEdit {
  const at = locate(slides, index);
  const slide = slides[at];
  if (!slide || imageIndex < 0 || imageIndex >= slide.images.length) return kept(slides, at, "That picture is already gone.");
  const images = slide.images.filter((_, position) => position !== imageIndex);
  const words = `${slide.title} ${slide.context}`.trim().length;
  const next = slides.map((item, position) => position === at
    ? { ...item, images, picture: images.length > 0 && words < 12 && !item.chart && item.numbers.length === 0, pace: undefined }
    : item);
  return { slides: next, index: at, note: "Picture removed." };
}

export function showSlide(slides: SourceSlide[], index: number, show: SlideShow): SlideEdit {
  const at = locate(slides, index);
  const slide = slides[at];
  if (!slide || !canFlip(slide)) return kept(slides, at, "A picture of a chart stays a picture.");
  const next = slides.map((item, position) => (position === at ? { ...item, show } : item));
  const note = show === "table" ? "Shown as a table. The cells stay." : show === "chart" ? "Shown as a chart. The cells stay." : "Chart and table.";
  return { slides: next, index: at, note };
}

export type SlideImage = {
  index: number;
  mime: "image/png" | "image/jpeg";
  base64: string;
};

function imageMime(name: string): SlideImage["mime"] | null {
  const lower = name.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  return null;
}

/** Pictures on slides that have no native chart. Tiny files and other formats are skipped. */
export async function slideImages(buffer: ArrayBuffer, indexes: number[]): Promise<SlideImage[]> {
  const wanted = new Set(indexes);
  if (wanted.size === 0) return [];
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return [];
  }
  const names = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((left, right) => Number(left.match(/(\d+)\.xml$/i)?.[1] ?? 0) - Number(right.match(/(\d+)\.xml$/i)?.[1] ?? 0));
  const found: SlideImage[] = [];
  for (const name of names) {
    const index = names.indexOf(name) + 1;
    if (!wanted.has(index) || found.length >= 6) continue;
    const rels = await zip.file(name.replace("ppt/slides/", "ppt/slides/_rels/").replace(/\.xml$/i, ".xml.rels"))?.async("string");
    if (!rels) continue;
    const xml = await zip.file(name)?.async("string");
    if (!xml) continue;
    const embeds = new Set([...xml.matchAll(/\br:embed="([^"]+)"/g)].map((match) => match[1] ?? "").filter(Boolean));
    let best: { bytes: Uint8Array; mime: SlideImage["mime"] } | null = null;
    for (const tag of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = attr(tag[0], "Id");
      const target = attr(tag[0], "Target");
      const type = attr(tag[0], "Type");
      if (!id || !target || !embeds.has(id) || !/\/image$/i.test(type)) continue;
      const mime = imageMime(target);
      if (!mime) continue;
      const path = resolvePath(name, target);
      const file = zip.file(path);
      if (!file) continue;
      const bytes = await file.async("uint8array");
      if (bytes.length < 8000 || bytes.length > 8_000_000) continue;
      if (!best || bytes.length > best.bytes.length) best = { bytes, mime };
    }
    if (!best) continue;
    let binary = "";
    for (let cursor = 0; cursor < best.bytes.length; cursor += 0x8000) {
      binary += String.fromCharCode(...best.bytes.subarray(cursor, cursor + 0x8000));
    }
    found.push({ index, mime: best.mime, base64: btoa(binary) });
  }
  return found;
}

/** A picture read replaces only a slide that has no native chart. */
export function applyPictureReads(deck: SourceDeck, reads: { index: number; text: string }[]): SourceDeck {
  const slides = deck.slides.map((slide) => {
    if (slide.chart) return slide;
    const text = reads.find((item) => item.index === slide.index)?.text ?? "";
    if (!text.trim()) return slide;
    const read = pictureFromModel(text);
    if (!read) return slide;
    if (read.chart) {
      return { ...slide, picture: false, title: slide.title.trim() || read.title, chart: read.chart };
    }
    if (read.numbers.length === 0 && !read.title) return slide;
    return {
      ...slide,
      picture: false,
      title: slide.title.trim() || read.title,
      numbers: read.numbers.length > 0 ? read.numbers : slide.numbers,
    };
  });
  const next: SourceDeck = { ...deck, slides };
  next.intent = intentOf(slides, deck.title);
  return next;
}
