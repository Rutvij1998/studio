import type { SlidePiece, SourceSlide } from "./book";
import { directedComposition, directionFits, type Composition } from "./sourceDeck";

export type DirectorRole = "stack" | "metrics" | "cascade" | "hero" | "versus" | "path" | "proof" | "sheet";

export type DeckFacts = {
  repeated: string[];
  slides: Array<{
    index: number;
    title: string;
    chart: boolean;
    table: boolean;
    images: number;
    lines: string[];
  }>;
};

export type FigureBind = { display: string; label: string };

export type SlideDirection = {
  index: number;
  role: DirectorRole;
  quiet: string[];
  figures: FigureBind[];
};

export type DirectorPlan = {
  drop: string[];
  keep: number[];
  slides: SlideDirection[];
};

export const DIRECTOR_GUIDE = `You direct a restyle before it is shown. Return one JSON object and nothing else.
{"keep":[1],"drop":[""],"slides":[{"index":2,"role":"cascade","quiet":[],"figures":[{"display":"","label":""}]}]}
keep lists the slides that already work and must stay: the cover, every chart, every table, every picture slide, and any slide whose layout is already clear.
drop quotes chrome only. Include every string from repeated, and a one-word or two-word chip such as Internal. Quote it exactly.
slides lists only a slide that needs a change, and never a slide you put in keep. role is stack, cascade, metrics, hero, versus, path, proof, or sheet. cascade is two or three short figures of one unit, sized by those values. metrics is a row of figures that each keep their sentence.
figures copies a number already on that slide and the words that should sit under it. If the line under the number repeats the number, use the sentence instead. quiet is only the step number 1 or 2 in front of a name.
Do not add a number, a percent, a date, or a sentence. Do not rewrite the deck.`;

const ROLES = new Set<DirectorRole>(["stack", "metrics", "cascade", "hero", "versus", "path", "proof", "sheet"]);

function norm(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

function digits(text: string): string[] {
  return text.match(/\d[\d,]*(?:\.\d+)?%?/g) ?? [];
}

function pieceLines(piece: SlidePiece): string[] {
  if (piece.kind === "prose") return piece.text.split("\n");
  if (piece.kind === "list") return [piece.title, ...piece.points];
  return [piece.title, piece.detail, piece.display, piece.label];
}

function corpus(slide: SourceSlide): string {
  const lines = [slide.title, slide.context, ...(slide.pieces ?? []).flatMap(pieceLines)];
  return norm(lines.join(" "));
}

function lineHits(slides: SourceSlide[], line: string): number {
  const key = norm(line);
  return slides.filter((slide) => corpus(slide).toLowerCase().includes(key.toLowerCase()) && key.length > 0).length;
}

function sourceHits(slides: SourceSlide[], line: string): number {
  const key = norm(line).toLowerCase();
  const seen = new Set<number>();
  for (const slide of slides) {
    if (!key || !corpus(slide).toLowerCase().includes(key)) continue;
    seen.add(slide.from ?? slide.index);
  }
  return seen.size;
}

/** A footer or a classification chip. A sentence that ends the slide stays. */
export function chromeLine(slides: SourceSlide[], line: string): boolean {
  const key = norm(line);
  if (key.length < 2 || key.length > 80) return false;
  if (/[.!?]"?$/.test(key)) return false;
  const hits = lineHits(slides, key);
  const chip = key.length <= 24 && !/\d/.test(key) && hits >= 2;
  const sources = new Set(slides.map((slide) => slide.from ?? slide.index)).size;
  const repeated = sourceHits(slides, key) >= Math.max(3, Math.ceil(sources * 0.5))
    || hits >= Math.max(4, Math.ceil(slides.length * 0.4));
  return chip || repeated;
}

function quoteFrom(haystack: string, phrase: string): string | null {
  const hay = norm(haystack);
  const needle = norm(phrase);
  if (needle.length < 2) return null;
  const at = hay.toLowerCase().indexOf(needle.toLowerCase());
  if (at < 0) return null;
  return hay.slice(at, at + needle.length);
}

export function deckFacts(slides: SourceSlide[]): DeckFacts {
  const counts = new Map<string, number>();
  for (const slide of slides) {
    const seen = new Set<string>();
    for (const line of [slide.title, slide.context, ...(slide.pieces ?? []).flatMap(pieceLines)]) {
      const key = norm(line);
      if (key.length < 12 || key.length > 80 || seen.has(key)) continue;
      seen.add(key);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
  }
  const repeated = [...counts.entries()]
    .filter(([key, count]) => count >= 3 && chromeLine(slides, key))
    .sort((left, right) => right[1] - left[1])
    .slice(0, 8)
    .map(([key]) => key);
  return {
    repeated,
    slides: slides.slice(0, 24).map((slide) => ({
      index: slide.index,
      title: norm(slide.title).slice(0, 90),
      chart: Boolean(slide.chart),
      table: Boolean(slide.table),
      images: slide.images.length,
      lines: slide.chart || slide.table ? [] : [slide.context, ...(slide.pieces ?? []).flatMap(pieceLines)]
        .map((line) => norm(line))
        .filter((line) => line.length > 0)
        .slice(0, 8)
        .map((line) => line.slice(0, 140)),
    })),
  };
}

function asRole(value: unknown): DirectorRole | null {
  return typeof value === "string" && ROLES.has(value as DirectorRole) ? value as DirectorRole : null;
}

/** Keep only the decisions that quote this deck. A bad field is dropped, not the whole plan. */
export function acceptPlan(slides: SourceSlide[], raw: unknown): DirectorPlan {
  const body = raw && typeof raw === "object" ? raw as { drop?: unknown; keep?: unknown; slides?: unknown } : {};
  const known = new Set(slides.map((slide) => slide.index));
  const keep = new Set<number>();
  if (Array.isArray(body.keep)) {
    for (const item of body.keep) {
      if (typeof item === "number" && known.has(item)) keep.add(item);
    }
  }
  const drop: string[] = [];
  if (Array.isArray(body.drop)) {
    for (const item of body.drop) {
      if (typeof item !== "string" || drop.length >= 8) continue;
      const key = norm(item);
      if (!chromeLine(slides, key) || drop.some((line) => norm(line).toLowerCase() === key.toLowerCase())) continue;
      drop.push(key);
    }
  }
  const planned: SlideDirection[] = [];
  const rows = Array.isArray(body.slides) ? body.slides : [];
  for (const item of rows) {
    if (!item || typeof item !== "object" || planned.length >= 24) continue;
    const row = item as { index?: unknown; role?: unknown; quiet?: unknown; figures?: unknown };
    if (typeof row.index !== "number") continue;
    const slide = slides.find((entry) => entry.index === row.index);
    if (!slide || keep.has(slide.index) || slide.chart || slide.table || slide.images.length > 0) continue;
    const role = asRole(row.role) ?? "stack";
    const text = corpus(slide);
    const quiet: string[] = [];
    if (Array.isArray(row.quiet)) {
      for (const value of row.quiet) {
        const key = norm(typeof value === "string" ? value : "");
        if (!/^[12]$/.test(key)) continue;
        const card = (slide.pieces ?? []).find((piece) => piece.kind === "card" && norm(piece.display) === key && /[A-Za-z]/.test(piece.title));
        if (card) quiet.push(key);
      }
    }
    const figures: FigureBind[] = [];
    if (Array.isArray(row.figures)) {
      for (const value of row.figures) {
        if (!value || typeof value !== "object" || figures.length >= 6) continue;
        const figure = value as { display?: unknown; label?: unknown };
        if (typeof figure.display !== "string" || typeof figure.label !== "string") continue;
        const display = norm(figure.display);
        const label = quoteFrom(text, figure.label);
        if (!display || !label || norm(label).toLowerCase() === display.toLowerCase()) continue;
        if (digits(label).some((run) => !digits(text).includes(run))) continue;
        const card = (slide.pieces ?? []).find((piece) => piece.kind === "card" && (norm(piece.display) === display || norm(piece.title) === display));
        if (!card || card.kind !== "card") continue;
        figures.push({ display, label });
      }
    }
    const drawable = role !== "sheet" && directionFits(role, slide.pieces ?? [], slide.title);
    if (!drawable && quiet.length === 0 && figures.length === 0) continue;
    planned.push({ index: slide.index, role: drawable ? role : "sheet", quiet, figures });
  }
  return { drop, keep: [...keep], slides: planned };
}

function withoutLine(text: string, lines: string[]): string {
  return text
    .split("\n")
    .map((item) => item.trim())
    .filter((item) => item && !lines.some((line) => norm(item).toLowerCase() === norm(line).toLowerCase()))
    .join("\n");
}

function stripPiece(piece: SlidePiece, lines: string[]): SlidePiece | null {
  const hit = (text: string) => lines.some((line) => norm(text).toLowerCase() === norm(line).toLowerCase());
  if (piece.kind === "prose") {
    if (hit(piece.text)) return null;
    const text = withoutLine(piece.text, lines);
    return text.trim() ? { ...piece, text } : null;
  }
  if (piece.kind === "list") {
    const points = piece.points.filter((point) => !hit(point));
    const title = hit(piece.title) ? "" : piece.title;
    if (!title && points.length === 0) return null;
    return { ...piece, title, points };
  }
  let title = piece.title;
  let detail = withoutLine(piece.detail, lines);
  let label = hit(piece.label) ? "" : piece.label;
  if (hit(title) && (detail.trim() || piece.display)) title = "";
  else if (hit(title) && !detail.trim() && !piece.display) return null;
  if (!title.trim() && !detail.trim() && !piece.display && !label.trim()) return null;
  return { ...piece, title, detail, label };
}

function visible(slide: SourceSlide): boolean {
  if (slide.chart || slide.table || slide.images.length > 0) return true;
  for (const piece of slide.pieces ?? []) {
    if (piece.kind === "prose" && piece.text.trim()) return true;
    if (piece.kind === "list" && (piece.title.trim() || piece.points.some((point) => point.trim()))) return true;
    if (piece.kind === "card" && (piece.title.trim() || piece.detail.trim() || piece.display.trim() || piece.label.trim())) return true;
  }
  return slide.context.split("\n").some((line) => {
    const text = line.trim();
    return text.length > 0 && text !== "Continued.";
  });
}

function rebind(slide: SourceSlide, plan: SlideDirection): SourceSlide {
  const pieces = (slide.pieces ?? []).map((piece) => {
    if (piece.kind === "list") return { ...piece, points: [...piece.points] };
    if (piece.kind === "prose") return { ...piece };
    return { ...piece };
  });
  for (const key of plan.quiet) {
    const card = pieces.find((piece) => piece.kind === "card" && norm(piece.display) === key);
    if (!card || card.kind !== "card") continue;
    const title = norm(card.title).startsWith(key) ? card.title : `${key} ${card.title}`.trim();
    card.display = "";
    card.label = "";
    card.title = title;
  }
  for (const figure of plan.figures) {
    const card = pieces.find((piece) => piece.kind === "card" && (norm(piece.display) === figure.display || norm(piece.title) === figure.display));
    if (!card || card.kind !== "card") continue;
    const display = card.display || figure.display;
    card.display = display;
    card.label = figure.label;
    if (norm(card.title).toLowerCase() === norm(display).toLowerCase()) card.title = figure.label;
    if (norm(card.detail).toLowerCase() === norm(figure.label).toLowerCase()) card.detail = "";
  }
  const role = plan.role === "sheet" ? undefined : plan.role;
  const direction = role && directionFits(role, pieces, slide.title) ? role as Composition : undefined;
  const next = { ...slide, pieces, direction };
  if (direction && directedComposition(next) === direction) return next;
  return { ...next, direction: undefined };
}

/** Apply an accepted plan. Figures stay the strings already on the slide. */
export function applyPlan(slides: SourceSlide[], plan: DirectorPlan): { slides: SourceSlide[]; changed: boolean } {
  const drop = plan.drop;
  let changed = drop.length > 0;
  const next = slides.map((slide) => {
    const pieces = (slide.pieces ?? [])
      .map((piece) => stripPiece(piece, drop))
      .filter((piece): piece is SlidePiece => piece != null);
    const context = withoutLine(slide.context, drop);
    const planned = plan.slides.find((item) => item.index === slide.index);
    const base = { ...slide, context, pieces: pieces.length ? pieces : undefined };
    if (!planned) return base;
    changed = true;
    return rebind(base, planned);
  });
  const kept = next.filter((slide, index) => index === 0 || visible(slide));
  if (kept.length !== next.length) changed = true;
  const indexed = kept.map((slide, index) => ({ ...slide, index: index + 1 }));
  const same = !changed || JSON.stringify(indexed.map((slide) => [slide.title, slide.context, slide.pieces, slide.direction]))
    === JSON.stringify(slides.map((slide) => [slide.title, slide.context, slide.pieces, slide.direction]));
  return { slides: indexed, changed: !same };
}

/** Trim a request before it reaches the model. */
export function packFacts(raw: unknown): DeckFacts {
  const body = raw && typeof raw === "object" ? raw as { repeated?: unknown; slides?: unknown } : {};
  const repeated = Array.isArray(body.repeated)
    ? body.repeated.filter((item): item is string => typeof item === "string").map((item) => norm(item).slice(0, 80)).filter(Boolean).slice(0, 8)
    : [];
  const slides = Array.isArray(body.slides) ? body.slides.flatMap((item) => {
    if (!item || typeof item !== "object") return [];
    const row = item as { index?: unknown; title?: unknown; chart?: unknown; table?: unknown; images?: unknown; lines?: unknown };
    if (typeof row.index !== "number") return [];
    const lines = Array.isArray(row.lines)
      ? row.lines.filter((line): line is string => typeof line === "string").map((line) => norm(line).slice(0, 180)).filter(Boolean).slice(0, 14)
      : [];
    return [{
      index: row.index,
      title: typeof row.title === "string" ? norm(row.title).slice(0, 90) : "",
      chart: Boolean(row.chart),
      table: Boolean(row.table),
      images: typeof row.images === "number" ? Math.max(0, Math.min(8, row.images)) : 0,
      lines,
    }];
  }).slice(0, 24) : [];
  return { repeated, slides };
}
