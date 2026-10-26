import { directedComposition, pieceAir, type Composition } from "./sourceDeck";
import type { SourceSlide, SlidePiece } from "./book";

export type GradeLetter = "A" | "B" | "C" | "D";

export type GradeMark = {
  index: number;
  title: string;
  severity: "empty" | "sparse";
  note: string;
};

export type DeckGrade = {
  grade: GradeLetter;
  summary: string;
  slides: GradeMark[];
};

export type GradeBrief = {
  index: number;
  title: string;
  composition: Composition;
  text: string;
  problem: string;
};

export const GRADE_GUIDE = `Grade the restyle. Return one JSON object and nothing else.
{"summary":"","slides":[{"index":1,"note":""}]}
Write one sentence for the summary, and one sentence for each slide that has a problem. Skip a slide whose problem is empty.
Use only the words already in that slide. Do not add a number, a percent, a date, or a fact. Do not rewrite the slide.`;

function wordsOf(pieces: SlidePiece[]): string[] {
  const text = pieces.map((piece) => {
    if (piece.kind === "prose") return piece.text;
    if (piece.kind === "card") return `${piece.title} ${piece.detail} ${piece.display} ${piece.label}`;
    return `${piece.title} ${piece.points.join(" ")}`;
  }).join(" ");
  return text.replace(/\s+/g, " ").trim().split(" ").filter(Boolean);
}

function excerpt(slide: SourceSlide): string {
  const pieces = slide.pieces ?? [];
  const body = wordsOf(pieces).join(" ");
  const text = [slide.title, slide.context, body].filter(Boolean).join(" ").replace(/\s+/g, " ").trim();
  return text.slice(0, 280);
}

/** A sentence beside a title-only card is the empty plate. Real lists keep the slide full. */
function blankPlate(pieces: SlidePiece[]): boolean {
  if (pieces.length !== 2) return false;
  const cards = pieces.filter((piece) => piece.kind === "card");
  const prose = pieces.filter((piece) => piece.kind === "prose");
  if (cards.length !== 1 || prose.length !== 1) return false;
  const card = cards[0];
  const line = prose[0];
  if (!card || card.kind !== "card" || !line || line.kind !== "prose") return false;
  if (card.display || card.detail.trim() || line.text.includes("\n") || line.text.length > 200) return false;
  return true;
}

function letterFor(empty: number, sparse: number): GradeLetter {
  if (empty >= 2 || empty + sparse >= 4) return "D";
  if (empty >= 1 || sparse >= 2) return "C";
  if (sparse === 1) return "B";
  return "A";
}

/** Read every restyled slide. The letter comes from the structure, not from a model. */
export function gradeDeck(slides: SourceSlide[]): DeckGrade {
  const marks: GradeMark[] = [];
  for (const slide of slides) {
    if (slide.chart || slide.table || slide.images.length > 0) continue;
    if (slide.index === 1 && (slide.pieces ?? []).length === 0) continue;
    const pieces = slide.pieces ?? [];
    const points = pieces.reduce((total, piece) => total + (piece.kind === "list" ? piece.points.length : 0), 0);
    const filled = pieces.some((piece) => piece.kind === "card" && (piece.display || piece.detail.trim()));
    const title = slide.title.replace(/\s+/g, " ").trim() || "Untitled";
    if (blankPlate(pieces)) {
      marks.push({
        index: slide.index,
        title,
        severity: "empty",
        note: "A sentence and a label were drawn as an empty card.",
      });
      continue;
    }
    const thin = wordsOf(pieces).length <= 18 && points < 3 && !filled && pieceAir(slide);
    if (thin) {
      marks.push({
        index: slide.index,
        title,
        severity: "sparse",
        note: "Most of this slide is empty. The line does not fill it.",
      });
    }
  }
  const empty = marks.filter((mark) => mark.severity === "empty").length;
  const sparse = marks.length - empty;
  const grade = letterFor(empty, sparse);
  const count = marks.length;
  const summary = count === 0
    ? "Every slide holds together."
    : `${count} ${count === 1 ? "slide needs" : "slides need"} another look.`;
  return { grade, summary, slides: marks };
}

export function gradeBrief(slides: SourceSlide[]): { grade: DeckGrade; brief: GradeBrief[] } {
  const grade = gradeDeck(slides);
  const brief = slides.map((slide) => ({
    index: slide.index,
    title: slide.title.replace(/\s+/g, " ").trim().slice(0, 90),
    composition: directedComposition(slide),
    text: excerpt(slide),
    problem: grade.slides.find((mark) => mark.index === slide.index)?.note ?? "",
  }));
  return { grade, brief };
}

function digits(text: string): string[] {
  return text.match(/\d[\d,]*(?:\.\d+)?%?/g) ?? [];
}

/** A model note may only repeat numbers that are already on the slide. */
export function acceptNote(source: string, note: string): string | null {
  const clean = note.replace(/\s+/g, " ").trim();
  if (clean.length < 8 || clean.length > 180) return null;
  const allowed = new Set(digits(source));
  for (const value of digits(clean)) {
    if (!allowed.has(value)) return null;
  }
  return clean;
}
