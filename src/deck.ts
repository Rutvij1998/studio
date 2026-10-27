import PptxGenJS from "pptxgenjs";
import { coverCopy } from "./brief";
import { conversationOf, formatCount, hasBroughtIn, viewBook, type Book, type Proof, type SlidePiece, type SourceChart, type SourceImage, type SourceSlide, type SourceTable } from "./book";
import { deckFileName, type DeckOptions, type TypeSize } from "./options";
import { slidesFor } from "./slides";
import { compactVolume, isResults, proofPhrase, storyLine, storyReadLines } from "./story";
import { factLine, factReadLines, isFacts } from "./facts";
import { chartFit, comboPair } from "./fit";
import { columnPoints, formatGridPoint, gridChartNote, gridLine, gridProgress, isGrid } from "./grid";
import { isOutline, outlineLine } from "./outline";
import { isSeries, rateOutlook, readingLine, seriesProgressList, trackNote, trackTotals } from "./reading";
import { coverWave } from "./dotWave";
import { axisMax, axisSpan, bulletLines, cardFace, chartEyebrow, chartFoot, compareSplit, directedComposition, copyBlocks, displayPx, figureBody, figurePx, figureShare, fillBandIndex, formatPoint, gridAcross, heroAside, imageSize, isContinued, isSource, isTimeTitle, leadCopy, pictureLayout, pieceAir, pieceBands, pieceFit, pieceFrame, plotMode, posterPx, proofPx, proofSplit, rowDisplay, sameFigure, slideRole, slideShow, spareHeaders, tableFromSlide, tableHasNumbers, type PieceBand } from "./sourceDeck";

const INK = "14003B";
const BRAND = "3200BE";
const PINK = "FF96FF";
const STAGE = "0B0027";
const MUTED = "635D7C";
const FAINT = "837D9B";
const LINE = "E4E1EC";
const WHITE = "FFFFFF";
const NEG = "E03B5B";
const POS = "1F9E72";
const NEU = "CAC6DA";
const APPROVED = "7C8CFF";
const OPEN = "B7AFC9";
const REST = "D4CEDF";

const TONE: Record<string, string> = {
  fulfilled: PINK,
  approved: APPROVED,
  open: OPEN,
  rest: REST,
};

const SCALE: Record<TypeSize, { title: number; rate: number; body: number; kicker: number }> = {
  s: { title: 32, rate: 54, body: 15, kicker: 11 },
  m: { title: 40, rate: 66, body: 18, kicker: 12 },
  l: { title: 48, rate: 78, body: 20, kicker: 13 },
};

type Slide = PptxGenJS.Slide;
type Pres = PptxGenJS & {
  shapes: typeof PptxGenJS.shapes;
  charts: typeof PptxGenJS.charts;
};

function fitTitle(text: string, size: number): number {
  if (text.length > 22) return Math.max(28, size - 12);
  if (text.length > 16) return Math.max(30, size - 6);
  return size;
}

export async function deckBlob(book: Book, options: DeckOptions): Promise<{ blob: Blob; fileName: string }> {
  const pres = new PptxGenJS() as Pres;
  pres.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  pres.layout = "DECK";
  const results = isResults(book);
  const series = isSeries(book);
  const source = isSource(book);
  const facts = isFacts(book);
  const grid = isGrid(book);
  const outline = isOutline(book);
  const kind = source ? "slides" : series ? "briefing" : grid ? "sheet" : facts ? "figures" : results ? "results" : outline ? "outline" : "scorecard";
  pres.title = book.client ? `${book.client} ${kind}` : source ? "From the deck" : series ? "Briefing" : grid ? "Sheet" : facts ? "Figures" : results ? "Results" : outline ? "Outline" : "Scorecard";
  pres.author = "Studio";
  pres.subject = source ? "From the deck" : series ? "Briefing" : grid ? "Sheet" : facts ? "Figures" : results ? "Results" : outline ? "Outline" : "Client scorecard";

  const view = viewBook(book);
  const type = SCALE[options.type];
  for (const slide of slidesFor(book)) {
    if (slide.id === "origin") addOrigin(pres, book, slide.track ?? 1, type);
    else if (slide.id === "cover") addCover(pres, book, type);
    else if (slide.id === "scope") addScope(pres, book, type);
    else if (slide.id === "track") addTrack(pres, book, slide.track ?? 0, type);
    else if (slide.id === "compare") addCompare(pres, book, type);
    else if (slide.id === "result") addResult(pres, book, type);
    else if (slide.id === "facts") addFacts(pres, book, slide.track ?? 0, type);
    else if (slide.id === "progress") addProgress(pres, book, type);
    else if (slide.id === "outlook") addOutlook(pres, book, type);
    else if (slide.id === "grid") addGridColumn(pres, book, slide.track ?? 0, type);
    else if (slide.id === "combo") addCombo(pres, book, type);
    else if (slide.id === "card") addCard(pres, book, slide.track ?? 0, type);
    else if (slide.id === "read") addRead(pres, book, type);
    else if (slide.id === "book") addBook(pres, book, view, type);
    else if (slide.id === "mix") addMix(pres, book, view, options, type);
    else if (slide.id === "versus") addVersus(pres, book, view, type);
    else if (slide.id === "listening") addListening(pres, book, options, type);
    else if (slide.id === "journey") addJourney(pres, book, type);
    else if (slide.id === "months") addMonths(pres, book, type);
    else if (slide.id === "source") addSource(pres, book, type);
    else addNext(pres, book, view, type);
  }

  const blob = (await pres.write({ outputType: "blob" })) as Blob;
  const fileName = deckFileName(book.client);
  const named = source
    ? fileName.replace(/scorecard\.pptx$/, "slides.pptx")
    : series
      ? fileName.replace(/scorecard\.pptx$/, "briefing.pptx")
      : grid
        ? fileName.replace(/scorecard\.pptx$/, "sheet.pptx")
        : facts
          ? fileName.replace(/scorecard\.pptx$/, "figures.pptx")
          : results
            ? fileName.replace(/scorecard\.pptx$/, "results.pptx")
            : isOutline(book)
              ? fileName.replace(/scorecard\.pptx$/, "outline.pptx")
              : fileName;
  return { blob, fileName: named };
}

const COVER = "2D0BB0";
const CARD_LINE = "E0C3FF";
const CARD_RADIUS = 0.2;

/** The wave lives on the title slide only, behind the words. */
function paintCoverWave(pres: Pres, slide: Slide) {
  for (const dot of coverWave()) {
    const d = dot.r * 13.333;
    slide.addShape(pres.shapes.OVAL, {
      x: dot.x * 13.333 - d,
      y: dot.y * 7.5 - d,
      w: d * 2,
      h: d * 2,
      fill: { color: dot.fill, transparency: Math.round((1 - dot.opacity) * 100) },
      line: { color: dot.fill, width: 0 },
    });
  }
}

function kicker(slide: Slide, text: string, x: number, y: number, color: string, size: number) {
  slide.addText(text.toUpperCase(), {
    x,
    y,
    w: 6,
    h: 0.32,
    fontFace: "Calibri",
    fontSize: size,
    bold: true,
    color,
    margin: 0,
    charSpacing: 1.5,
  });
}

function imageData(value: string): string | null {
  const match = value.match(/^data:(image\/[a-z0-9.+-]+);base64,([A-Za-z0-9+/=]+)$/i);
  if (!match?.[1] || !match[2]) return null;
  return `${match[1]};base64,${match[2]}`;
}

function placeKept(slide: Slide, images: SourceImage[], x: number, y: number, w: number, h: number) {
  const ready = images
    .slice(0, 4)
    .map((image) => ({ data: imageData(image.dataUrl), size: imageSize(image.dataUrl) }))
    .filter((image): image is { data: string; size: { width: number; height: number } | null } => Boolean(image.data));
  if (ready.length === 0) return;
  const layout = pictureLayout(images.slice(0, ready.length));
  const count = ready.length;
  const cols = layout === "pair" || layout === "grid" ? Math.min(2, count) : layout === "row" ? count : 1;
  const rows = layout === "stack" ? count : Math.ceil(count / cols);
  const gap = 0.16;
  const cellW = (w - gap * (cols - 1)) / cols;
  const cellH = (h - gap * (rows - 1)) / rows;
  ready.forEach((image, index) => {
    const column = index % cols;
    const row = Math.floor(index / cols);
    const left = x + column * (cellW + gap);
    const top = y + row * (cellH + gap);
    const ratio = image.size && image.size.height > 0 ? image.size.width / image.size.height : cellW / cellH;
    const fittedW = Math.min(cellW, cellH * ratio);
    const fittedH = Math.min(cellH, fittedW / ratio);
    slide.addImage({
      data: image.data,
      x: left + (cellW - fittedW) / 2,
      y: top + Math.max(0, (cellH - fittedH) / 2),
      w: fittedW,
      h: fittedH,
      altText: "Picture from the deck",
    });
  });
}

function placeMark(pres: Pres, slide: Slide, logo: string) {
  const data = imageData(logo);
  if (!data) return;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 10.55,
    y: 0.38,
    w: 2.15,
    h: 0.78,
    fill: { color: WHITE },
    rectRadius: 0.1,
  });
  slide.addImage({
    data,
    x: 10.7,
    y: 0.5,
    w: 1.85,
    h: 0.52,
    sizing: { type: "contain", w: 1.85, h: 0.52 },
    altText: "Client mark",
  });
}

function addCover(
  pres: Pres,
  book: Book,
  type: (typeof SCALE)["m"],
) {
  const slide = pres.addSlide();
  slide.background = { color: COVER };
  paintCoverWave(pres, slide);
  const series = isSeries(book);
  const facts = isFacts(book);
  const grid = isGrid(book);
  const outline = isOutline(book);
  if (book.logo) placeMark(pres, slide, book.logo);
  kicker(slide, series ? "Briefing" : grid ? "Sheet" : facts ? "Figures" : isResults(book) ? "Results" : outline ? "Outline" : "Client scorecard", 0.7, 0.55, PINK, type.kicker);
  const title = book.client || (outline ? book.outline?.subject : "") || "Untitled book";
  const meta = [book.market, book.period].filter(Boolean).join("  ·  ") || (facts || series || grid || isResults(book) || outline ? "" : "Market and period");
  slide.addText(title, {
    x: 0.7,
    y: 1.15,
    w: book.logo ? 9.4 : 11.5,
    h: meta ? 1.55 : 0.95,
    fontFace: "Calibri",
    fontSize: fitTitle(title, type.title + 14),
    bold: true,
    color: WHITE,
    margin: 0,
    valign: "top",
  });
  if (meta) slide.addText(meta, {
    x: 0.7,
    y: 2.85,
    w: 10,
    h: 0.36,
    fontFace: "Calibri",
    fontSize: type.body,
    color: "C8C4D4",
    margin: 0,
    valign: "top",
  });
  const coverLine = outline && book.outline
    ? outlineLine(book.outline).slice(0, 280)
    : series && book.reading
    ? readingLine(book.reading).slice(0, 280)
    : grid && book.grid
      ? gridLine(book.grid).slice(0, 280)
      : facts && book.facts
        ? factLine(book.facts).slice(0, 280)
        : book.story && isResults(book)
          ? storyLine(book.story, book.period || book.story.period)
          : coverCopy(book.brief, 280);
  slide.addText(coverLine || "Add a brief, a spreadsheet, or an existing deck.", {
    x: 0.7,
    y: meta ? 3.32 : 2.22,
    w: 10.2,
    h: 1.55,
    fontFace: "Calibri",
    fontSize: type.body + 2,
    color: "E4E0EE",
    margin: 0,
    valign: "top",
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.7,
    y: 6.55,
    w: 11.9,
    h: 0.01,
    fill: { color: "FFFFFF", transparency: 80 },
  });
  const chapters = slidesFor(book)
    .filter((item) => item.id !== "cover")
    .map((item) => item.label)
    .join("     ");
  slide.addText(chapters, {
    x: 0.7,
    y: 6.7,
    w: 11.9,
    h: 0.36,
    fontFace: "Calibri",
    fontSize: 13,
    color: "A39CB8",
    margin: 0,
  });
}

function addScope(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  const cards = book.reading?.scope.slice(0, 4) ?? [];
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "What counts", 0.65, 0.38, "C23FC2", type.kicker);
  slide.addText(book.client || book.reading?.title || "The count", {
    x: 0.65,
    y: 0.72,
    w: 12,
    h: 0.6,
    fontFace: "Calibri",
    fontSize: type.title - 4,
    bold: true,
    color: INK,
    margin: 0,
  });
  const columns = cards.length >= 4 ? 2 : 1;
  const rows = Math.ceil(cards.length / columns) || 1;
  const gap = 0.2;
  const width = columns === 1 ? 12 : (12 - gap) / 2;
  const height = Math.min(2.4, (5.4 - gap * (rows - 1)) / rows);
  cards.forEach((card, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const x = 0.65 + column * (width + gap);
    const y = 1.6 + row * (height + gap);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x,
      y,
      w: width,
      h: height,
      fill: { color: "F7F5FB" },
      rectRadius: 0.1,
    });
    slide.addText(card.label, {
      x: x + 0.28,
      y: y + 0.18,
      w: width - 0.5,
      h: 0.3,
      fontFace: "Calibri",
      fontSize: 13,
      bold: true,
      color: "C23FC2",
      margin: 0,
    });
    slide.addText(card.text, {
      x: x + 0.28,
      y: y + 0.52,
      w: width - 0.56,
      h: height - 0.7,
      fontFace: "Calibri",
      fontSize: card.text.length > 180 ? 13 : 15,
      color: INK,
      margin: 0,
    });
  });
}

function addTrack(pres: Pres, book: Book, index: number, type: (typeof SCALE)["m"]) {
  const track = book.reading?.tracks[index];
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  if (!track) return;
  const totals = trackTotals(track);
  kicker(slide, track.name, 0.55, 0.32, "C23FC2", type.kicker);
  slide.addText("Containment", {
    x: 0.55,
    y: 0.64,
    w: 6,
    h: 0.5,
    fontFace: "Calibri",
    fontSize: type.title - 6,
    bold: true,
    color: INK,
    margin: 0,
  });
  slide.addText(`${totals.rate}%`, {
    x: 7.4,
    y: 0.4,
    w: 5.3,
    h: 0.7,
    fontFace: "Calibri",
    fontSize: type.rate - 18,
    bold: true,
    color: INK,
    align: "right",
    margin: 0,
  });
  slide.addText(`${formatCount(totals.part)} ${track.partLabel.toLowerCase()} of ${formatCount(totals.total)} ${track.unit}`, {
    x: 6.2,
    y: 1.12,
    w: 6.5,
    h: 0.3,
    fontFace: "Calibri",
    fontSize: 13,
    color: MUTED,
    align: "right",
    margin: 0,
  });
  const points = track.points;
  slide.addChart(pres.charts.LINE, [
    {
      name: track.name,
      labels: points.map((point) => point.label),
      values: points.map((point) => point.rate),
    },
  ], {
    x: 0.4,
    y: 1.5,
    w: 12.5,
    h: 5.05,
    showLegend: false,
    showValue: false,
    chartColors: [BRAND],
    chartArea: { fill: { color: WHITE } },
    plotArea: { fill: { color: WHITE } },
    catAxisLabelColor: MUTED,
    valAxisLabelColor: MUTED,
    catAxisLabelFontSize: points.length > 12 ? 8 : 11,
    catAxisLabelFontFace: "Calibri",
    valAxisLabelFontFace: "Calibri",
    valAxisMinVal: 0,
    valAxisMaxVal: 100,
    lineSize: 2.5,
  });
  slide.addText(trackNote(track), {
    x: 0.55,
    y: 6.95,
    w: 12.2,
    h: 0.4,
    fontFace: "Calibri",
    fontSize: 11,
    color: MUTED,
    margin: 0,
  });
}

function addCompare(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  const tracks = book.reading?.tracks.slice(0, 3) ?? [];
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "Side by side", 0.65, 0.4, "C23FC2", type.kicker);
  slide.addText(book.period || book.client || "The series", {
    x: 0.65,
    y: 0.78,
    w: 12,
    h: 0.55,
    fontFace: "Calibri",
    fontSize: type.title - 6,
    bold: true,
    color: INK,
    margin: 0,
  });
  const gap = 0.24;
  const width = (12 - gap * (tracks.length - 1)) / Math.max(tracks.length, 1);
  tracks.forEach((track, index) => {
    const totals = trackTotals(track);
    const x = 0.65 + index * (width + gap);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x,
      y: 1.7,
      w: width,
      h: 4.3,
      fill: { color: "F7F5FB" },
      rectRadius: 0.1,
    });
    slide.addText(track.name, {
      x: x + 0.3,
      y: 1.95,
      w: width - 0.6,
      h: 0.35,
      fontFace: "Calibri",
      fontSize: 16,
      color: MUTED,
      margin: 0,
    });
    slide.addText(`${totals.rate}%`, {
      x: x + 0.3,
      y: 2.4,
      w: width - 0.6,
      h: 0.9,
      fontFace: "Calibri",
      fontSize: type.rate - 10,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText(`${formatCount(totals.part)} of ${formatCount(totals.total)} ${track.unit}`, {
      x: x + 0.3,
      y: 3.4,
      w: width - 0.6,
      h: 0.4,
      fontFace: "Calibri",
      fontSize: 14,
      color: INK,
      margin: 0,
    });
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.3,
      y: 4.15,
      w: width - 0.6,
      h: 0.14,
      fill: { color: "E4E1EC" },
      rectRadius: 0.07,
    });
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.3,
      y: 4.15,
      w: Math.max(0.08, (width - 0.6) * (Math.min(100, totals.rate) / 100)),
      h: 0.14,
      fill: { color: BRAND },
      rectRadius: 0.07,
    });
  });
  slide.addText("Each rate is that series alone. Calls and chats stay in their own totals.", {
    x: 0.65,
    y: 6.3,
    w: 12,
    h: 0.35,
    fontFace: "Calibri",
    fontSize: 14,
    color: MUTED,
    margin: 0,
  });
}

function addProgress(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  const view = book.grid ? gridProgress(book.grid) : null;
  if (view) {
    kicker(slide, view.title, 0.65, 0.4, "C23FC2", type.kicker);
    slide.addText(view.column, {
      x: 0.65,
      y: 0.78,
      w: 12,
      h: 0.55,
      fontFace: "Calibri",
      fontSize: type.title - 8,
      bold: true,
      color: INK,
      margin: 0,
    });
    const share = view.kind === "share";
    const cards = [
      { label: share ? "Progress" : view.leftLabel, value: view.left, note: share ? "Latest row" : "Previous row" },
      { label: share ? "Not progress" : view.rightLabel, value: view.right, note: share ? "The rest of 100" : "Latest row" },
    ];
    cards.forEach((card, index) => {
      const x = 0.65 + index * 6.2;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x,
        y: 1.8,
        w: 5.8,
        h: 3.6,
        fill: { color: "F3F0F8" },
        rectRadius: 0.12,
      });
      slide.addText(card.label, {
        x: x + 0.35,
        y: 2.05,
        w: 5.1,
        h: 0.35,
        fontFace: "Calibri",
        fontSize: 14,
        bold: true,
        color: MUTED,
        margin: 0,
      });
      slide.addText(share ? `${formatGridPoint(card.value)}%` : formatGridPoint(card.value), {
        x: x + 0.35,
        y: 2.6,
        w: 5.1,
        h: 1.1,
        fontFace: "Calibri",
        fontSize: 40,
        bold: true,
        color: INK,
        margin: 0,
      });
      slide.addText(card.note, {
        x: x + 0.35,
        y: 4.0,
        w: 5.1,
        h: 0.4,
        fontFace: "Calibri",
        fontSize: 16,
        color: INK,
        margin: 0,
      });
    });
    slide.addText(view.note, {
      x: 0.65,
      y: 5.7,
      w: 12,
      h: 0.7,
      fontFace: "Calibri",
      fontSize: 14,
      color: MUTED,
      margin: 0,
    });
    return;
  }
  const cards = book.reading ? seriesProgressList(book.reading) : [];
  kicker(slide, cards.length === 1 ? cards[0]?.direction ?? "Progress" : "Progress", 0.65, 0.4, "C23FC2", type.kicker);
  slide.addText("Latest full month", {
    x: 0.65,
    y: 0.78,
    w: 12,
    h: 0.5,
    fontFace: "Calibri",
    fontSize: type.title - 8,
    bold: true,
    color: INK,
    margin: 0,
  });
  const wide = cards.length < 2;
  cards.slice(0, 2).forEach((card, index) => {
    const x = 0.65 + index * 6.2;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x,
      y: 1.6,
      w: wide ? 12 : 5.8,
      h: 3.8,
      fill: { color: "F3F0F8" },
      rectRadius: 0.12,
    });
    slide.addText(`${card.name}  ·  ${card.direction}  ·  ${card.label}`, {
      x: x + 0.35,
      y: 1.85,
      w: (wide ? 12 : 5.8) - 0.7,
      h: 0.4,
      fontFace: "Calibri",
      fontSize: 14,
      bold: true,
      color: MUTED,
      margin: 0,
    });
    slide.addText(formatCount(card.contained), {
      x: x + 0.35,
      y: 2.45,
      w: wide ? 5.4 : 5.1,
      h: 0.8,
      fontFace: "Calibri",
      fontSize: 32,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText("Contained", {
      x: x + 0.35,
      y: 3.25,
      w: 5,
      h: 0.3,
      fontFace: "Calibri",
      fontSize: 14,
      color: INK,
      margin: 0,
    });
    slide.addText(formatCount(card.rest), {
      x: x + 0.35,
      y: 3.7,
      w: 5,
      h: 0.7,
      fontFace: "Calibri",
      fontSize: 32,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText("Not contained", {
      x: x + 0.35,
      y: 4.45,
      w: 5,
      h: 0.3,
      fontFace: "Calibri",
      fontSize: 14,
      color: INK,
      margin: 0,
    });
  });
  slide.addText("Latest full month only. This is not a sum of the months. A short month and a tiny month are left out.", {
    x: 0.65,
    y: 5.7,
    w: 12,
    h: 0.7,
    fontFace: "Calibri",
    fontSize: 14,
    color: MUTED,
    margin: 0,
  });
}

function addOutlook(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  const rows = (book.reading?.tracks ?? []).map((track) => rateOutlook(track)).filter((row) => row != null).slice(0, 3);
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "Outlook", 0.65, 0.4, "C23FC2", type.kicker);
  slide.addText("One period ahead", {
    x: 0.65,
    y: 0.78,
    w: 12,
    h: 0.5,
    fontFace: "Calibri",
    fontSize: type.title - 8,
    bold: true,
    color: INK,
    margin: 0,
  });
  const width = rows.length > 1 ? 5.8 : 12;
  rows.forEach((row, index) => {
    const x = 0.65 + index * 6.2;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x,
      y: 1.7,
      w: width,
      h: 3.5,
      fill: { color: "F3F0F8" },
      rectRadius: 0.12,
    });
    slide.addText(row.name, {
      x: x + 0.35,
      y: 1.95,
      w: width - 0.7,
      h: 0.35,
      fontFace: "Calibri",
      fontSize: 14,
      bold: true,
      color: MUTED,
      margin: 0,
    });
    slide.addText(`${row.rate}%`, {
      x: x + 0.35,
      y: 2.5,
      w: width - 0.7,
      h: 0.9,
      fontFace: "Calibri",
      fontSize: 40,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText(`${row.label}. Line from ${row.from} to ${row.to}.`, {
      x: x + 0.35,
      y: 3.6,
      w: width - 0.7,
      h: 0.8,
      fontFace: "Calibri",
      fontSize: 16,
      color: INK,
      margin: 0,
    });
  });
  slide.addText("Projection is a straight line from the full months. It is not a result. A short month and a tiny month are left out of the line.", {
    x: 0.65,
    y: 5.5,
    w: 12,
    h: 0.8,
    fontFace: "Calibri",
    fontSize: 14,
    color: MUTED,
    margin: 0,
  });
}

function addGridColumn(pres: Pres, book: Book, index: number, type: (typeof SCALE)["m"]) {
  const grid = book.grid;
  const column = grid?.columns[index];
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  if (!grid || !column) return;
  const points = columnPoints(grid, index);
  const peak = Math.max(0, ...points.map((point) => point.value));
  const max = column.percent && peak <= 100 ? 100 : axisMax(peak);
  const latest = [...points].reverse().find((point) => !point.projected);
  kicker(slide, grid.cadence === "rows" ? "By row" : grid.cadence === "year" ? "By year" : "By month", 0.55, 0.32, "C23FC2", type.kicker);
  slide.addText(column.name, {
    x: 0.55,
    y: 0.64,
    w: 7,
    h: 0.5,
    fontFace: "Calibri",
    fontSize: type.title - 8,
    bold: true,
    color: INK,
    margin: 0,
  });
  if (latest) {
    slide.addText(column.percent ? `${formatGridPoint(latest.value)}%` : formatGridPoint(latest.value), {
      x: 7.6,
      y: 0.4,
      w: 5.1,
      h: 0.6,
      fontFace: "Calibri",
      fontSize: 28,
      bold: true,
      color: INK,
      align: "right",
      margin: 0,
    });
    slide.addText(`Latest row, ${latest.label}. Not a total.`, {
      x: 6.2,
      y: 1.05,
      w: 6.5,
      h: 0.3,
      fontFace: "Calibri",
      fontSize: 13,
      color: MUTED,
      align: "right",
      margin: 0,
    });
  }
  const fit = chartFit(grid.labels, [{ name: column.name, values: column.values, percent: column.percent }]);
  if (fit.kind === "line" || fit.kind === "bars" || fit.kind === "donut") {
    const drawn = fit.kind === "donut" ? points.filter((point) => !point.projected) : points;
    const labels = drawn.map((point) => point.label);
    const values = drawn.map((point) => point.value);
    const chartType = fit.kind === "line" ? pres.charts.LINE : fit.kind === "donut" ? pres.charts.DOUGHNUT : pres.charts.BAR;
    if (fit.kind === "donut") {
      slide.addChart(chartType, [{ name: column.name, labels, values }], {
        x: 2.4,
        y: 1.45,
        w: 8.4,
        h: 4.8,
        showLegend: true,
        legendPos: "r",
        holeSize: 62,
        chartColors: ["FF96FF", "7C8CFF", "B7AFC9", "D4CEDF"],
        chartArea: { fill: { color: WHITE } },
        plotArea: { fill: { color: WHITE } },
      });
    } else if (fit.kind === "line") {
      slide.addChart(chartType, [{ name: column.name, labels, values }], {
        x: 0.4,
        y: 1.5,
        w: 12.5,
        h: 4.9,
        showLegend: false,
        showValue: false,
        lineSize: 2.5,
        chartColors: [BRAND],
        chartArea: { fill: { color: WHITE } },
        plotArea: { fill: { color: WHITE } },
        catAxisLabelColor: MUTED,
        valAxisLabelColor: MUTED,
        catAxisLabelFontSize: labels.length > 12 ? 8 : 11,
        catAxisLabelFontFace: "Calibri",
        valAxisLabelFontFace: "Calibri",
        valAxisMinVal: 0,
        valAxisMaxVal: max,
      });
    } else {
      slide.addChart(chartType, [{ name: column.name, labels: [...labels].reverse(), values: [...values].reverse() }], {
        x: 0.4,
        y: 1.45,
        w: 12.5,
        h: 4.9,
        barDir: "bar",
        barGrouping: "clustered",
        showLegend: false,
        chartColors: [BRAND],
        chartArea: { fill: { color: WHITE } },
        plotArea: { fill: { color: WHITE } },
        catAxisLabelColor: MUTED,
        valAxisLabelColor: MUTED,
        catAxisLabelFontFace: "Calibri",
        valAxisLabelFontFace: "Calibri",
        valAxisMinVal: 0,
        valAxisMaxVal: max,
      });
    }
    slide.addText(gridChartNote(grid, index), {
      x: 0.55,
      y: 6.85,
      w: 12.2,
      h: 0.45,
      fontFace: "Calibri",
      fontSize: 12,
      color: MUTED,
      margin: 0,
    });
    return;
  }
  const plotX = 0.7;
  const plotY = 1.7;
  const plotH = 4.15;
  const plotW = 12.1;
  const gap = points.length > 16 ? 0.04 : 0.08;
  const barW = points.length > 0 ? Math.max(0.12, (plotW - gap * Math.max(0, points.length - 1)) / points.length) : 0.2;
  points.forEach((point, pointIndex) => {
    const x = plotX + pointIndex * (barW + gap);
    const height = point.value <= 0 || max <= 0 ? 0.04 : Math.max(0.06, plotH * Math.min(1, point.value / max));
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + Math.max(0, (barW - 0.28) / 2),
      y: plotY + (plotH - height),
      w: Math.min(0.28, barW),
      h: height,
      fill: { color: point.projected ? "C23FC2" : BRAND },
      rectRadius: 0.04,
    });
    slide.addText(point.label, {
      x,
      y: plotY + plotH + 0.04,
      w: barW,
      h: 0.42,
      fontFace: "Calibri",
      fontSize: 9,
      color: MUTED,
      align: "center",
      margin: 0,
    });
  });
  slide.addText(gridChartNote(grid, index), {
    x: 0.55,
    y: 6.85,
    w: 12.2,
    h: 0.45,
    fontFace: "Calibri",
    fontSize: 12,
    color: MUTED,
    margin: 0,
  });
}

function addCombo(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  const grid = book.grid;
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  if (!grid) return;
  const pair = comboPair(
    grid.labels,
    grid.columns.map((column) => ({ name: column.name, values: column.values, percent: column.percent })),
  );
  if (!pair) return;
  const count = grid.columns[pair.count];
  const rate = grid.columns[pair.rate];
  if (!count || !rate) return;
  const labels: string[] = [];
  const countValues: number[] = [];
  const rateValues: number[] = [];
  grid.labels.forEach((label, index) => {
    const countValue = count.values[index];
    const rateValue = rate.values[index];
    if (countValue == null || rateValue == null) return;
    labels.push(label);
    countValues.push(countValue);
    rateValues.push(rateValue);
  });
  const countMax = axisMax(Math.max(0, ...countValues));
  kicker(slide, "Together", 0.55, 0.32, "C23FC2", type.kicker);
  slide.addText(`${count.name} and ${rate.name}`, {
    x: 0.55,
    y: 0.64,
    w: 12,
    h: 0.5,
    fontFace: "Calibri",
    fontSize: type.title - 10,
    bold: true,
    color: INK,
    margin: 0,
  });
  const combo = [
    {
      type: pres.charts.BAR,
      data: [{ name: count.name, labels, values: countValues }],
      options: { chartColors: [BRAND], barGrouping: "clustered", barDir: "col" },
    },
    {
      type: pres.charts.LINE,
      data: [{ name: rate.name, labels, values: rateValues }],
      options: { chartColors: ["C23FC2"], secondaryValAxis: true },
    },
  ];
  const comboOptions = {
    x: 0.35,
    y: 1.3,
    w: 12.6,
    h: 5.1,
    showLegend: true,
    legendPos: "b" as const,
    valAxes: [
      { valAxisMinVal: 0, valAxisMaxVal: countMax, valAxisTitle: count.name, showValAxisTitle: true, valAxisLabelColor: MUTED },
      { valAxisMinVal: 0, valAxisMaxVal: 100, valAxisTitle: rate.name, showValAxisTitle: true, valAxisLabelColor: MUTED },
    ],
    catAxes: [{ catAxisLabelColor: MUTED, catAxisLabelFontSize: labels.length > 10 ? 9 : 11 }, { catAxisHidden: true }],
  };
  slide.addChart(combo, comboOptions as unknown as []);
  slide.addText(`${count.name} is the column. ${rate.name} is the line. Rows are not added together.`, {
    x: 0.55,
    y: 6.85,
    w: 12.2,
    h: 0.4,
    fontFace: "Calibri",
    fontSize: 12,
    color: MUTED,
    margin: 0,
  });
}

function addCard(pres: Pres, book: Book, index: number, type: (typeof SCALE)["m"]) {
  const card = book.outline?.cards[index];
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  if (!card) return;
  kicker(slide, card.kicker, 0.7, 0.55, "C23FC2", type.kicker);
  slide.addText(card.title, {
    x: 0.7,
    y: 1.05,
    w: 11.8,
    h: 1.4,
    fontFace: "Calibri",
    fontSize: card.title.length > 28 ? type.title - 8 : type.title,
    bold: true,
    color: INK,
    margin: 0,
  });
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.7,
    y: 2.7,
    w: 11.9,
    h: 3.15,
    fill: { color: "F3F0F8" },
    rectRadius: 0.16,
  });
  slide.addText(card.hint.trim() || "Add the point to mention.", {
    x: 1.05,
    y: 3.05,
    w: 11.2,
    h: 2.45,
    fontFace: "Calibri",
    fontSize: card.hint.length > 180 ? 18 : 24,
    color: INK,
    margin: 0,
  });
  slide.addText("Placeholder. No figure was in the brief.", {
    x: 0.7,
    y: 6.7,
    w: 11.8,
    h: 0.35,
    fontFace: "Calibri",
    fontSize: 14,
    color: MUTED,
    margin: 0,
  });
}

function addRead(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  const lines = book.facts
    ? factReadLines(book.facts)
    : book.story
      ? storyReadLines(book.story, book.period || book.story.period)
      : [];
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "What it says", 0.65, 0.4, "C23FC2", type.kicker);
  slide.addText(book.client || "The brief", {
    x: 0.65,
    y: 0.78,
    w: 12,
    h: 0.55,
    fontFace: "Calibri",
    fontSize: type.title - 8,
    bold: true,
    color: INK,
    margin: 0,
  });
  const height = Math.min(1.15, 4.6 / Math.max(lines.length, 1));
  lines.forEach((line, index) => {
    const y = 1.7 + index * (height + 0.12);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.65,
      y,
      w: 12,
      h: height,
      fill: { color: "F3F0F8" },
      rectRadius: 0.1,
    });
    slide.addText(line, {
      x: 0.9,
      y: y + 0.12,
      w: 11.5,
      h: height - 0.24,
      fontFace: "Calibri",
      fontSize: line.length > 140 ? 14 : 16,
      color: INK,
      margin: 0,
    });
  });
}

function addResult(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  const story = book.story;
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "The result", 0.65, 0.38, "C23FC2", type.kicker);
  slide.addText(book.client || story?.client || "Untitled", {
    x: 0.65,
    y: 0.72,
    w: 12,
    h: 0.7,
    fontFace: "Calibri",
    fontSize: type.title,
    bold: true,
    color: INK,
    margin: 0,
  });
  const proofs = story?.proofs.slice(0, 4) ?? [];
  let cursor = 1.6;
  if (story && story.volume != null && story.volume > 0) {
    slide.addText(compactVolume(story.volume), {
      x: 0.65,
      y: 1.55,
      w: 3.4,
      h: 1.15,
      fontFace: "Calibri",
      fontSize: type.rate,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText(story.unit || "Count", {
      x: 4.15,
      y: 1.78,
      w: 8,
      h: 0.42,
      fontFace: "Calibri",
      fontSize: type.body + 6,
      bold: true,
      color: INK,
      margin: 0,
    });
    const when = [story.hedge, story.verb, book.period || story.period].filter(Boolean).join("  ·  ");
    if (when) {
      slide.addText(when, {
        x: 4.15,
        y: 2.24,
        w: 8,
        h: 0.32,
        fontFace: "Calibri",
        fontSize: 15,
        color: MUTED,
        margin: 0,
      });
    }
    cursor = 3.15;
  }
  const columns = proofs.length >= 4 ? 2 : Math.max(proofs.length, 1);
  const rows = Math.ceil(proofs.length / columns) || 1;
  const gap = 0.22;
  const width = (12 - gap * (columns - 1)) / columns;
  const height = rows === 1 ? 2.35 : 1.7;
  proofs.forEach((proof, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const x = 0.65 + column * (width + gap);
    const y = cursor + row * (height + 0.16);
    drawProof(pres, slide, proof, x, y, width, height, type);
  });
  const note = proofs.some((proof) => proof.relation !== "exact")
    ? "Each bar stops at the figure in the brief. Above means at least that high."
    : "Each bar is the rate stated in the brief.";
  if (proofs.length > 0) {
    const noteY = Math.min(6.95, cursor + rows * height + (rows - 1) * 0.16 + 0.2);
    slide.addText(note, {
      x: 0.65,
      y: noteY,
      w: 12,
      h: 0.32,
      fontFace: "Calibri",
      fontSize: 13,
      color: MUTED,
      margin: 0,
    });
  }
}

function addFacts(pres: Pres, book: Book, index: number, type: (typeof SCALE)["m"]) {
  const group = book.facts?.groups[index];
  if (!group) return;
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, group.name || "The result", 0.65, 0.38, "C23FC2", type.kicker);
  slide.addText(book.client || "Untitled", {
    x: 0.65,
    y: 0.72,
    w: 12,
    h: 0.7,
    fontFace: "Calibri",
    fontSize: type.title,
    bold: true,
    color: INK,
    margin: 0,
  });
  const proofs = group.rates.slice(0, 4).map((rate) => ({ label: rate.label, relation: rate.relation, value: rate.value }));
  let cursor = 1.6;
  if (group.volume && group.volume.value > 0) {
    const floor = group.volume.relation === "least" ? "at least" : group.volume.relation === "exact" ? "" : group.volume.relation;
    slide.addText(compactVolume(group.volume.value), {
      x: 0.65,
      y: 1.55,
      w: 3.4,
      h: 1.15,
      fontFace: "Calibri",
      fontSize: type.rate,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText(group.volume.unit || "Count", {
      x: 4.15,
      y: 1.78,
      w: 8,
      h: 0.42,
      fontFace: "Calibri",
      fontSize: type.body + 6,
      bold: true,
      color: INK,
      margin: 0,
    });
    const when = [group.volume.hedge, floor, book.facts?.verb, book.period].filter(Boolean).join("  ·  ");
    if (when) {
      slide.addText(when, {
        x: 4.15,
        y: 2.24,
        w: 8,
        h: 0.32,
        fontFace: "Calibri",
        fontSize: 15,
        color: MUTED,
        margin: 0,
      });
    }
    cursor = 3.15;
  }
  const columns = proofs.length >= 4 ? 2 : Math.max(proofs.length, 1);
  const rows = Math.ceil(proofs.length / columns) || 1;
  const gap = 0.22;
  const width = (12 - gap * (columns - 1)) / columns;
  const height = rows === 1 ? 2.35 : 1.7;
  proofs.forEach((proof, proofIndex) => {
    const column = proofIndex % columns;
    const row = Math.floor(proofIndex / columns);
    const x = 0.65 + column * (width + gap);
    const y = cursor + row * (height + 0.16);
    drawProof(pres, slide, proof, x, y, width, height, type);
  });
  const separate = (book.facts?.groups.length ?? 0) > 1;
  const note = separate
    ? "This slide is only this series. Calls and chats are not added together."
    : proofs.some((proof) => proof.relation !== "exact")
      ? "Each bar stops at the figure in the brief. Above means at least that high."
      : proofs.length > 0
        ? "Each bar is the rate stated in the brief."
        : "The figure is the one written in the brief.";
  const noteY = proofs.length > 0 ? Math.min(6.95, cursor + rows * height + (rows - 1) * 0.16 + 0.2) : Math.min(6.95, cursor + 0.4);
  slide.addText(note, {
    x: 0.65,
    y: noteY,
    w: 12,
    h: 0.32,
    fontFace: "Calibri",
    fontSize: 13,
    color: MUTED,
    margin: 0,
  });
}

function drawProof(
  pres: Pres,
  slide: Slide,
  proof: Proof,
  x: number,
  y: number,
  w: number,
  h: number,
  type: (typeof SCALE)["m"],
) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x,
    y,
    w,
    h,
    fill: { color: "F7F5FB" },
    rectRadius: 0.1,
  });
  slide.addText(proof.label, {
    x: x + 0.28,
    y: y + 0.2,
    w: w - 0.5,
    h: 0.3,
    fontFace: "Calibri",
    fontSize: 14,
    color: MUTED,
    margin: 0,
  });
  slide.addText(proofPhrase(proof), {
    x: x + 0.28,
    y: y + 0.52,
    w: w - 0.5,
    h: 0.55,
    fontFace: "Calibri",
    fontSize: h < 2 ? 22 : type.title - 4,
    bold: true,
    color: INK,
    margin: 0,
  });
  const track = w - 0.56;
  const fill = Math.max(0.08, track * (Math.min(100, Math.max(0, proof.value)) / 100));
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.28,
    y: y + h - 0.48,
    w: track,
    h: 0.14,
    fill: { color: "F3F0F8" },
    rectRadius: 0.07,
  });
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.28,
    y: y + h - 0.48,
    w: fill,
    h: 0.14,
    fill: { color: BRAND },
    rectRadius: 0.07,
  });
}

function addBook(
  pres: Pres,
  book: Book,
  view: ReturnType<typeof viewBook>,
  type: (typeof SCALE)["m"],
) {
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "The book", 0.65, 0.4, "C23FC2", type.kicker);
  slide.addText(book.client || "Untitled book", {
    x: 0.65,
    y: 0.72,
    w: 12,
    h: 0.78,
    fontFace: "Calibri",
    fontSize: type.title,
    bold: true,
    color: INK,
    margin: 0,
  });
  slide.addText(`${view.rate}%`, {
    x: 0.65,
    y: 1.7,
    w: 3.6,
    h: 1.25,
    fontFace: "Calibri",
    fontSize: type.rate,
    bold: true,
    color: INK,
    margin: 0,
  });
  slide.addText("of active subscribers", {
    x: 4.3,
    y: 1.95,
    w: 6,
    h: 0.4,
    fontFace: "Calibri",
    fontSize: type.body + 4,
    color: INK,
    margin: 0,
  });
  slide.addText(`${formatCount(view.claims)} claims on ${formatCount(view.active)} active`, {
    x: 4.3,
    y: 2.4,
    w: 7,
    h: 0.35,
    fontFace: "Calibri",
    fontSize: type.body,
    color: MUTED,
    margin: 0,
  });
  const track = 12;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.65,
    y: 3.05,
    w: track,
    h: 0.16,
    fill: { color: "F3F0F8" },
    rectRadius: 0.08,
  });
  const width = Math.max(view.rate > 0 ? 0.12 : 0, Math.min(track, (Math.min(view.rate, 100) / 100) * track));
  if (width > 0) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.65,
      y: 3.05,
      w: width,
      h: 0.16,
      fill: { color: BRAND },
      rectRadius: 0.08,
    });
  }
  slide.addText(
    view.over
      ? "Claims are higher than the active book. The bar stops at the end of the track."
      : "Bar length is the claim rate, and it stops at 100%.",
    { x: 0.65, y: 3.32, w: 12, h: 0.32, fontFace: "Calibri", fontSize: 13, color: MUTED, margin: 0 },
  );
  const parts = [
    { label: "Fulfilled", count: view.fulfilled, share: view.fulfilledShare, color: PINK },
    { label: "Approved, still open", count: view.approvedOpen, share: view.openShare, color: APPROVED },
    { label: "Not approved", count: view.notApproved, share: view.notShare, color: OPEN },
  ];
  parts.forEach((part, index) => {
    const x = 0.65 + index * 4.15;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x,
      y: 3.85,
      w: 3.95,
      h: 2.9,
      fill: { color: "F7F5FB" },
      rectRadius: 0.1,
    });
    slide.addShape(pres.shapes.OVAL, {
      x: x + 0.28,
      y: 4.13,
      w: 0.18,
      h: 0.18,
      fill: { color: part.color },
    });
    slide.addText(part.label, {
      x: x + 0.58,
      y: 4.07,
      w: 3.1,
      h: 0.32,
      fontFace: "Calibri",
      fontSize: 14,
      color: MUTED,
      margin: 0,
    });
    slide.addText(formatCount(part.count), {
      x: x + 0.28,
      y: 4.55,
      w: 3.4,
      h: 0.85,
      fontFace: "Calibri",
      fontSize: type.title,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText(`${part.share}% of active`, {
      x: x + 0.28,
      y: 5.45,
      w: 3.4,
      h: 0.35,
      fontFace: "Calibri",
      fontSize: 15,
      bold: true,
      color: BRAND,
      margin: 0,
    });
  });
}

function addMix(
  pres: Pres,
  book: Book,
  view: ReturnType<typeof viewBook>,
  options: DeckOptions,
  type: (typeof SCALE)["m"],
) {
  const slide = pres.addSlide();
  slide.background = { color: STAGE };
  kicker(slide, "The mix", 0.65, 0.4, PINK, type.kicker);
  slide.addText(book.client || "The book", {
    x: 0.65,
    y: 0.72,
    w: 6,
    h: 0.72,
    fontFace: "Calibri",
    fontSize: type.title,
    bold: true,
    color: WHITE,
    margin: 0,
  });
  const note =
    options.mix === "columns"
      ? "Column height follows the count. The tallest slice sets the scale."
      : options.mix === "meter"
        ? "The bar is the whole mix. Length is the count."
        : view.ringIsClaimMix
          ? "The ring is the claim mix. The rate is still claims divided by active subscribers."
          : "The ring is the whole book.";
  slide.addText(note, {
    x: 0.65,
    y: 1.45,
    w: 5.4,
    h: 1.1,
    fontFace: "Calibri",
    fontSize: type.body,
    color: "D4D0E0",
    margin: 0,
  });

  const labels = view.slices.map((slice) => slice.label);
  const colors = view.slices.map((slice) => TONE[slice.tone] ?? REST);
  if (view.slices.length === 0) {
    slide.addText("Add a spreadsheet to fill the mix.", {
      x: 0.65,
      y: 3,
      w: 10,
      h: 0.5,
      fontFace: "Calibri",
      fontSize: type.body,
      color: "D4D0E0",
      margin: 0,
    });
    return;
  }

  if (options.mix === "donut") {
    slide.addChart(pres.charts.DOUGHNUT, [{ name: "Mix", labels, values: view.slices.map((slice) => slice.count) }], {
      x: 6.3,
      y: 1.15,
      w: 6.4,
      h: 5.2,
      showLegend: false,
      showPercent: false,
      showValue: false,
      holeSize: 64,
      chartColors: colors,
      chartArea: { fill: { color: STAGE } },
      plotArea: { fill: { color: STAGE } },
    });
    slide.addText(`${view.rate}%`, {
      x: 8.15,
      y: 3.15,
      w: 2.7,
      h: 0.7,
      fontFace: "Calibri",
      fontSize: 32,
      bold: true,
      color: WHITE,
      align: "center",
      margin: 0,
    });
  } else if (options.mix === "columns") {
    slide.addChart(
      pres.charts.BAR,
      view.slices.map((slice) => ({
        name: slice.label,
        labels,
        values: labels.map((label) => (label === slice.label ? slice.count : 0)),
      })),
      {
        x: 6.1,
        y: 1.3,
        w: 6.7,
        h: 5.4,
        barDir: "col",
        barGrouping: "stacked",
        showLegend: false,
        showValue: false,
        chartColors: colors,
        chartArea: { fill: { color: STAGE } },
        plotArea: { fill: { color: STAGE } },
        catAxisLabelColor: "E4E0EE",
        valAxisLabelColor: "A39CB8",
        catAxisLabelFontSize: 11,
        valAxisLabelFontSize: 11,
        valGridLine: { color: "3A3358", size: 0.5 },
        catGridLine: { style: "none" },
      },
    );
  } else {
    slide.addChart(
      pres.charts.BAR,
      view.slices.map((slice) => ({
        name: slice.label,
        labels: ["Mix"],
        values: [slice.count],
      })),
      {
        x: 0.4,
        y: 2.7,
        w: 12.5,
        h: 2.4,
        barDir: "bar",
        barGrouping: "stacked",
        showLegend: false,
        showValue: false,
        chartColors: colors,
        chartArea: { fill: { color: STAGE } },
        plotArea: { fill: { color: STAGE } },
        catAxisHidden: true,
        valAxisHidden: true,
        valGridLine: { style: "none" },
        catGridLine: { style: "none" },
      },
    );
  }

  view.slices.forEach((slice, index) => {
    const y = 2.75 + index * 0.68;
    if (options.mix === "meter") return;
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7,
      y: y + 0.08,
      w: 0.16,
      h: 0.16,
      fill: { color: TONE[slice.tone] ?? REST },
    });
    slide.addText(slice.label, {
      x: 1.05,
      y,
      w: 3.2,
      h: 0.32,
      fontFace: "Calibri",
      fontSize: 15,
      color: WHITE,
      margin: 0,
    });
    slide.addText(`${formatCount(slice.count)}   ${slice.percent}%`, {
      x: 1.05,
      y: y + 0.34,
      w: 3.4,
      h: 0.26,
      fontFace: "Calibri",
      fontSize: 12,
      color: "C8C4D4",
      margin: 0,
    });
  });
  if (options.mix === "meter") {
    view.slices.forEach((slice, index) => {
      const x = 0.7 + (index % 4) * 3.1;
      const y = 5.5 + Math.floor(index / 4) * 0.4;
      slide.addShape(pres.shapes.OVAL, {
        x,
        y: y + 0.06,
        w: 0.16,
        h: 0.16,
        fill: { color: TONE[slice.tone] ?? REST },
      });
      slide.addText(`${slice.label}  ${slice.percent}%`, {
        x: x + 0.28,
        y,
        w: 2.7,
        h: 0.3,
        fontFace: "Calibri",
        fontSize: 13,
        color: WHITE,
        margin: 0,
      });
    });
  }
}

function addVersus(
  pres: Pres,
  book: Book,
  view: ReturnType<typeof viewBook>,
  type: (typeof SCALE)["m"],
) {
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "Versus", 0.65, 0.4, "C23FC2", type.kicker);
  slide.addText(`${book.client || "This book"}  and  ${book.peerName || "the comparison"}`, {
    x: 0.65,
    y: 0.78,
    w: 12,
    h: 0.55,
    fontFace: "Calibri",
    fontSize: type.title - 4,
    bold: true,
    color: INK,
    margin: 0,
  });
  const rows = [
    { label: "Claim rate", note: "of active subscribers", left: view.rate, right: view.peerRate },
    { label: "Fulfilled", note: "of active subscribers", left: view.fulfilledShare, right: view.peerFulfilledRate },
  ];
  rows.forEach((row, index) => {
    const y = 1.7 + index * 2.4;
    slide.addText(row.label, {
      x: 0.65,
      y,
      w: 12,
      h: 0.36,
      fontFace: "Calibri",
      fontSize: 16,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText(row.note, {
      x: 0.65,
      y: y + 0.38,
      w: 6,
      h: 0.28,
      fontFace: "Calibri",
      fontSize: 13,
      color: MUTED,
      margin: 0,
    });
    drawSide(pres, slide, 0.65, y + 0.82, book.client || "This book", row.left, row.left >= row.right);
    drawSide(pres, slide, 6.95, y + 0.82, book.peerName || "Peer", row.right, row.right >= row.left);
  });
}

function drawSide(pres: Pres, slide: Slide, x: number, y: number, name: string, rate: number, lead: boolean) {
  slide.addText(`${rate}%`, {
    x,
    y,
    w: 2.2,
    h: 0.5,
    fontFace: "Calibri",
    fontSize: 28,
    bold: true,
    color: lead ? BRAND : INK,
    margin: 0,
  });
  slide.addText(name, {
    x: x + 2.3,
    y: y + 0.08,
    w: 3.4,
    h: 0.36,
    fontFace: "Calibri",
    fontSize: 14,
    color: MUTED,
    margin: 0,
  });
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x,
    y: y + 0.6,
    w: 5.6,
    h: 0.14,
    fill: { color: "F3F0F8" },
    rectRadius: 0.07,
  });
  const width = Math.max(rate > 0 ? 0.1 : 0, Math.min(5.6, (Math.min(rate, 100) / 100) * 5.6));
  if (width > 0) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x,
      y: y + 0.6,
      w: width,
      h: 0.14,
      fill: { color: lead ? BRAND : "B7AFC9" },
      rectRadius: 0.07,
    });
  }
}

function addListening(pres: Pres, book: Book, options: DeckOptions, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  const talk = conversationOf(book.listening);
  kicker(slide, "Health", 0.65, 0.38, "C23FC2", type.kicker);
  slide.addText(`${talk.health}%`, {
    x: 0.6,
    y: 0.7,
    w: 4.2,
    h: 1.15,
    fontFace: "Calibri",
    fontSize: type.rate,
    bold: true,
    color: INK,
    margin: 0,
  });
  slide.addText("of mentions are positive or neutral.", {
    x: 0.65,
    y: 1.9,
    w: 4.4,
    h: 0.7,
    fontFace: "Calibri",
    fontSize: type.body,
    color: MUTED,
    margin: 0,
  });
  slide.addText(`${formatCount(talk.total)} mentions`, {
    x: 0.65,
    y: 2.6,
    w: 4,
    h: 0.32,
    fontFace: "Calibri",
    fontSize: 14,
    color: FAINT,
    margin: 0,
  });
  const labels = book.listening.map((row) => row.label);
  if (options.listening === "columns") {
    slide.addChart(
      pres.charts.BAR,
      [
        { name: "Negative", labels, values: book.listening.map((row) => row.negative) },
        { name: "Neutral", labels, values: book.listening.map((row) => row.neutral) },
        { name: "Positive", labels, values: book.listening.map((row) => row.positive) },
      ],
      {
        x: 5.1,
        y: 0.45,
        w: 7.7,
        h: 6.4,
        barDir: "col",
        barGrouping: "stacked",
        chartColors: [NEG, NEU, POS],
        showLegend: true,
        legendPos: "b",
        showValue: false,
        chartArea: { fill: { color: WHITE } },
        catAxisLabelColor: MUTED,
        valAxisLabelColor: FAINT,
        valGridLine: { color: LINE, size: 0.5 },
        catGridLine: { style: "none" },
      },
    );
    return;
  }
  book.listening.slice(0, 6).forEach((row, index) => {
    const y = 3.15 + index * 0.65;
    const max = Math.max(1, ...book.listening.map((item) => item.count));
    const width = Math.max(0.15, (row.count / max) * 7.2);
    slide.addText(row.label, {
      x: 5.2,
      y,
      w: 3.2,
      h: 0.28,
      fontFace: "Calibri",
      fontSize: 14,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText(formatCount(row.count), {
      x: 10.6,
      y,
      w: 1.8,
      h: 0.28,
      fontFace: "Calibri",
      fontSize: 14,
      color: MUTED,
      align: "right",
      margin: 0,
    });
    const x = 5.2;
    const barY = y + 0.32;
    let cursor = x;
    const parts = [
      { n: row.negative, color: NEG },
      { n: row.neutral, color: NEU },
      { n: row.positive, color: POS },
    ];
    const sum = parts.reduce((total, part) => total + part.n, 0) || 1;
    parts.forEach((part) => {
      if (part.n <= 0) return;
      const partW = (part.n / sum) * width;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: cursor,
        y: barY,
        w: partW,
        h: 0.12,
        fill: { color: part.color },
      });
      cursor += partW;
    });
  });
}

function addJourney(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "Journey", 0.65, 0.38, "C23FC2", type.kicker);
  slide.addText("Where the tone turns", {
    x: 0.65,
    y: 0.72,
    w: 10,
    h: 0.5,
    fontFace: "Calibri",
    fontSize: type.title,
    bold: true,
    color: INK,
    margin: 0,
  });
  const rows = book.journey.slice(0, 5);
  rows.forEach((stage, index) => {
    const y = 1.55 + index * 1.05;
    slide.addText(stage.label, {
      x: 0.65,
      y,
      w: 2.5,
      h: 0.32,
      fontFace: "Calibri",
      fontSize: 16,
      bold: true,
      color: INK,
      margin: 0,
    });
    slide.addText(`${formatCount(stage.mentions)} mentions`, {
      x: 0.65,
      y: y + 0.3,
      w: 2.5,
      h: 0.24,
      fontFace: "Calibri",
      fontSize: 12,
      color: MUTED,
      margin: 0,
    });
    const half = 3.6;
    const cx = 3.5 + half;
    const neu = (stage.neutral / 2 / 100) * half;
    const neg = (stage.negative / 100) * half;
    const pos = (stage.positive / 100) * half;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 3.5,
      y: y + 0.18,
      w: half * 2,
      h: 0.16,
      fill: { color: "F3F0F8" },
      rectRadius: 0.06,
    });
    if (neg > 0) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: cx - neu - neg,
        y: y + 0.18,
        w: neg,
        h: 0.16,
        fill: { color: NEG },
      });
    }
    if (neu > 0) {
      slide.addShape(pres.shapes.RECTANGLE, { x: cx - neu, y: y + 0.18, w: neu, h: 0.16, fill: { color: NEU } });
      slide.addShape(pres.shapes.RECTANGLE, { x: cx, y: y + 0.18, w: neu, h: 0.16, fill: { color: NEU } });
    }
    if (pos > 0) {
      slide.addShape(pres.shapes.RECTANGLE, { x: cx + neu, y: y + 0.18, w: pos, h: 0.16, fill: { color: POS } });
    }
    slide.addShape(pres.shapes.RECTANGLE, { x: cx - 0.01, y: y + 0.08, w: 0.02, h: 0.36, fill: { color: INK } });
    slide.addText(`${stage.negative}%`, {
      x: 10.9,
      y: y + 0.08,
      w: 0.9,
      h: 0.3,
      fontFace: "Calibri",
      fontSize: 14,
      bold: true,
      color: "D4254C",
      margin: 0,
    });
    slide.addText(`${stage.positive}%`, {
      x: 11.8,
      y: y + 0.08,
      w: 0.9,
      h: 0.3,
      fontFace: "Calibri",
      fontSize: 14,
      bold: true,
      color: "0F6E4C",
      margin: 0,
    });
  });
}

function addMonths(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "Months", 0.65, 0.36, "C23FC2", type.kicker);
  slide.addText(book.client || "Claims by month", {
    x: 0.65,
    y: 0.68,
    w: 10,
    h: 0.48,
    fontFace: "Calibri",
    fontSize: type.title,
    bold: true,
    color: INK,
    margin: 0,
  });
  const labels = book.months.map((month) => month.label);
  slide.addChart(
    pres.charts.BAR,
    [
      { name: "Not approved", labels, values: book.months.map((month) => month.notApproved) },
      { name: "Approved, still open", labels, values: book.months.map((month) => month.open) },
      { name: "Fulfilled", labels, values: book.months.map((month) => month.fulfilled) },
    ],
    {
      x: 0.45,
      y: 1.3,
      w: 12.4,
      h: 5.7,
      barDir: "col",
      barGrouping: "stacked",
      chartColors: [OPEN, APPROVED, PINK],
      showLegend: true,
      legendPos: "b",
      showValue: false,
      chartArea: { fill: { color: WHITE } },
      catAxisLabelColor: MUTED,
      valAxisLabelColor: FAINT,
      valGridLine: { color: LINE, size: 0.5 },
      catGridLine: { style: "none" },
    },
  );
}

function addSource(pres: Pres, book: Book, type: (typeof SCALE)["m"]) {
  if (!hasBroughtIn(book)) return;
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  kicker(slide, "Brought in", 0.65, 0.4, "C23FC2", type.kicker);
  slide.addText(book.sourceName || "Decoded text", {
    x: 0.65,
    y: 0.78,
    w: 12,
    h: 0.5,
    fontFace: "Calibri",
    fontSize: type.title - 6,
    bold: true,
    color: INK,
    margin: 0,
  });
  const text = book.sourceText.trim();
  const shown = text.length > 1600 ? `${text.slice(0, 1600).trimEnd()}…` : text;
  slide.addText(shown, {
    x: 0.65,
    y: 1.5,
    w: 12,
    h: 5.4,
    fontFace: "Calibri",
    fontSize: shown.length > 700 ? 14 : type.body,
    color: INK,
    margin: 0,
    valign: "top",
  });
}

function addNext(
  pres: Pres,
  book: Book,
  view: ReturnType<typeof viewBook>,
  type: (typeof SCALE)["m"],
) {
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 4.6, h: 7.5, fill: { color: STAGE } });
  kicker(slide, "What to do", 0.45, 0.5, PINK, type.kicker);
  slide.addText(book.client || "This book", {
    x: 0.45,
    y: 1.05,
    w: 3.8,
    h: 0.8,
    fontFace: "Calibri",
    fontSize: 22,
    bold: true,
    color: WHITE,
    margin: 0,
  });
  const seriesLead = isSeries(book) ? book.reading?.tracks[0] : undefined;
  const seriesTotals = seriesLead ? trackTotals(seriesLead) : null;
  const factVolume = isFacts(book) ? book.facts?.groups.find((group) => (group.volume?.value ?? 0) > 0)?.volume : undefined;
  const results = !seriesTotals && !factVolume && isResults(book) && book.story?.volume != null && book.story.volume > 0;
  slide.addText(seriesTotals ? `${seriesTotals.rate}%` : factVolume ? compactVolume(factVolume.value) : results ? compactVolume(book.story?.volume ?? 0) : `${view.rate}%`, {
    x: 0.45,
    y: 2.3,
    w: 3.8,
    h: 0.9,
    fontFace: "Calibri",
    fontSize: type.rate - 8,
    bold: true,
    color: WHITE,
    margin: 0,
  });
  slide.addText(
    seriesTotals ? `${seriesLead?.name ?? "Series"} containment` : factVolume ? factVolume.unit || "Count" : results ? book.story?.unit || "Count" : "of active subscribers",
    {
    x: 0.45,
    y: 3.25,
    w: 3.7,
    h: 0.35,
    fontFace: "Calibri",
    fontSize: 14,
    color: "D4D0E0",
    margin: 0,
  });
  const text = book.recommendation.trim() || "Write the one line. The deck will not invent a recommendation.";
  slide.addText(text, {
    x: 5.15,
    y: 1.5,
    w: 7.5,
    h: 4.6,
    fontFace: "Calibri",
    fontSize: text.length > 180 ? type.body : type.title - 8,
    bold: true,
    color: text ? INK : FAINT,
    margin: 0,
    valign: "middle",
  });
}

const CHART_COLORS = ["3200BE", "9A2F9A", "1F9E72", "5C6BC0"];
const TRACK = "F3F0F8";

function speak(slide: Slide, source: SourceSlide) {
  const text = source.spoken?.trim();
  if (text) slide.addNotes(text);
}

function addOrigin(pres: Pres, book: Book, index: number, type: (typeof SCALE)["m"]) {
  const source = book.sourceDeck?.slides.find((slide) => slide.index === index) ?? null;
  if (!source) return;
  const role = slideRole(source);
  const structured = (source.pieces ?? []).some((piece) => piece.kind === "card" || piece.kind === "list");
  if (role === "cover") addOriginCover(pres, book, source, type);
  else if (source.table && !source.chart) addOriginTable(pres, source, type);
  else if (role !== "chart" && structured) addOriginPieces(pres, source, type);
  else if (role === "chart" && source.chart && slideShow(source) === "table") addOriginTable(pres, source, type);
  else if (role === "chart" && source.chart) addOriginChart(pres, source, type);
  else if (role === "figures") addOriginNumbers(pres, source, type);
  else if (role === "pictures" && source.picture && source.images.length === 0) addOriginProse(pres, source, type, true);
  else addOriginSheet(pres, source, type);
}

function addOriginPieces(pres: Pres, source: SourceSlide, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  speak(slide, source);
  slide.background = { color: WHITE };
  const frame = pieceFrame(pieceFit(source));
  const size = frame.type;
  const x0 = 0.58;
  const width = 12.17;
  const floor = 7.5 - 36 / 96;
  const pieces = source.pieces ?? [];
  const composition = directedComposition(source);
  const air = pieceAir(source) && source.images.length === 0;
  let cursor = frame.inch(40);
  if (isContinued(source)) {
    kicker(slide, "Continued", x0, cursor, "C23FC2", type.kicker);
    cursor += frame.inch(38);
  }
  if (source.eyebrow) {
    slide.addText(source.eyebrow.toUpperCase(), {
      x: x0, y: cursor, w: width, h: 0.3,
      fontFace: "Calibri", fontSize: 12, bold: true, color: "C23FC2", margin: 0, charSpacing: 1.1,
    });
    cursor += frame.inch(28);
  }
  if (source.title.trim()) {
    const chars = Math.max(8, Math.floor(1168 / (size.title * 0.56)));
    const lines = source.title.trim().split("\n").reduce((total, line) => total + Math.max(1, Math.ceil(line.trim().length / chars)), 0);
    const titleH = frame.inch(lines * Math.ceil(size.title * 1.12));
    slide.addText(source.title.trim(), {
      x: x0, y: cursor, w: width, h: titleH,
      fontFace: "Calibri", fontSize: frame.pt(size.title), bold: true, color: INK, margin: 0, valign: "top",
    });
    cursor += titleH + frame.inch(16);
  }
  const bandInches = (band: PieceBand, narrow: boolean): number => {
    if (band.kind === "prose") {
      const parts = frame.proseParts(band.text);
      let height = 0;
      parts.forEach((part, index) => {
        if (index) height += frame.inch(10);
        height += frame.inch(Math.max(part.height, 18) * (narrow && part.kind !== "list" ? 1.8 : 1));
      });
      return height + frame.inch(16);
    }
    if (band.kind === "lists") {
      if (narrow) {
        const timed = band.items.filter((item) => isTimeTitle(item.title)).length >= 2;
        return band.items.reduce((total, item) => {
          const step = timed ? 0.4 : 0;
          return total + frame.inch(frame.listHeight([item])) + step + frame.inch(12);
        }, 0);
      }
      return frame.inch(frame.listHeight(band.items)) + frame.inch(16);
    }
    const across = narrow ? 1 : gridAcross(band.tone, band.items.length);
    const gap = frame.inch(14);
    let height = 0;
    for (let row = 0; row < band.items.length; row += across) {
      const slice = band.items.slice(row, row + across);
      height += frame.inch(Math.max(...slice.map((item) => frame.cardHeight(item, narrow ? 2 : slice.length)))) + gap;
    }
    return height + frame.inch(2);
  };
  const paint = (items: SlidePiece[], xStart: number, yStart: number, box: number, narrow: boolean, grow: "none" | "cards" | "lists" = "none") => {
    let topCursor = yStart;
    const drawn = pieceBands(items);
    const fillAt = grow === "lists"
      ? drawn.findIndex((band) => band.kind === "lists")
      : grow === "cards"
        ? fillBandIndex(drawn)
        : -1;
    const natural = drawn.reduce((sum, band) => sum + bandInches(band, narrow), 0);
    const extra = fillAt >= 0 ? Math.max(0, floor - (yStart + natural)) : 0;
    for (let bandIndex = 0; bandIndex < drawn.length; bandIndex += 1) {
      const band = drawn[bandIndex];
      if (!band) continue;
      const closing = bandIndex === drawn.length - 1 && bandIndex > 0 && band.kind === "prose" && band.text.length <= 240;
      if (band.kind === "prose") {
        const parts = frame.proseParts(band.text);
        parts.forEach((part, index) => {
          if (index) topCursor += frame.inch(10);
          const height = frame.inch(Math.max(part.height, 18) * (narrow && part.kind !== "list" ? 1.8 : 1));
          const marked = closing && part.kind !== "list";
          if (marked) {
            slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
              x: xStart, y: topCursor, w: box, h: height,
              fill: { color: "F7F5FB" }, rectRadius: CARD_RADIUS,
              line: { color: CARD_LINE, width: 0.75 },
            });
            slide.addShape(pres.shapes.RECTANGLE, {
              x: xStart, y: topCursor, w: 0.06, h: height, fill: { color: BRAND },
            });
          }
          if (part.kind === "list") {
            slide.addText(part.points.map((point, pointIndex) => ({
              text: point,
              options: { bullet: { code: "25CF", color: "E91E8C" }, breakLine: pointIndex < part.points.length - 1, color: "322C49" },
            })), {
              x: xStart, y: topCursor, w: box, h: height,
              fontFace: "Calibri", fontSize: frame.pt(size.bullet), paraSpaceAfter: 4, margin: 0, valign: "top",
            });
          } else {
            slide.addText(part.text, {
              x: xStart + (marked ? 0.16 : 0), y: topCursor, w: box - (marked ? 0.16 : 0), h: height,
              fontFace: "Calibri", fontSize: frame.pt(size.lead), color: "322C49", margin: 0, valign: "middle",
            });
          }
          topCursor += height;
        });
        topCursor += frame.inch(16);
        continue;
      }
      if (band.kind === "lists") {
        const timed = band.items.filter((item) => isTimeTitle(item.title)).length >= 2;
        const columns = narrow ? 1 : Math.max(1, band.items.length);
        const gap = frame.inch(16);
        const columnW = (box - gap * (columns - 1)) / columns;
        const height = frame.inch(frame.listHeight(band.items) * (narrow && band.items.length > 1 ? band.items.length : 1))
          + (!narrow && bandIndex === fillAt ? extra : 0);
        if (narrow) {
          band.items.forEach((item, index) => {
            const titleLines = item.title ? Math.max(1, item.title.split("\n").length) : 0;
            const titleH = titleLines ? frame.inch(titleLines * Math.ceil(size.column * 1.25) + 10) : 0;
            const itemH = frame.inch(frame.listHeight([item]));
            const stepH = timed ? 0.2 : 0;
            const pad = timed ? 0.1 : 0;
            if (timed) {
              slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
                x: xStart, y: topCursor, w: box, h: itemH + stepH + pad * 2,
                fill: { color: WHITE }, rectRadius: CARD_RADIUS,
                line: { color: CARD_LINE, width: 0.75 },
              });
              slide.addText(String(index + 1).padStart(2, "0"), {
                x: xStart + pad, y: topCursor + pad, w: box - pad * 2, h: stepH,
                fontFace: "Calibri", fontSize: 11, bold: true, color: "C23FC2", margin: 0,
              });
            }
            if (item.title) {
              slide.addText(item.title, {
                x: xStart + pad, y: topCursor + pad + stepH, w: box - pad * 2, h: titleH,
                fontFace: "Calibri", fontSize: frame.pt(size.column), bold: true, color: INK, margin: 0, valign: "top",
              });
            }
            slide.addText(item.points.map((point, pointIndex) => ({
              text: point,
              options: { bullet: { code: "25CF", color: "E91E8C" }, breakLine: pointIndex < item.points.length - 1, color: "322C49" },
            })), {
              x: xStart + pad, y: topCursor + pad + titleH + stepH, w: box - pad * 2, h: Math.max(0.3, itemH - titleH),
              fontFace: "Calibri", fontSize: frame.pt(size.bullet), paraSpaceAfter: 4, margin: 0, valign: "top",
            });
            topCursor += itemH + stepH + pad * 2 + frame.inch(12);
          });
          continue;
        }
        band.items.forEach((item, index) => {
          const x = xStart + index * (columnW + gap);
          const titleLines = item.title ? Math.max(1, item.title.split("\n").length) : 0;
          const titleH = titleLines ? frame.inch(titleLines * Math.ceil(size.column * 1.25) + 10) : 0;
          const stepH = timed ? 0.2 : 0;
          const pad = timed ? 0.1 : 0;
          if (timed) {
            slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
              x, y: topCursor, w: columnW, h: height,
              fill: { color: WHITE }, rectRadius: CARD_RADIUS,
              line: { color: CARD_LINE, width: 0.75 },
            });
          }
          let top = topCursor + pad;
          if (timed) {
            slide.addText(String(index + 1).padStart(2, "0"), {
              x: x + pad, y: top, w: columnW - pad * 2, h: stepH,
              fontFace: "Calibri", fontSize: 11, bold: true, color: "C23FC2", margin: 0,
            });
            top += stepH;
          }
          if (item.title) {
            slide.addText(item.title, {
              x: x + pad, y: top, w: columnW - pad * 2, h: titleH,
              fontFace: "Calibri", fontSize: frame.pt(size.column), bold: true, color: INK, margin: 0, valign: "top",
            });
            top += titleH;
          }
          slide.addText(item.points.map((point, pointIndex) => ({
            text: point,
            options: { bullet: { code: "25CF", color: "E91E8C" }, breakLine: pointIndex < item.points.length - 1, color: "322C49" },
          })), {
            x: x + pad, y: top, w: columnW - pad * 2, h: Math.max(0.3, height - titleH - stepH - pad * 2),
            fontFace: "Calibri", fontSize: frame.pt(size.bullet), paraSpaceAfter: 4, margin: 0, valign: "top",
          });
        });
        topCursor += height + frame.inch(16);
        continue;
      }
      if (band.kind === "cards" && composition === "cascade" && band.tone === "stat" && !narrow && band.items[0]) {
        const items = band.items;
        const displays = items.map((item) => item.display);
        const ceiling = air ? (items.length > 2 ? 96 : 128) : size.stat;
        const sizes = items.map((item) => figurePx(item.display, displays, ceiling));
        const maxPx = Math.max(...sizes);
        const labelH = frame.inch(Math.ceil(size.statLabel * 1.3));
        const detailH = frame.inch(Math.ceil(size.statBody * 1.4) + 8);
        const tallestDetail = Math.max(...items.map((item) => (item.detail.trim() ? detailH : 0)));
        const bar = sameFigure(displays);
        const rowH = frame.inch(maxPx) + labelH + tallestDetail + (bar ? 0.28 : 0);
        const gap = 0.28;
        const colW = (box - gap * (items.length - 1)) / items.length;
        items.forEach((item, index) => {
          const px = sizes[index] ?? ceiling;
          const x = xStart + index * (colW + gap);
          slide.addText(item.display, {
            x, y: topCursor + frame.inch(maxPx - px), w: colW, h: frame.inch(px),
            fontFace: "Calibri", fontSize: frame.pt(px), bold: true, color: INK, margin: 0, valign: "bottom",
          });
          let top = topCursor + frame.inch(maxPx);
          slide.addText(item.label || item.title, {
            x, y: top, w: colW, h: labelH,
            fontFace: "Calibri", fontSize: frame.pt(size.statLabel), bold: true, color: INK, margin: 0, valign: "top",
          });
          top += labelH;
          if (item.detail.trim()) {
            slide.addText(item.detail, {
              x, y: top + 0.04, w: colW, h: detailH,
              fontFace: "Calibri", fontSize: frame.pt(size.statBody), color: "322C49", margin: 0, valign: "top",
            });
          }
          if (bar) {
            drawTrack(pres, slide, x, topCursor + rowH - 0.1, colW, 0.08, "3200BE", figureShare(item.display, displays));
          }
        });
        topCursor += rowH + frame.inch(16);
        continue;
      }
      const cards = band.items;
      const displays = band.tone === "stat" ? cards.map((item) => item.display) : [];
      const bareShare = sameFigure(displays) && cards.every((item) => item.display && !item.detail.trim());
      const across = narrow ? 1 : gridAcross(band.tone, cards.length);
      const gap = frame.inch(14);
      const bonus = bandIndex === fillAt ? extra / Math.max(1, Math.ceil(cards.length / across)) : 0;
      for (let row = 0; row < cards.length; row += across) {
        const slice = cards.slice(row, row + across);
        const cardW = (box - gap * (slice.length - 1)) / slice.length;
        const heights = slice.map((item) => frame.cardHeight(item, narrow ? 2 : slice.length));
        const naturalH = frame.inch(Math.max(...heights));
        const rowH = naturalH + bonus;
        const plainStats = composition === "metrics" && band.tone === "stat" && slice.every((item) => !cardFace(item.title));
        if (plainStats) {
          slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: xStart, y: topCursor, w: box, h: rowH,
            fill: { color: STAGE },
            rectRadius: 0.22,
            line: { color: STAGE, width: 0 },
          });
          slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: xStart + 0.18, y: topCursor + 0.1, w: Math.max(0.4, box - 0.36), h: 0.045,
            fill: { color: "C23FC2" },
            rectRadius: 0.02,
            line: { color: "C23FC2", width: 0 },
          });
        }
        slice.forEach((item, index) => {
          const x = xStart + index * (cardW + gap);
          const face = cardFace(item.title);
          const faceFill = face === "green" ? "E7F6EF" : face === "amber" ? "FFF4D6" : face === "red" ? "F8E8EE" : face === "ink" ? STAGE : WHITE;
          const faceInk = face === "green" ? "0E6B45" : face === "amber" ? "8A5A00" : face === "red" ? "8E2348" : face === "ink" ? "F4F1FA" : INK;
          const faceBody = face === "ink" ? "D9D3E8" : face ? faceInk : "322C49";
          const openStat = plainStats || (composition === "metrics" && band.tone === "stat" && !face);
          if (!openStat) {
            slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
              x, y: topCursor, w: cardW, h: rowH,
              fill: { color: faceFill },
              rectRadius: CARD_RADIUS,
              line: face ? { color: faceFill, width: 0 } : { color: CARD_LINE, width: 0.75 },
            });
          }
          const inset = 0.16;
          const shift = bonus > 0.08 && !openStat ? bonus / 2 : 0;
          const boxBottom = topCursor + naturalH + shift;
          let top = topCursor + 0.12 + shift;
          if (item.display) {
            const displayH = frame.inch(size.stat);
            slide.addText(item.display, {
              x: x + inset, y: top, w: cardW - inset * 2, h: displayH,
              fontFace: "Calibri", fontSize: frame.pt(size.stat), bold: true, color: openStat ? WHITE : faceInk, margin: 0, valign: "top",
            });
            top += displayH;
            const labelH = frame.inch(Math.ceil(size.statLabel * 1.3));
            slide.addText(item.label || item.title, {
              x: x + inset, y: top, w: cardW - inset * 2, h: labelH,
              fontFace: "Calibri", fontSize: frame.pt(size.statLabel), bold: true, color: openStat ? WHITE : faceInk, margin: 0, valign: "top",
            });
            top += labelH + 0.04;
            if (item.detail) {
              slide.addText(item.detail, {
                x: x + inset, y: top, w: cardW - inset * 2, h: Math.max(0.24, boxBottom - top - 0.1),
                fontFace: "Calibri", fontSize: frame.pt(size.statBody), color: openStat ? "D9D3E8" : faceBody, margin: 0, valign: "top",
              });
            }
            if (bareShare) {
              drawTrack(pres, slide, x + inset, topCursor + rowH - 0.14, cardW - inset * 2, 0.045, openStat ? "C23FC2" : "3200BE", figureShare(item.display, displays));
            }
            return;
          }
          const titleFont = band.tone === "chips" ? size.chip : size.cardTitle;
          const titleLines = Math.max(1, item.title.split("\n").length);
          const titleH = frame.inch(titleLines * Math.ceil(titleFont * 1.3));
          slide.addText(item.title, {
            x: x + inset, y: top, w: cardW - inset * 2, h: titleH,
            fontFace: "Calibri", fontSize: frame.pt(titleFont), bold: true, color: faceInk, margin: 0, valign: "top",
          });
          if (item.detail) {
            slide.addText(item.detail, {
              x: x + inset, y: top + titleH + 0.04, w: cardW - inset * 2, h: Math.max(0.24, boxBottom - (top + titleH) - 0.12),
              fontFace: "Calibri", fontSize: frame.pt(size.cardBody), color: faceBody, margin: 0, valign: "top",
            });
          }
        });
        topCursor += rowH + gap;
      }
      topCursor += frame.inch(2);
    }
    return topCursor;
  };
  const proof = composition === "proof" ? proofSplit(pieces) : null;
  const compare = composition === "versus" ? compareSplit(pieces) : null;
  const aside = composition === "hero" ? heroAside(pieces) : null;
  const poster = composition === "poster" && pieces[0]?.kind === "prose" ? pieces[0].text : "";
  if (poster) {
    const px = posterPx(poster);
    const chars = Math.max(8, Math.floor(1168 / (px * 0.5)));
    const lines = poster.split("\n").reduce((total, line) => total + Math.max(1, Math.ceil(line.trim().length / chars)), 0);
    const textH = Math.max(0.5, (lines * Math.ceil(px * 1.05)) / 96);
    const y = Math.max(cursor, floor - textH);
    slide.addText(poster, {
      x: x0, y, w: width, h: Math.min(textH, floor - y),
      fontFace: "Calibri", fontSize: frame.pt(px), bold: true, color: INK, margin: 0, valign: "bottom",
    });
    cursor = floor;
  } else if (proof) {
    const gap = 0.28;
    const leftW = width * 0.6;
    const rightW = width - leftW - gap;
    const regionH = Math.max(1.6, floor - cursor);
    const px = proofPx(proof.text);
    const chars = Math.max(8, Math.floor((leftW * 96) / (px * 0.5)));
    const lines = Math.max(1, Math.ceil(proof.text.trim().length / chars));
    const textH = Math.min(regionH, (lines * Math.ceil(px * 1.08)) / 96 + 0.06);
    slide.addText(proof.text, {
      x: x0, y: cursor + regionH - textH, w: leftW, h: textH,
      fontFace: "Calibri", fontSize: frame.pt(px), bold: true, color: INK, margin: 0, valign: "bottom",
    });
    const card = proof.card;
    const face = cardFace(card.title);
    const fill = face === "green" ? "E7F6EF" : face === "amber" ? "FFF4D6" : face === "red" ? "F8E8EE" : face === "ink" ? STAGE : WHITE;
    const ink = face === "green" ? "0E6B45" : face === "amber" ? "8A5A00" : face === "red" ? "8E2348" : face === "ink" ? "F4F1FA" : INK;
    const body = face === "ink" ? "D9D3E8" : face ? ink : "322C49";
    const x = x0 + leftW + gap;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: cursor, w: rightW, h: regionH,
      fill: { color: fill },
      rectRadius: CARD_RADIUS,
      line: face ? { color: fill, width: 0 } : { color: CARD_LINE, width: 0.75 },
    });
    const displayH = card.display ? frame.inch(size.stat) : 0;
    const titleText = card.display ? card.label || card.title : card.title;
    const titleFont = card.display ? size.statLabel : size.cardTitle;
    const titleH = frame.inch(Math.max(1, titleText.split("\n").length) * Math.ceil(titleFont * 1.25));
    const detailH = card.detail ? frame.inch(Math.max(1, card.detail.split("\n").length) * Math.ceil(size.cardBody * 1.4) + 8) : 0;
    let top = cursor + regionH - (displayH + titleH + detailH + 0.2) - 0.16;
    if (card.display) {
      slide.addText(card.display, {
        x: x + 0.18, y: top, w: rightW - 0.36, h: displayH,
        fontFace: "Calibri", fontSize: frame.pt(size.stat), bold: true, color: ink, margin: 0, valign: "top",
      });
      top += displayH;
    }
    slide.addText(titleText, {
      x: x + 0.18, y: top, w: rightW - 0.36, h: titleH,
      fontFace: "Calibri", fontSize: frame.pt(titleFont), bold: true, color: ink, margin: 0, valign: "top",
    });
    if (card.detail) {
      slide.addText(card.detail, {
        x: x + 0.18, y: top + titleH + 0.06, w: rightW - 0.36, h: detailH,
        fontFace: "Calibri", fontSize: frame.pt(size.cardBody), color: body, margin: 0, valign: "top",
      });
    }
    cursor += regionH;
  } else if (compare) {
    if (compare.before.length) cursor = paint(compare.before, x0, cursor, width, false);
    const gap = 0.22;
    const colW = (width - gap) / 2;
    const leftNatural = pieceBands(compare.left).reduce((sum, band) => sum + bandInches(band, true), 0);
    const rightNatural = pieceBands(compare.right).reduce((sum, band) => sum + bandInches(band, true), 0);
    const splitH = Math.max(leftNatural, rightNatural, 1);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x0, y: cursor, w: colW, h: splitH,
      fill: { color: "F4F1FB" },
      rectRadius: 0.18,
      line: { color: "F4F1FB", width: 0 },
    });
    const left = paint(compare.left, x0, cursor, colW, true);
    const right = paint(compare.right, x0 + colW + gap, cursor, colW, true);
    const spineH = Math.max(splitH, left - cursor, right - cursor);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x0 + colW + gap / 2 - 0.01,
      y: cursor + 0.08,
      w: 0.02,
      h: Math.max(0.3, spineH - 0.16),
      fill: { color: "C23FC2" },
      rectRadius: 0.02,
      line: { color: "C23FC2", width: 0 },
    });
    cursor = Math.max(left, right, cursor + splitH);
  } else if (aside) {
    const gap = 0.2;
    const statW = width * 0.34;
    const restEnd = paint(aside.rest, x0 + statW + gap, cursor, width - statW - gap, false);
    const statH = Math.max(1.6, restEnd - cursor, source.images.length ? 0 : floor - cursor);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x0, y: cursor, w: statW, h: statH,
      fill: { color: STAGE },
      rectRadius: CARD_RADIUS,
      line: { color: STAGE, width: 0 },
    });
    const heroPx = displayPx(aside.stat.display, air ? 120 : size.stat);
    const displayH = frame.inch(heroPx);
    const labelH = 0.55;
    const detailH = aside.stat.detail ? 0.7 : 0;
    const block = displayH + 0.08 + labelH + (detailH ? 0.18 + detailH : 0);
    let top = cursor + Math.max(0.2, (statH - block) / 2);
    slide.addText(aside.stat.display, {
      x: x0 + 0.18, y: top, w: statW - 0.36, h: displayH,
      fontFace: "Calibri", fontSize: frame.pt(heroPx), bold: true, color: WHITE, margin: 0, valign: "top",
    });
    top += displayH + 0.06;
    slide.addText(aside.stat.label || aside.stat.title, {
      x: x0 + 0.18, y: top, w: statW - 0.36, h: labelH,
      fontFace: "Calibri", fontSize: frame.pt(size.statLabel), bold: true, color: WHITE, margin: 0, valign: "top",
    });
    if (aside.stat.detail) {
      slide.addText(aside.stat.detail, {
        x: x0 + 0.18, y: top + labelH + 0.16, w: statW - 0.36, h: detailH,
        fontFace: "Calibri", fontSize: frame.pt(size.statBody), color: "D9D3E8", margin: 0, valign: "top",
      });
    }
    cursor = Math.max(restEnd, cursor + statH);
  } else {
    const grow = !air ? "none" : composition === "path" ? "lists" : composition === "stack" || composition === "metrics" || composition === "cascade" ? "cards" : "none";
    cursor = paint(pieces, x0, cursor, width, false, grow);
  }
  if (source.images.length && cursor < 6.3) placeKept(slide, source.images, x0, cursor, width, Math.max(0.8, 7.05 - cursor));
}

function addOriginCover(pres: Pres, book: Book, source: SourceSlide, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  speak(slide, source);
  slide.background = { color: COVER };
  paintCoverWave(pres, slide);
  if (book.logo) placeMark(pres, slide, book.logo);
  const title = source.title.trim() || book.sourceDeck?.title || "Untitled";
  const wide = title.length > 42;
  const titleSize = wide ? Math.max(28, type.title - 4) : title.length > 22 ? type.title + 4 : type.title + 14;
  const titleTop = 1.25;
  const widthIn = book.logo ? 9.2 : 11.8;
  const charsPerLine = Math.max(16, Math.floor((widthIn * 72) / (titleSize * 0.45)));
  const lines = Math.min(4, title.split("\n").reduce((total, line) => total + Math.max(1, Math.ceil(line.length / charsPerLine)), 0));
  const titleH = 0.12 + lines * (titleSize / 72) * 1.2;
  const titleLines = title.split("\n");
  const accent = titleLines.length > 1 ? titleLines[titleLines.length - 1] ?? "" : "";
  const lead = accent ? titleLines.slice(0, -1).join("\n") : title;
  slide.addText(accent
    ? [
        { text: `${lead}\n`, options: { color: WHITE } },
        { text: accent, options: { color: "FF97FE" } },
      ]
    : title, {
    x: 0.75,
    y: titleTop,
    w: widthIn,
    h: titleH,
    fontFace: "Calibri",
    fontSize: titleSize,
    bold: true,
    color: WHITE,
    margin: 0,
    valign: "top",
  });
  const presenter = source.context.trim();
  if (presenter) {
    slide.addText(presenter, {
      x: 0.75,
      y: titleTop + titleH + 0.14,
      w: 11,
      h: 0.42,
      fontFace: "Calibri",
      fontSize: 22,
      color: "E4E0EE",
      margin: 0,
    });
  }
}

function addOriginSheet(pres: Pres, source: SourceSlide, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  speak(slide, source);
  slide.background = { color: WHITE };
  const bullets = bulletLines(source);
  const continued = isContinued(source);
  const lead = leadCopy(source);
  const spokenOnly = Boolean(source.spoken?.trim()) && !lead && bullets.length === 0 && source.images.length === 0;
  const images = source.images;
  let cursor = 0.46;
  if (continued) {
    kicker(slide, "Continued", 0.7, cursor, "C23FC2", type.kicker);
    cursor = 0.82;
  }
  if (source.eyebrow) {
    slide.addText(source.eyebrow.toUpperCase(), {
      x: 0.7, y: cursor, w: 12, h: 0.3,
      fontFace: "Calibri", fontSize: 12, bold: true, color: "C23FC2", margin: 0, charSpacing: 1.1,
    });
    cursor += 0.36;
  }
  if (source.title.trim()) {
    const titleLines = Math.max(1, source.title.trim().split("\n").length);
    const titleH = spokenOnly ? 1.5 : titleLines > 1 ? 0.48 * titleLines : images.length ? 0.52 : 0.72;
    slide.addText(source.title.trim(), {
      x: 0.7,
      y: spokenOnly ? 2.15 : cursor,
      w: 12,
      h: titleH,
      fontFace: "Calibri",
      fontSize: spokenOnly ? 40 : titleLines > 1 ? Math.max(26, type.title - 8) : images.length ? Math.max(24, type.title - 12) : type.title - 2,
      bold: true,
      color: INK,
      margin: 0,
      valign: "top",
    });
    cursor += titleH + 0.12;
  }
  if (lead) {
    const height = images.length ? Math.min(1.15, 0.36 + lead.length / 180) : Math.min(4.8, 0.46 * Math.max(1, Math.ceil(lead.length / 78)));
    slide.addText(lead, {
      x: 0.7,
      y: cursor,
      w: images.length ? 12 : 10.4,
      h: height,
      fontFace: "Calibri",
      fontSize: images.length ? 16 : lead.length > 240 ? 18 : 22,
      color: "322C49",
      margin: 0,
      valign: "top",
    });
    cursor += height + 0.1;
  }
  if (bullets.length) {
    const size = bullets.length > 8 || images.length ? 15 : 20;
    const height = images.length
      ? Math.min(2.3, 0.34 * bullets.length + 0.08)
      : Math.min(5.5, 0.4 * bullets.length + 0.16);
    slide.addText(bullets.map((text, index) => ({
      text,
      options: { bullet: { code: "25CF", color: "E91E8C" }, breakLine: index < bullets.length - 1, color: INK },
    })), {
      x: 0.7,
      y: cursor,
      w: 12,
      h: height,
      fontFace: "Calibri",
      fontSize: size,
      paraSpaceAfter: bullets.length > 8 ? 4 : 8,
      valign: "top",
      margin: 0,
    });
    cursor += height + 0.12;
  }
  if (images.length && cursor < 5.4) {
    const top = cursor + 0.06;
    placeKept(slide, images, 0.7, top, 12, 7.12 - top);
  }
}

function addOriginTable(pres: Pres, source: SourceSlide, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  speak(slide, source);
  slide.background = { color: WHITE };
  const table = tableFromSlide(source);
  const pictures = source.images.length > 0;
  let cursor = 0.42;
  if (isContinued(source)) {
    kicker(slide, "Continued", 0.55, cursor, "C23FC2", type.kicker);
    cursor = 0.78;
  }
  if (source.eyebrow) {
    slide.addText(source.eyebrow.toUpperCase(), {
      x: 0.55, y: cursor, w: 12.2, h: 0.28,
      fontFace: "Calibri", fontSize: 12, bold: true, color: "C23FC2", margin: 0, charSpacing: 1.1,
    });
    cursor += 0.32;
  }
  const title = source.title.trim() || "Table";
  const titleLines = Math.max(1, Math.ceil(title.length / 46));
  const titleH = 0.38 * titleLines;
  slide.addText(title, {
    x: 0.55,
    y: cursor,
    w: 12.2,
    h: titleH,
    fontFace: "Calibri",
    fontSize: titleLines > 1 ? type.title - 12 : type.title - 8,
    bold: true,
    color: INK,
    margin: 0,
    valign: "top",
  });
  cursor += titleH + 0.12;
  const pieces = source.pieces ?? [];
  const lead = pieces.length
    ? pieces.filter((piece) => piece.kind === "prose").map((piece) => piece.text).join("\n")
    : source.context;
  const after = pieces.filter((piece) => piece.kind !== "prose");
  const reserve = after.length ? 1.15 : 0;
  if (lead.trim()) cursor = placeCopy(slide, lead, 0.55, cursor, 12.2, pictures ? 2.4 : 3.15);
  const floor = (pictures ? 5.35 : 6.9) - reserve;
  if (table) addSourceTable(slide, table, 0.55, cursor, 12.2, Math.max(1.2, floor - cursor));
  cursor = floor + 0.08;
  after.forEach((piece) => {
    if (piece.kind === "card") {
      slide.addText(piece.title, {
        x: 0.55, y: cursor, w: 12.2, h: 0.28,
        fontFace: "Calibri", fontSize: 14, bold: true, color: INK, margin: 0,
      });
      cursor += 0.28;
      if (piece.detail) {
        slide.addText(piece.detail, {
          x: 0.55, y: cursor, w: 12.2, h: 0.42,
          fontFace: "Calibri", fontSize: 13, color: "322C49", margin: 0,
        });
        cursor += 0.46;
      }
    }
  });
  if (pictures) placeKept(slide, source.images, 0.55, 5.5, 12.2, 1.55);
}

function addSourceTable(slide: Slide, table: SourceTable, x: number, y: number, w: number, h: number) {
  const headers = table.headers.length ? table.headers : ["Item"];
  const textual = !tableHasNumbers(table);
  const header = headers.map((text) => ({
    text: text || " ",
    options: { fill: { color: INK }, color: WHITE, bold: true, align: "left" as const },
  }));
  const body = table.rows.map((row, index) => headers.map((_, cell) => ({
    text: row[cell] || " ",
    options: {
      fill: { color: index % 2 ? "F7F5FB" : WHITE },
      color: INK,
      align: (!textual && cell === headers.length - 1 ? "right" : "left") as "left" | "right",
      bold: !textual && cell === headers.length - 1,
    },
  })));
  const lead = Math.min(w * 0.62, w - 0.7 * Math.max(1, headers.length - 1));
  const rest = (w - lead) / Math.max(1, headers.length - 1);
  const colW = textual ? headers.map(() => w / headers.length) : headers.map((_, index) => (index === 0 ? lead : rest));
  slide.addTable([header, ...body], {
    x,
    y,
    w,
    h,
    colW,
    border: { pt: 0.25, color: LINE },
    fontFace: "Calibri",
    fontSize: table.rows.length > 8 ? 10 : 12,
    valign: "middle",
  });
}

function addFunnelDrawing(pres: Pres, slide: Slide, chart: SourceChart, table: SourceTable | null, x: number, y: number, w: number, h: number) {
  const series = chart.series[0];
  const values = series?.values ?? [];
  const max = Math.max(1, ...values.map((value) => Math.max(0, value)));
  const rows = Math.max(1, chart.categories.length);
  const gap = 0.08;
  const rowH = Math.min(0.62, (h - gap * (rows - 1)) / rows);
  chart.categories.forEach((category, index) => {
    const value = values[index] ?? 0;
    const span = Math.max(w * 0.38, w * (Math.max(0, value) / max));
    const top = y + index * (rowH + gap);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + (w - span) / 2,
      y: top,
      w: span,
      h: rowH,
      fill: { color: index % 2 ? "5B3BE0" : BRAND },
      rectRadius: 0.08,
    });
    const written = rowDisplay(table, category, series ? [series] : [], formatPoint(value));
    slide.addText(`${category}   ${written}`, {
      x: x + (w - span) / 2 + 0.12,
      y: top,
      w: span - 0.24,
      h: rowH,
      fontFace: "Calibri",
      fontSize: 12,
      bold: true,
      color: WHITE,
      margin: 0,
      valign: "middle",
    });
  });
}

function drawTrack(pres: Pres, slide: Slide, x: number, y: number, w: number, h: number, fill: string, span: number) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: TRACK },
    rectRadius: h / 2,
  });
  const bar = Math.max(0, Math.min(w, w * span));
  if (bar < 0.04) return;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w: bar, h,
    fill: { color: fill },
    rectRadius: Math.min(h / 2, bar / 2),
  });
}

function drawScore(pres: Pres, slide: Slide, chart: SourceChart, table: SourceTable | null, x: number, y: number, w: number, h: number) {
  const series = chart.series[0];
  if (!series) return;
  const cats = chart.categories.slice(0, 8);
  const values = cats.map((_, index) => series.values[index] ?? 0);
  const max = Math.max(1, ...values.map((value) => Math.max(0, value)));
  const named = cats.findIndex((category) => /^total\b/i.test(category.trim()));
  const lead = named >= 0 ? named : values.reduce((best, value, index) => (value > (values[best] ?? 0) ? index : best), 0);
  const rest = cats.map((_, index) => index).filter((index) => index !== lead);
  const hero = rowDisplay(table, cats[lead] ?? "", [series], formatPoint(values[lead] ?? 0));
  const wide = rest.length > 3;
  const heroPx = displayPx(hero, wide ? 84 : 128);
  const tones = ["3200BE", "9A2F9A", "7C8CFF", "5C6BC0", "1F9E72", "C23FC2", "5B3BE0"];
  if (!rest.length) {
    slide.addText((cats[lead] ?? "").toUpperCase(), {
      x, y: y + 0.4, w, h: 0.28,
      fontFace: "Calibri", fontSize: 12, bold: true, color: "C23FC2", margin: 0, charSpacing: 1.1,
    });
    slide.addText(hero, {
      x, y: y + 0.7, w, h: heroPx / 96,
      fontFace: "Calibri", fontSize: heroPx * 0.75, bold: true, color: INK, margin: 0,
    });
    drawTrack(pres, slide, x, y + 0.78 + heroPx / 96, w, 0.16, "3200BE", Math.max(0, values[lead] ?? 0) / max);
    return;
  }
  if (wide) {
    const heroH = heroPx / 96;
    slide.addText((cats[lead] ?? "").toUpperCase(), {
      x, y, w, h: 0.26,
      fontFace: "Calibri", fontSize: 12, bold: true, color: "C23FC2", margin: 0, charSpacing: 1.1,
    });
    slide.addText(hero, {
      x, y: y + 0.26, w, h: heroH,
      fontFace: "Calibri", fontSize: heroPx * 0.75, bold: true, color: INK, margin: 0, valign: "top",
    });
    const beamY = y + 0.32 + heroH;
    drawTrack(pres, slide, x, beamY, w * 0.72, 0.14, "3200BE", Math.max(0, values[lead] ?? 0) / max);
    const gap = 0.12;
    const colW = (w - gap * (rest.length - 1)) / rest.length;
    const base = y + h - 0.58;
    const tower = Math.max(0.8, base - (beamY + 0.55));
    rest.forEach((index, slot) => {
      const left = x + slot * (colW + gap);
      const value = values[index] ?? 0;
      const ratio = value <= 0 ? 0 : Math.max(0, value) / max;
      const bh = ratio <= 0 ? 0 : Math.max(0.08, tower * ratio);
      const written = rowDisplay(table, cats[index] ?? "", [series], formatPoint(value));
      slide.addText(written, {
        x: left, y: base - bh - 0.32, w: colW, h: 0.28,
        fontFace: "Calibri", fontSize: displayPx(written, 22) * 0.75, bold: true, color: INK, margin: 0,
      });
      if (bh > 0) {
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
          x: left, y: base - bh, w: colW, h: bh,
          fill: { color: tones[slot] ?? "3200BE" },
          rectRadius: 0.06,
        });
      }
      slide.addText(cats[index] ?? "", {
        x: left, y: base + 0.04, w: colW, h: 0.5,
        fontFace: "Calibri", fontSize: 11, color: "322C49", margin: 0, valign: "top",
      });
    });
    return;
  }
  const leftW = w * 0.54;
  const base = y + Math.max(2.4, h - 0.72);
  slide.addText((cats[lead] ?? "").toUpperCase(), {
    x, y: base - 0.34 - heroPx / 96, w: leftW, h: 0.26,
    fontFace: "Calibri", fontSize: 12, bold: true, color: "C23FC2", margin: 0, charSpacing: 1.1,
  });
  slide.addText(hero, {
    x, y: base - 0.08 - heroPx / 96, w: leftW, h: heroPx / 96,
    fontFace: "Calibri", fontSize: heroPx * 0.75, bold: true, color: INK, margin: 0, valign: "bottom",
  });
  drawTrack(pres, slide, x, base - 0.02, leftW, 0.16, "3200BE", Math.max(0, values[lead] ?? 0) / max);
  const gap = 0.16;
  const rightX = x + leftW + 0.28;
  const rightW = Math.max(1.2, w - leftW - 0.28);
  const colW = Math.min(1.7, (rightW - gap * (rest.length - 1)) / rest.length);
  const tower = Math.max(1.5, base - y - 0.55);
  rest.forEach((index, slot) => {
    const left = rightX + slot * (colW + gap);
    const value = values[index] ?? 0;
    const ratio = value <= 0 ? 0 : Math.max(0, value) / max;
    const bh = ratio <= 0 ? 0 : Math.max(0.08, tower * ratio);
    const written = rowDisplay(table, cats[index] ?? "", [series], formatPoint(value));
    slide.addText(written, {
      x: left, y: base - bh - 0.4, w: colW, h: 0.36,
      fontFace: "Calibri", fontSize: displayPx(written, 28) * 0.75, bold: true, color: INK, margin: 0,
    });
    if (bh > 0) {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: left, y: base - bh, w: colW, h: bh,
        fill: { color: tones[slot] ?? "3200BE" },
        rectRadius: 0.08,
      });
    }
    slide.addText(cats[index] ?? "", {
      x: left, y: base + 0.06, w: colW, h: 0.62,
      fontFace: "Calibri", fontSize: 11, color: "322C49", margin: 0, valign: "top",
    });
  });
}

function drawRank(pres: Pres, slide: Slide, chart: SourceChart, table: SourceTable | null, x: number, y: number, w: number, h: number) {
  const series = chart.series[0];
  if (!series) return;
  const cats = chart.categories.slice(0, 14);
  const values = cats.map((_, index) => series.values[index] ?? 0);
  const max = Math.max(1, ...values.map((value) => Math.max(0, value)));
  const named = cats.findIndex((category) => /^total\b/i.test(category.trim()));
  const lead = named >= 0 ? named : values.reduce((best, value, index) => (value > (values[best] ?? 0) ? index : best), 0);
  const rest = cats.map((_, index) => index).filter((index) => index !== lead);
  const beaconH = 1.15;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h: beaconH,
    fill: { color: STAGE },
    rectRadius: 0.18,
    line: { color: STAGE, width: 0 },
  });
  slide.addText(cats[lead] ?? "", {
    x: x + 0.28, y: y + 0.22, w: w - 3.5, h: 0.72,
    fontFace: "Calibri", fontSize: 16, bold: true, color: "FF96FF", margin: 0, valign: "middle",
  });
  const hero = rowDisplay(table, cats[lead] ?? "", [series], formatPoint(values[lead] ?? 0));
  const heroPt = displayPx(hero, 72) * 0.75;
  slide.addText(hero, {
    x: x + w - 3.15, y: y + 0.14, w: 2.8, h: 0.88,
    fontFace: "Calibri", fontSize: heroPt, bold: true, color: WHITE, align: "right", margin: 0, valign: "middle",
  });
  const rowH = Math.min(0.32, (h - beaconH - 0.14) / Math.max(1, rest.length));
  rest.forEach((index, slot) => {
    const top = y + beaconH + 0.14 + slot * rowH;
    const cat = cats[index] ?? "";
    slide.addText(cat, {
      x, y: top, w: 3.15, h: rowH,
      fontFace: "Calibri", fontSize: rest.length > 10 ? 11 : 13, color: INK, margin: 0, valign: "middle",
    });
    const value = values[index] ?? 0;
    drawTrack(pres, slide, x + 3.25, top + (rowH - 0.08) / 2, w - 4.55, 0.08, "3200BE", value <= 0 ? 0 : Math.max(0, value) / max);
    slide.addText(rowDisplay(table, cat, [series], formatPoint(values[index] ?? 0)), {
      x: x + w - 1.15, y: top, w: 1.15, h: rowH,
      fontFace: "Calibri", fontSize: 13, bold: true, color: INK, align: "right", margin: 0, valign: "middle",
    });
  });
}

function drawMatrix(pres: Pres, slide: Slide, chart: SourceChart, table: SourceTable | null, x: number, y: number, w: number, h: number) {
  const cats = chart.categories.slice(0, 12);
  const series = chart.series.slice(0, 4);
  const spare = spareHeaders(table, series);
  const labelW = 2.45;
  const cols = Math.max(1, series.length + spare.length);
  const colW = (w - labelW) / cols;
  slide.addText("Each column has its own scale. The figure is the cell as written.", {
    x, y, w, h: 0.26,
    fontFace: "Calibri", fontSize: 11, color: MUTED, margin: 0,
  });
  const head = y + 0.3;
  series.forEach((item, index) => {
    slide.addText(item.name.toUpperCase(), {
      x: x + labelW + index * colW, y: head, w: colW - 0.08, h: 0.22,
      fontFace: "Calibri", fontSize: 10, bold: true, color: MUTED, align: "right", margin: 0,
    });
  });
  spare.forEach((header, index) => {
    slide.addText(header.toUpperCase(), {
      x: x + labelW + (series.length + index) * colW, y: head, w: colW - 0.08, h: 0.22,
      fontFace: "Calibri", fontSize: 10, bold: true, color: MUTED, align: "right", margin: 0,
    });
  });
  const maxes = series.map((item) => Math.max(1, ...cats.map((_, index) => Math.max(0, item.values[index] ?? 0))));
  const rowH = Math.min(0.4, (h - 0.62) / Math.max(1, cats.length));
  cats.forEach((cat, index) => {
    const top = head + 0.26 + index * rowH;
    const row = table?.rows.find((item) => (item[0] ?? "").trim() === cat);
    slide.addText(cat, {
      x, y: top, w: labelW - 0.12, h: rowH,
      fontFace: "Calibri", fontSize: 12, color: INK, margin: 0, valign: "middle",
    });
    series.forEach((item, seriesIndex) => {
      const left = x + labelW + seriesIndex * colW;
      const value = item.values[index] ?? 0;
      const peakAt = cats.reduce((best, _, valueIndex) => ((item.values[valueIndex] ?? 0) > (item.values[best] ?? 0) ? valueIndex : best), 0);
      const peak = index === peakAt && value > 0;
      const numW = Math.min(0.85, colW * 0.34);
      drawTrack(pres, slide, left, top + (rowH - 0.08) / 2, Math.max(0.3, colW - numW - 0.16), 0.08, CHART_COLORS[seriesIndex] ?? "3200BE", Math.max(0, value) / (maxes[seriesIndex] ?? 1));
      slide.addText(rowDisplay(table, cat, [item], formatPoint(value)), {
        x: left + colW - numW - 0.08, y: top, w: numW, h: rowH,
        fontFace: "Calibri", fontSize: peak ? 14 : 12, bold: true, color: peak ? BRAND : INK, align: "right", margin: 0, valign: "middle",
      });
    });
    spare.forEach((header, spareIndex) => {
      const column = table?.headers.findIndex((item) => item.trim() === header) ?? -1;
      slide.addText(column >= 0 ? row?.[column] ?? "" : "", {
        x: x + labelW + (series.length + spareIndex) * colW, y: top, w: colW - 0.08, h: rowH,
        fontFace: "Calibri", fontSize: 12, bold: true, color: INK, align: "right", margin: 0, valign: "middle",
      });
    });
  });
}

function drawSegment(pres: Pres, slide: Slide, x1: number, y1: number, x2: number, y2: number, color: string, width = 2.25, transparency = 0) {
  slide.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    w: Math.max(0.01, Math.abs(x2 - x1)),
    h: Math.max(0.01, Math.abs(y2 - y1)),
    flipV: y2 < y1,
    line: { color, width, transparency },
  });
}

function drawTrend(pres: Pres, slide: Slide, chart: SourceChart, x: number, y: number, w: number, h: number) {
  const cats = chart.categories.slice(0, 16);
  const series = chart.series.slice(0, 4);
  const span = axisSpan(series.flatMap((item) => cats.map((_, index) => Math.max(0, item.values[index] ?? 0))));
  const left = x + 0.7;
  const plotW = Math.max(1, w - 0.85);
  const plotTop = y + 0.28;
  const plotH = Math.max(1.2, h - 0.72);
  const range = span.max - span.min || 1;
  const xOf = (index: number) => (cats.length <= 1 ? left + plotW / 2 : left + (plotW * index) / (cats.length - 1));
  const yOf = (value: number) => plotTop + plotH - ((Math.max(span.min, value) - span.min) / range) * plotH;
  const ticks = [span.max, span.min + range / 2, span.min];
  ticks.forEach((tick) => {
    const yy = yOf(tick);
    slide.addShape(pres.shapes.LINE, { x: left, y: yy, w: plotW, h: 0, line: { color: "E4E1EC", width: 1 } });
    slide.addText(formatPoint(tick), {
      x: x, y: yy - 0.12, w: 0.62, h: 0.24,
      fontFace: "Calibri", fontSize: 10, color: MUTED, align: "right", margin: 0,
    });
  });
  series.forEach((item, seriesIndex) => {
    const color = CHART_COLORS[seriesIndex] ?? "3200BE";
    const points = cats.map((_, index) => ({ x: xOf(index), y: yOf(item.values[index] ?? 0) }));
    for (let index = 0; index < points.length - 1; index += 1) {
      const from = points[index];
      const to = points[index + 1];
      if (from && to) {
        drawSegment(pres, slide, from.x, from.y, to.x, to.y, color, 9, 82);
        drawSegment(pres, slide, from.x, from.y, to.x, to.y, color, 2.25, 0);
      }
    }
    const peakAt = cats.reduce((best, _, index) => ((item.values[index] ?? 0) > (item.values[best] ?? -Infinity) ? index : best), 0);
    points.forEach((point, index) => {
      const peak = index === peakAt;
      const dot = peak ? 0.11 : 0.07;
      slide.addShape(pres.shapes.OVAL, {
        x: point.x - dot, y: point.y - dot, w: dot * 2, h: dot * 2,
        fill: { color },
        line: { color: "FFFFFF", width: 1.25 },
      });
    });
    if (series.length !== 1 || cats.length > 8) {
      const peak = points[peakAt];
      if (peak) {
        slide.addText(formatPoint(item.values[peakAt] ?? 0), {
          x: peak.x - 0.4, y: peak.y - 0.34, w: 0.8, h: 0.22,
          fontFace: "Calibri", fontSize: 12, bold: true, color: INK, align: "center", margin: 0,
        });
      }
    }
    if (series.length === 1 && cats.length <= 8) {
      points.forEach((point, index) => {
        slide.addText(formatPoint(item.values[index] ?? 0), {
          x: point.x - 0.4, y: point.y - 0.36, w: 0.8, h: 0.22,
          fontFace: "Calibri", fontSize: 11, bold: true, color: INK, align: "center", margin: 0,
        });
      });
    }
  });
  cats.forEach((cat, index) => {
    const cx = xOf(index);
    slide.addText(cat, {
      x: cx - 0.45, y: plotTop + plotH + 0.06, w: 0.9, h: 0.28,
      fontFace: "Calibri", fontSize: cats.length > 8 ? 9 : 11, color: MUTED, align: "center", margin: 0,
    });
  });
}

function drawColumns(pres: Pres, slide: Slide, chart: SourceChart, table: SourceTable | null, x: number, y: number, w: number, h: number) {
  const cats = chart.categories.slice(0, 8);
  const series = chart.series.slice(0, 4);
  const stacked = chart.kind === "stacked";
  const peaks = cats.map((_, index) => stacked
    ? series.reduce((total, item) => total + Math.max(0, item.values[index] ?? 0), 0)
    : Math.max(0, ...series.map((item) => item.values[index] ?? 0)));
  const max = axisMax(Math.max(0, ...peaks));
  const gap = 0.16;
  const groupW = (w - gap * Math.max(0, cats.length - 1)) / Math.max(1, cats.length);
  const plotH = Math.max(1.2, h - 0.5);
  const base = y + plotH;
  slide.addShape(pres.shapes.LINE, { x, y: base, w, h: 0, line: { color: "E4E1EC", width: 1.25 } });
  cats.forEach((cat, index) => {
    const groupX = x + index * (groupW + gap);
    if (stacked) {
      let cursor = 0;
      const barW = Math.min(0.62, groupW * 0.55);
      const bx = groupX + (groupW - barW) / 2;
      series.forEach((item, seriesIndex) => {
        const value = Math.max(0, item.values[index] ?? 0);
        const bh = (value / max) * (plotH - 0.2);
        if (bh > 0.02) {
          slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: bx, y: base - cursor - bh, w: barW, h: bh,
            fill: { color: CHART_COLORS[seriesIndex] ?? "3200BE" },
            rectRadius: 0.04,
          });
        }
        cursor += bh;
      });
    } else {
      const barW = Math.min(0.46, (groupW - 0.12) / Math.max(1, series.length));
      const barsW = series.length * barW + Math.max(0, series.length - 1) * 0.06;
      const start = groupX + (groupW - barsW) / 2;
      const lonePeak = series.length === 1
        ? series[0]!.values.reduce((best, value, valueIndex) => (value > (series[0]!.values[best] ?? -Infinity) ? valueIndex : best), 0)
        : -1;
      series.forEach((item, seriesIndex) => {
        const value = Math.max(0, item.values[index] ?? 0);
        const bh = (value / max) * (plotH - 0.28);
        const bx = start + seriesIndex * (barW + 0.06);
        const tone = series.length === 1 && index !== lonePeak ? "D9D0F5" : CHART_COLORS[seriesIndex] ?? "3200BE";
        if (bh > 0.02) {
          slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: bx, y: base - bh, w: barW, h: bh,
            fill: { color: tone },
            rectRadius: 0.05,
          });
        }
        slide.addText(rowDisplay(table, cat, [item], formatPoint(value)), {
          x: bx - 0.1, y: base - bh - 0.24, w: barW + 0.2, h: 0.22,
          fontFace: "Calibri", fontSize: 11, bold: true, color: INK, align: "center", margin: 0,
        });
      });
    }
    slide.addText(cat, {
      x: groupX, y: base + 0.06, w: groupW, h: 0.32,
      fontFace: "Calibri", fontSize: 12, color: MUTED, align: "center", margin: 0,
    });
  });
}

function addOriginChart(pres: Pres, source: SourceSlide, type: (typeof SCALE)["m"]) {
  const chart = source.chart as SourceChart;
  const slide = pres.addSlide();
  speak(slide, source);
  slide.background = { color: WHITE };
  const mode = plotMode(chart);
  const eyebrow = chartEyebrow(chart);
  let cursor = 0.32;
  if (eyebrow) {
    kicker(slide, eyebrow, 0.55, cursor, "C23FC2", type.kicker);
    cursor = 0.64;
  }
  slide.addText(source.title.trim() || "Chart", {
    x: 0.55,
    y: cursor,
    w: 12.2,
    h: 0.48,
    fontFace: "Calibri",
    fontSize: type.title - 6,
    bold: true,
    color: INK,
    margin: 0,
  });
  cursor += 0.52;
  const lines = source.context.split(/\n/).map((line) => line.trim()).filter(Boolean);
  const points = lines.filter((line) => /^[•●▪‣]/.test(line)).map((line) => line.replace(/^[•●▪‣]\s*/, ""));
  const prose = lines.filter((line) => !/^[•●▪‣]/.test(line)).join(" ");
  if (prose) {
    const height = Math.min(1.05, 0.26 + prose.length / 240);
    slide.addText(prose.slice(0, 420), {
      x: 0.55,
      y: cursor,
      w: 12.2,
      h: height,
      fontFace: "Calibri",
      fontSize: 15,
      color: "322C49",
      margin: 0,
      valign: "top",
    });
    cursor += height + 0.08;
  }
  if (points.length) {
    const height = Math.min(1.35, 0.28 * points.length + 0.04);
    slide.addText(points.map((text, index) => ({
      text: text.slice(0, 180),
      options: { bullet: { code: "25CF", color: "E91E8C" }, breakLine: index < points.length - 1, color: "322C49" },
    })), {
      x: 0.55,
      y: cursor,
      w: 12.2,
      h: height,
      fontFace: "Calibri",
      fontSize: 14,
      paraSpaceAfter: 3,
      margin: 0,
      valign: "top",
    });
    cursor += height + 0.06;
  }
  const chartTop = cursor + 0.04;
  const chartH = Math.max(1.5, 6.72 - chartTop);
  const inline = mode === "score" || mode === "rank" || mode === "matrix" || mode === "trend" || mode === "columns";
  const limit = mode === "rank" ? 14 : mode === "trend" || mode === "funnel" ? 16 : mode === "matrix" ? 12 : mode === "score" ? 8 : 12;
  const labels = chart.categories.slice(0, limit);
  const data = chart.series.slice(0, chart.kind === "donut" ? 1 : 4).map((series) => ({
    name: series.name,
    labels: [...labels],
    values: labels.map((_, index) => series.values[index] ?? 0),
  }));
  const table = source.table;
  const beside = slideShow(source) === "chart" ? null : table;
  const side = source.images.length > 0 && !inline;
  const chartW = !inline && (beside || side) ? 8.05 : 12.2;
  const colors = data.length === 1 ? ["3200BE"] : [...CHART_COLORS];
  if (mode === "score") drawScore(pres, slide, chart, table, 0.55, chartTop, chartW, chartH);
  else if (mode === "rank") drawRank(pres, slide, chart, table, 0.55, chartTop, chartW, chartH);
  else if (mode === "matrix") drawMatrix(pres, slide, chart, table, 0.55, chartTop, chartW, chartH);
  else if (mode === "trend") drawTrend(pres, slide, chart, 0.45, chartTop, chartW, chartH);
  else if (mode === "columns") drawColumns(pres, slide, chart, table, 0.55, chartTop, chartW, chartH);
  else if (mode === "funnel") {
    const drawnChart = {
      ...chart,
      categories: labels,
      series: chart.series.map((series) => ({ ...series, values: labels.map((_, index) => series.values[index] ?? 0) })),
    };
    addFunnelDrawing(pres, slide, drawnChart, table, 0.45, chartTop, chartW, chartH);
  } else if (mode === "donut") {
    slide.addChart(pres.charts.DOUGHNUT, data, {
      x: table || side ? 0.4 : 2.2,
      y: chartTop,
      w: table || side ? chartW : 8.6,
      h: chartH,
      showLegend: true,
      legendPos: "r",
      legendFontSize: 14,
      legendColor: INK,
      showPercent: false,
      showValue: false,
      holeSize: 62,
      chartColors: ["3200BE", "9A2F9A", "7C8CFF", "D4CEDF"],
      chartArea: { fill: { color: WHITE } },
      plotArea: { fill: { color: WHITE } },
    });
  } else {
    const ceiling = axisMax(Math.max(0, ...labels.map((_, index) => data.reduce((total, series) => total + Math.max(0, series.values[index] ?? 0), 0))));
    slide.addChart(pres.charts.BAR, data, {
      x: 0.4,
      y: chartTop,
      w: chartW,
      h: chartH,
      barDir: "col",
      barGrouping: "stacked",
      showLegend: data.length > 1,
      legendPos: "b",
      showValue: false,
      chartColors: colors,
      chartArea: { fill: { color: WHITE } },
      plotArea: { fill: { color: WHITE } },
      catAxisLabelColor: MUTED,
      valAxisLabelColor: MUTED,
      catAxisLabelFontSize: 11,
      catAxisLabelFontFace: "Calibri",
      valAxisLabelFontFace: "Calibri",
      valAxisMinVal: 0,
      valAxisMaxVal: ceiling,
    });
  }
  const extra = [
    labels.length < chart.categories.length ? `First ${labels.length} of ${chart.categories.length} points.` : "",
    chart.series.length > 1 && mode === "donut" ? "The first series is drawn." : "",
    source.extraCharts > 0 ? "Another chart is on this slide and is not drawn." : "",
    chartFoot(chart),
  ].filter(Boolean).join(" ");
  if (!inline && beside) addSourceTable(slide, beside, 8.65, chartTop, 4.2, chartH);
  else if (side) placeKept(slide, source.images, 9.15, chartTop, 3.7, chartH);
  if (extra) {
    slide.addText(extra, {
      x: 0.55,
      y: 6.95,
      w: 12.2,
      h: 0.28,
      fontFace: "Calibri",
      fontSize: 11,
      color: MUTED,
      margin: 0,
    });
  }
}

function placeCopy(slide: Slide, text: string, x: number, y: number, w: number, bottom: number): number {
  let cursor = y;
  for (const block of copyBlocks(text)) {
    if (cursor >= bottom - 0.3) break;
    if (block.kind === "list") {
      const room = Math.max(1, Math.floor((bottom - cursor) / 0.3));
      const points = block.points.slice(0, room);
      const height = Math.min(bottom - cursor, 0.3 * points.length + 0.04);
      slide.addText(points.map((line, index) => ({
        text: line,
        options: { bullet: { code: "25CF", color: "E91E8C" }, breakLine: index < points.length - 1, color: "322C49" },
      })), {
        x, y: cursor, w, h: height,
        fontFace: "Calibri", fontSize: 15, paraSpaceAfter: 4, margin: 0, valign: "top",
      });
      cursor += height + 0.08;
      continue;
    }
    const lines = block.text.split("\n").map((line) => line.trim()).filter(Boolean);
    const across = Math.max(24, Math.floor((w * 72) / 9));
    const wrapped = lines.reduce((total, line) => total + Math.max(1, Math.ceil(line.length / across)), 0);
    const height = Math.min(Math.max(0.3, bottom - cursor), 0.28 * wrapped + 0.08);
    slide.addText(lines.map((line, index) => ({
      text: line,
      options: { breakLine: index < lines.length - 1 },
    })), {
      x, y: cursor, w, h: height,
      fontFace: "Calibri", fontSize: 16, color: "322C49", margin: 0, valign: "top",
    });
    cursor += height + 0.08;
  }
  return cursor;
}

function addOriginNumbers(pres: Pres, source: SourceSlide, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  speak(slide, source);
  slide.background = { color: WHITE };
  const numbers = source.numbers.slice(0, 6);
  kicker(slide, numbers.length === 1 ? "Figure" : "Figures", 0.65, 0.38, "C23FC2", type.kicker);
  slide.addText(source.title.trim() || "Figures", {
    x: 0.65,
    y: 0.72,
    w: 12,
    h: 0.6,
    fontFace: "Calibri",
    fontSize: type.title - 4,
    bold: true,
    color: INK,
    margin: 0,
  });
  if (numbers.length <= 1) {
    const number = numbers[0];
    const body = figureBody(source).trim();
    const busy = copyBlocks(body).some((block) => block.kind === "list") || body.length > 180;
    const numberW = busy ? 0.9 : 12;
    slide.addText(number?.display || "", {
      x: 0.65,
      y: busy ? 1.45 : 1.7,
      w: numberW,
      h: busy ? 0.85 : 1.2,
      fontFace: "Calibri",
      fontSize: number && number.display.length > 8 ? 36 : busy ? 40 : type.rate,
      bold: true,
      color: INK,
      margin: 0,
      valign: "middle",
    });
    if (number?.label) {
      slide.addText(number.label, {
        x: busy ? 1.55 : 0.65,
        y: busy ? 1.78 : 3.05,
        w: busy ? 10 : 12,
        h: 0.45,
        fontFace: "Calibri",
        fontSize: 20,
        color: INK,
        margin: 0,
        valign: "middle",
      });
    }
    if (body) placeCopy(slide, body, 0.65, busy ? 2.5 : 3.6, source.images.length ? 7.4 : 12, 7.15);
    if (source.images.length) placeKept(slide, source.images, 8.5, 3.5, 4.2, 2.8);
    return;
  }
  const columns = numbers.length === 2 || numbers.length === 4 ? 2 : 3;
  const rows = Math.ceil(numbers.length / columns);
  const gap = 0.18;
  const width = (12 - gap * (columns - 1)) / columns;
  const height = rows === 1 ? 2.15 : 1.7;
  numbers.forEach((number, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const x = 0.65 + column * (width + gap);
    const y = 1.6 + row * (height + gap);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x,
      y,
      w: width,
      h: height,
      fill: { color: "F7F5FB" },
      rectRadius: 0.1,
    });
    slide.addText(number.display, {
      x: x + 0.24,
      y: y + 0.28,
      w: width - 0.48,
      h: 0.7,
      fontFace: "Calibri",
      fontSize: number.display.length > 8 ? 22 : 32,
      bold: true,
      color: INK,
      margin: 0,
    });
    if (number.label) {
      slide.addText(number.label, {
        x: x + 0.24,
        y: y + 1.1,
        w: width - 0.48,
        h: 0.4,
        fontFace: "Calibri",
        fontSize: 14,
        color: MUTED,
        margin: 0,
      });
    }
  });
  const body = figureBody(source).trim();
  if (body) placeCopy(slide, body, 0.65, Math.min(6.4, 1.6 + rows * (height + gap) + 0.12), 12, 7.15);
  if (source.images.length) placeKept(slide, source.images, 0.65, 5.35, 12, 1.6);
}

function addOriginProse(pres: Pres, source: SourceSlide, type: (typeof SCALE)["m"], picture: boolean) {
  const slide = pres.addSlide();
  speak(slide, source);
  slide.background = { color: WHITE };
  const images = source.images;
  const side = images.length > 0 && !picture;
  kicker(slide, picture ? "Kept" : `Slide ${source.index}`, 0.65, 0.42, "C23FC2", type.kicker);
  if (source.title.trim() || !picture) {
    slide.addText(source.title.trim() || "Context", {
      x: 0.65,
      y: 0.82,
      w: side ? 7.2 : 12,
      h: 0.7,
      fontFace: "Calibri",
      fontSize: type.title - 4,
      bold: true,
      color: INK,
      margin: 0,
      valign: "top",
    });
  }
  if (picture && images.length > 0) {
    placeKept(slide, images, 0.7, source.title.trim() ? 1.8 : 1.15, 11.9, source.title.trim() ? 5 : 5.6);
    return;
  }
  const text = picture
    ? "This slide is a picture. The file did not store a title, a figure, or a chart as text."
    : source.context.trim().slice(0, 1600) || "This slide had no readable text.";
  slide.addText(text, {
    x: 0.65,
    y: 1.8,
    w: side ? 7.1 : 12,
    h: 4.8,
    fontFace: "Calibri",
    fontSize: text.length > 700 ? 14 : type.body,
    color: INK,
    margin: 0,
    valign: "top",
  });
  if (side) placeKept(slide, images, 8.15, 1.7, 4.5, 4.8);
}
