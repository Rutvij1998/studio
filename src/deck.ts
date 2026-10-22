import PptxGenJS from "pptxgenjs";
import { coverCopy } from "./brief";
import { conversationOf, formatCount, hasBroughtIn, viewBook, type Book, type Proof, type SourceChart, type SourceImage, type SourceSlide, type SourceTable } from "./book";
import { deckFileName, type DeckOptions, type TypeSize } from "./options";
import { slidesFor } from "./slides";
import { compactVolume, isResults, proofPhrase, storyLine, storyReadLines } from "./story";
import { factLine, factReadLines, isFacts } from "./facts";
import { chartFit, comboPair } from "./fit";
import { columnPoints, formatGridPoint, gridChartNote, gridLine, gridProgress, isGrid } from "./grid";
import { isOutline, outlineLine } from "./outline";
import { isSeries, rateOutlook, readingLine, seriesProgressList, trackNote, trackTotals } from "./reading";
import { axisMax, axisSpan, bulletLines, chartEyebrow, chartFoot, copyBlocks, figureBody, formatPoint, isContinued, isSource, leadCopy, pieceBands, pieceFit, pieceFrame, plotMode, rowDisplay, slideRole, spareHeaders } from "./sourceDeck";

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

function rasterSize(dataUrl: string): { width: number; height: number } | null {
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
        const width = ((bytes[cursor + 7] ?? 0) | 0) << 8 | (bytes[cursor + 8] ?? 0);
        return width > 0 && height > 0 ? { width, height } : null;
      }
      const size = ((bytes[cursor + 2] ?? 0) << 8) | (bytes[cursor + 3] ?? 0);
      if (size < 2) break;
      cursor += 2 + size;
    }
  }
  return null;
}

function placeKept(slide: Slide, images: SourceImage[], x: number, y: number, w: number, h: number) {
  const ready = images
    .slice(0, 4)
    .map((image) => ({ data: imageData(image.dataUrl), size: rasterSize(image.dataUrl) }))
    .filter((image): image is { data: string; size: { width: number; height: number } | null } => Boolean(image.data));
  if (ready.length === 0) return;
  const cols = ready.length === 1 ? 1 : ready.length === 4 ? 2 : ready.length;
  const rows = Math.ceil(ready.length / cols);
  const gap = 0.12;
  const cellW = (w - gap * (cols - 1)) / cols;
  const cellH = (h - gap * (rows - 1)) / rows;
  ready.forEach((image, index) => {
    const column = index % cols;
    const row = Math.floor(index / cols);
    const left = x + column * (cellW + gap);
    const top = y + row * (cellH + gap);
    const ratio = image.size && image.size.width > 0 ? image.size.width / image.size.height : cellW / cellH;
    const fittedW = Math.min(cellW, cellH * ratio);
    const fittedH = fittedW / ratio;
    slide.addImage({
      data: image.data,
      x: left + (cellW - fittedW) / 2,
      y: top,
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
  slide.background = { color: STAGE };
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

function addOrigin(pres: Pres, book: Book, index: number, type: (typeof SCALE)["m"]) {
  const source = book.sourceDeck?.slides.find((slide) => slide.index === index) ?? null;
  if (!source) return;
  const role = slideRole(source);
  const structured = (source.pieces ?? []).some((piece) => piece.kind === "card" || piece.kind === "list");
  if (role === "cover") addOriginCover(pres, book, source, type);
  else if (role !== "chart" && structured && source.images.length === 0) addOriginPieces(pres, source, type);
  else if (role === "chart" && source.chart) addOriginChart(pres, source, type);
  else if (role === "figures") addOriginNumbers(pres, source, type);
  else if (role === "pictures" && source.picture && source.images.length === 0) addOriginProse(pres, source, type, true);
  else addOriginSheet(pres, source, type);
}

function addOriginPieces(pres: Pres, source: SourceSlide, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  slide.background = { color: WHITE };
  const frame = pieceFrame(pieceFit(source));
  const size = frame.type;
  const x0 = 0.58;
  const width = 12.17;
  let cursor = frame.inch(40);
  if (isContinued(source)) {
    kicker(slide, "Continued", x0, cursor, "C23FC2", type.kicker);
    cursor += frame.inch(38);
  }
  if (source.title.trim()) {
    const chars = Math.max(8, Math.floor(1168 / (size.title * 0.56)));
    const lines = Math.max(1, Math.ceil(source.title.trim().length / chars));
    const titleH = frame.inch(lines * Math.ceil(size.title * 1.12));
    slide.addText(source.title.trim(), {
      x: x0, y: cursor, w: width, h: titleH,
      fontFace: "Calibri", fontSize: frame.pt(size.title), bold: true, color: INK, margin: 0, valign: "top",
    });
    cursor += titleH + frame.inch(16);
  }
  for (const band of pieceBands(source.pieces ?? [])) {
    if (band.kind === "prose") {
      const parts = frame.proseParts(band.text);
      parts.forEach((part, index) => {
        if (index) cursor += frame.inch(10);
        const height = frame.inch(Math.max(part.height, 18));
        if (part.kind === "list") {
          slide.addText(part.points.map((point, pointIndex) => ({
            text: point,
            options: { bullet: { code: "25A0", color: BRAND }, breakLine: pointIndex < part.points.length - 1, color: "322C49" },
          })), {
            x: x0, y: cursor, w: width, h: height,
            fontFace: "Calibri", fontSize: frame.pt(size.bullet), paraSpaceAfter: 4, margin: 0, valign: "top",
          });
        } else {
          slide.addText(part.text, {
            x: x0, y: cursor, w: width, h: height,
            fontFace: "Calibri", fontSize: frame.pt(size.lead), color: "322C49", margin: 0, valign: "top",
          });
        }
        cursor += height;
      });
      cursor += frame.inch(16);
      continue;
    }
    if (band.kind === "lists") {
      const columns = Math.max(1, band.items.length);
      const gap = frame.inch(16);
      const columnW = (width - gap * (columns - 1)) / columns;
      const height = frame.inch(frame.listHeight(band.items));
      band.items.forEach((item, index) => {
        const x = x0 + index * (columnW + gap);
        let top = cursor;
        const titleLines = item.title ? Math.max(1, item.title.split("\n").length) : 0;
        const titleH = titleLines ? frame.inch(titleLines * Math.ceil(size.column * 1.25) + 10) : 0;
        if (item.title) {
          slide.addText(item.title, {
            x, y: top, w: columnW, h: titleH,
            fontFace: "Calibri", fontSize: frame.pt(size.column), bold: true, color: INK, margin: 0, valign: "top",
          });
          top += titleH;
        }
        slide.addText(item.points.map((point, pointIndex) => ({
          text: point,
          options: { bullet: { code: "25A0", color: BRAND }, breakLine: pointIndex < item.points.length - 1, color: "322C49" },
        })), {
          x, y: top, w: columnW, h: Math.max(0.3, height - titleH),
          fontFace: "Calibri", fontSize: frame.pt(size.bullet), paraSpaceAfter: 4, margin: 0, valign: "top",
        });
      });
      cursor += height + frame.inch(16);
      continue;
    }
    const cards = band.items;
    const across = band.tone === "chips" ? 4 : 3;
    const gap = frame.inch(14);
    for (let row = 0; row < cards.length; row += across) {
      const slice = cards.slice(row, row + across);
      const wide = slice.length === 1 ? 1 : across;
      const cardW = slice.length === 1 ? width : (width - gap * (across - 1)) / across;
      const heights = slice.map((item) => frame.cardHeight(item, wide));
      const rowH = frame.inch(Math.max(...heights));
      slice.forEach((item, index) => {
        const x = x0 + index * (cardW + gap);
        slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
          x, y: cursor, w: cardW, h: rowH,
          fill: { color: "F7F5FB" },
          rectRadius: 0.1,
        });
        const inset = 0.16;
        let top = cursor + 0.12;
        if (item.display) {
          const displayH = frame.inch(size.stat);
          slide.addText(item.display, {
            x: x + inset, y: top, w: cardW - inset * 2, h: displayH,
            fontFace: "Calibri", fontSize: frame.pt(size.stat), bold: true, color: INK, margin: 0, valign: "top",
          });
          top += displayH;
          const labelH = frame.inch(Math.ceil(size.statLabel * 1.3));
          slide.addText(item.label || item.title, {
            x: x + inset, y: top, w: cardW - inset * 2, h: labelH,
            fontFace: "Calibri", fontSize: frame.pt(size.statLabel), bold: true, color: INK, margin: 0, valign: "top",
          });
          top += labelH + 0.04;
          if (item.detail) {
            slide.addText(item.detail, {
              x: x + inset, y: top, w: cardW - inset * 2, h: Math.max(0.24, cursor + rowH - top - 0.1),
              fontFace: "Calibri", fontSize: frame.pt(size.statBody), color: MUTED, margin: 0, valign: "top",
            });
          }
          return;
        }
        const titleFont = band.tone === "chips" ? size.chip : size.cardTitle;
        const titleLines = Math.max(1, item.title.split("\n").length);
        const titleH = frame.inch(titleLines * Math.ceil(titleFont * 1.3));
        slide.addText(item.title, {
          x: x + inset, y: top, w: cardW - inset * 2, h: titleH,
          fontFace: "Calibri", fontSize: frame.pt(titleFont), bold: true, color: INK, margin: 0, valign: "top",
        });
        if (item.detail) {
          slide.addText(item.detail, {
            x: x + inset, y: top + titleH + 0.04, w: cardW - inset * 2, h: Math.max(0.24, cursor + rowH - (top + titleH) - 0.12),
            fontFace: "Calibri", fontSize: frame.pt(size.cardBody), color: "322C49", margin: 0, valign: "top",
          });
        }
      });
      cursor += rowH + gap;
    }
    cursor += frame.inch(2);
  }
}

function addOriginCover(pres: Pres, book: Book, source: SourceSlide, type: (typeof SCALE)["m"]) {
  const slide = pres.addSlide();
  slide.background = { color: STAGE };
  if (book.logo) placeMark(pres, slide, book.logo);
  const title = source.title.trim() || book.sourceDeck?.title || "Untitled";
  const wide = title.length > 42;
  const titleSize = wide ? Math.max(28, type.title - 4) : title.length > 22 ? type.title + 4 : type.title + 14;
  const titleTop = 2.2;
  const widthIn = book.logo ? 9.2 : 11.8;
  const charsPerLine = Math.max(16, Math.floor((widthIn * 72) / (titleSize * 0.45)));
  const lines = Math.min(4, Math.ceil(title.length / charsPerLine));
  const titleH = 0.12 + lines * (titleSize / 72) * 1.2;
  slide.addText(title, {
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
  slide.background = { color: WHITE };
  const bullets = bulletLines(source);
  const continued = isContinued(source);
  const lead = leadCopy(source);
  const images = source.images;
  let cursor = 0.46;
  if (continued) {
    kicker(slide, "Continued", 0.7, cursor, "C23FC2", type.kicker);
    cursor = 0.82;
  }
  if (source.title.trim()) {
    slide.addText(source.title.trim(), {
      x: 0.7,
      y: cursor,
      w: 12,
      h: images.length ? 0.52 : 0.72,
      fontFace: "Calibri",
      fontSize: images.length ? Math.max(24, type.title - 12) : type.title - 2,
      bold: true,
      color: INK,
      margin: 0,
      valign: "top",
    });
    cursor += images.length ? 0.58 : 0.82;
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
      options: { bullet: { code: "25A0", color: BRAND }, breakLine: index < bullets.length - 1, color: INK },
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

function addSourceTable(slide: Slide, table: SourceTable, x: number, y: number, w: number, h: number) {
  const headers = table.headers.length ? table.headers : ["Item"];
  const header = headers.map((text) => ({
    text: text || " ",
    options: { fill: { color: INK }, color: WHITE, bold: true, align: "left" as const },
  }));
  const body = table.rows.map((row, index) => headers.map((_, cell) => ({
    text: row[cell] || " ",
    options: {
      fill: { color: index % 2 ? "F7F5FB" : WHITE },
      color: INK,
      align: (cell === headers.length - 1 ? "right" : "left") as "left" | "right",
      bold: cell === headers.length - 1,
    },
  })));
  const lead = Math.min(w * 0.62, w - 0.7 * Math.max(1, headers.length - 1));
  const rest = (w - lead) / Math.max(1, headers.length - 1);
  slide.addTable([header, ...body], {
    x,
    y,
    w,
    h,
    colW: headers.map((_, index) => (index === 0 ? lead : rest)),
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

function drawScore(pres: Pres, slide: Slide, chart: SourceChart, table: SourceTable | null, x: number, y: number, w: number) {
  const series = chart.series[0];
  if (!series) return;
  const cats = chart.categories.slice(0, 5);
  const values = cats.map((_, index) => series.values[index] ?? 0);
  const max = Math.max(1, ...values.map((value) => Math.max(0, value)));
  const named = cats.findIndex((category) => /^total\b/i.test(category.trim()));
  const lead = named >= 0 ? named : values.reduce((best, value, index) => (value > (values[best] ?? 0) ? index : best), 0);
  const rest = cats.map((_, index) => index).filter((index) => index !== lead);
  slide.addText(rowDisplay(table, cats[lead] ?? "", [series], formatPoint(values[lead] ?? 0)), {
    x, y, w: Math.min(w, 8), h: 0.9,
    fontFace: "Calibri", fontSize: 48, bold: true, color: INK, margin: 0,
  });
  slide.addText(cats[lead] ?? "", {
    x, y: y + 0.92, w: Math.min(w, 8), h: 0.32,
    fontFace: "Calibri", fontSize: 16, color: "322C49", margin: 0,
  });
  drawTrack(pres, slide, x, y + 1.32, w, 0.12, "3200BE", Math.max(0, values[lead] ?? 0) / max);
  if (!rest.length) return;
  const gap = 0.28;
  const colW = (w - gap * (rest.length - 1)) / rest.length;
  rest.forEach((index, slot) => {
    const left = x + slot * (colW + gap);
    const top = y + 1.75;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: left, y: top - 0.08, w: colW, h: 1.22,
      fill: { color: "F7F5FB" },
      rectRadius: 0.08,
    });
    slide.addText(rowDisplay(table, cats[index] ?? "", [series], formatPoint(values[index] ?? 0)), {
      x: left + 0.16, y: top, w: colW - 0.32, h: 0.5,
      fontFace: "Calibri", fontSize: 28, bold: true, color: INK, margin: 0,
    });
    slide.addText(cats[index] ?? "", {
      x: left + 0.16, y: top + 0.5, w: colW - 0.32, h: 0.32,
      fontFace: "Calibri", fontSize: 14, color: "322C49", margin: 0,
    });
    drawTrack(pres, slide, left + 0.16, top + 0.92, Math.max(0.4, colW - 0.48), 0.1, "3200BE", Math.max(0, values[index] ?? 0) / max);
  });
}

function drawRank(pres: Pres, slide: Slide, chart: SourceChart, table: SourceTable | null, x: number, y: number, w: number, h: number) {
  const series = chart.series[0];
  if (!series) return;
  const cats = chart.categories.slice(0, 14);
  const values = cats.map((_, index) => series.values[index] ?? 0);
  const max = Math.max(1, ...values.map((value) => Math.max(0, value)));
  const rowH = Math.min(0.36, h / Math.max(1, cats.length));
  cats.forEach((cat, index) => {
    const top = y + index * rowH;
    slide.addText(cat, {
      x, y: top, w: 3.15, h: rowH,
      fontFace: "Calibri", fontSize: cats.length > 10 ? 11 : 13, color: INK, margin: 0, valign: "middle",
    });
    drawTrack(pres, slide, x + 3.25, top + (rowH - 0.1) / 2, w - 4.55, 0.1, "3200BE", Math.max(0, values[index] ?? 0) / max);
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
      const numW = Math.min(0.85, colW * 0.34);
      drawTrack(pres, slide, left, top + (rowH - 0.08) / 2, Math.max(0.3, colW - numW - 0.16), 0.08, CHART_COLORS[seriesIndex] ?? "3200BE", Math.max(0, value) / (maxes[seriesIndex] ?? 1));
      slide.addText(rowDisplay(table, cat, [item], formatPoint(value)), {
        x: left + colW - numW - 0.08, y: top, w: numW, h: rowH,
        fontFace: "Calibri", fontSize: 12, bold: true, color: INK, align: "right", margin: 0, valign: "middle",
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
    points.forEach((point) => {
      slide.addShape(pres.shapes.OVAL, {
        x: point.x - 0.07, y: point.y - 0.07, w: 0.14, h: 0.14,
        fill: { color },
        line: { color: "FFFFFF", width: 1.25 },
      });
    });
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
      series.forEach((item, seriesIndex) => {
        const value = Math.max(0, item.values[index] ?? 0);
        const bh = (value / max) * (plotH - 0.28);
        const bx = start + seriesIndex * (barW + 0.06);
        if (bh > 0.02) {
          slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: bx, y: base - bh, w: barW, h: bh,
            fill: { color: CHART_COLORS[seriesIndex] ?? "3200BE" },
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
      options: { bullet: { code: "25A0", color: BRAND }, breakLine: index < points.length - 1, color: "322C49" },
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
  const limit = mode === "rank" ? 14 : mode === "trend" || mode === "funnel" ? 16 : mode === "matrix" ? 12 : mode === "score" ? 5 : 12;
  const labels = chart.categories.slice(0, limit);
  const data = chart.series.slice(0, chart.kind === "donut" ? 1 : 4).map((series) => ({
    name: series.name,
    labels: [...labels],
    values: labels.map((_, index) => series.values[index] ?? 0),
  }));
  const table = source.table;
  const side = source.images.length > 0 && !inline;
  const chartW = !inline && (table || side) ? 8.05 : 12.2;
  const colors = data.length === 1 ? ["3200BE"] : [...CHART_COLORS];
  if (mode === "score") drawScore(pres, slide, chart, table, 0.55, chartTop, chartW);
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
  if (!inline && table) addSourceTable(slide, table, 8.65, chartTop, 4.2, chartH);
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
        options: { bullet: { code: "25A0", color: BRAND }, breakLine: index < points.length - 1, color: "322C49" },
      })), {
        x, y: cursor, w, h: height,
        fontFace: "Calibri", fontSize: 15, paraSpaceAfter: 4, margin: 0, valign: "top",
      });
      cursor += height + 0.08;
      continue;
    }
    const lines = block.text.split("\n").map((line) => line.trim()).filter(Boolean);
    const height = Math.min(bottom - cursor, 0.28 * lines.length + 0.06);
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
