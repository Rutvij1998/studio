import JSZip from "jszip";
import * as XLSX from "xlsx";
import {
  SAMPLE_BOOK,
  SAMPLE_JOURNEY,
  SAMPLE_LISTENING,
  SAMPLE_MONTHS,
  largestRemainder,
  type Book,
  type JourneyStage,
  type ListeningRow,
  type MonthStack,
  type Track,
} from "./book";
import { gridFromMatrix, gridPeriod } from "./grid";
import { readingFromTracks, trackFromRows } from "./reading";

const TEXT_FIELDS = ["client", "market", "period", "peerName", "brief", "recommendation"] as const;
const NUMBER_FIELDS = [
  "active",
  "claims",
  "fulfilled",
  "approved",
  "peerActive",
  "peerClaims",
  "peerFulfilled",
] as const;

type TextField = (typeof TEXT_FIELDS)[number];
type NumberField = (typeof NUMBER_FIELDS)[number];
type Field = TextField | NumberField;

const ALIAS: Record<string, Field> = {
  client: "client",
  name: "client",
  "client name": "client",
  program: "client",
  market: "market",
  country: "market",
  period: "period",
  window: "period",
  active: "active",
  "active subscribers": "active",
  "active subs": "active",
  "active subscribers eop": "active",
  subscribers: "active",
  claims: "claims",
  "claim count": "claims",
  fulfilled: "fulfilled",
  "fulfilled claims": "fulfilled",
  approved: "approved",
  "approved claims": "approved",
  peer: "peerName",
  "peer name": "peerName",
  comparison: "peerName",
  "peer active": "peerActive",
  "peer active subscribers": "peerActive",
  "peer claims": "peerClaims",
  "peer fulfilled": "peerFulfilled",
  "peer fulfilled claims": "peerFulfilled",
  recommendation: "recommendation",
  "what to do": "recommendation",
  brief: "brief",
  prompt: "brief",
  notes: "brief",
};

export type SheetRead = {
  book: Book;
  sheetName: string;
  usedRow: number;
  note: string;
};

const norm = (value: unknown) =>
  String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

const asText = (value: unknown) => String(value ?? "").replace(/\s+/g, " ").trim();

function asCount(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return Math.max(0, Math.round(value));
  const text = asText(value).replace(/[$,%\s]/g, "").replace(/,/g, "");
  if (!text) return null;
  const parsed = Number(text);
  if (!Number.isFinite(parsed)) return null;
  return Math.max(0, Math.round(parsed));
}

function apply(book: Book, field: Field, value: unknown) {
  if ((TEXT_FIELDS as readonly string[]).includes(field)) {
    const text = asText(value);
    if (text) book[field as TextField] = text;
    return;
  }
  const count = asCount(value);
  if (count != null) book[field as NumberField] = count;
}

function rowObject(headers: string[], row: unknown[]): Partial<Record<Field, unknown>> {
  const out: Partial<Record<Field, unknown>> = {};
  headers.forEach((header, index) => {
    const field = ALIAS[norm(header)];
    if (!field) return;
    out[field] = row[index];
  });
  return out;
}

function score(record: Partial<Record<Field, unknown>>): number {
  return NUMBER_FIELDS.reduce((sum, field) => sum + (asCount(record[field]) ? 1 : 0), 0);
}

type Grid = (string | number | null)[][];

const SERIES_NAMES = new Set([
  "listening",
  "sentiment",
  "mentions",
  "journey",
  "stages",
  "months",
  "month",
  "trend",
]);

function sheetRows(workbook: XLSX.WorkBook, name: string): Grid {
  const grid = XLSX.utils.sheet_to_json<(string | number | null)[]>(workbook.Sheets[name]!, {
    header: 1,
    raw: true,
    defval: "",
  });
  return grid.filter((row) => row.some((cell) => asText(cell) !== ""));
}

function isBookGrid(rows: Grid): boolean {
  if (rows.length < 2) return false;
  const headers = rows[0]!.map((cell) => asText(cell));
  if (headers.some((header) => {
    const field = ALIAS[norm(header)];
    return field != null && (NUMBER_FIELDS as readonly string[]).includes(field);
  })) {
    return true;
  }
  return rows.some((row) => {
    const field = ALIAS[norm(row[0])];
    return field != null && (NUMBER_FIELDS as readonly string[]).includes(field);
  });
}

function columnIndex(headers: string[], aliases: string[]): number {
  return headers.findIndex((header) => aliases.includes(norm(header)));
}

function readListening(rows: Grid): ListeningRow[] {
  if (rows.length < 2) return [];
  const headers = rows[0]!.map((cell) => asText(cell));
  const labelAt = columnIndex(headers, ["label", "topic", "source", "name"]);
  const countAt = columnIndex(headers, ["count", "mentions", "volume"]);
  if (labelAt < 0 || countAt < 0) return [];
  const positiveAt = columnIndex(headers, ["positive", "pos"]);
  const neutralAt = columnIndex(headers, ["neutral", "neu"]);
  const negativeAt = columnIndex(headers, ["negative", "neg"]);
  const out: ListeningRow[] = [];
  for (const row of rows.slice(1)) {
    const label = asText(row[labelAt]);
    const count = asCount(row[countAt]);
    if (!label || count == null) continue;
    const positive = positiveAt < 0 ? null : asCount(row[positiveAt]);
    const neutral = neutralAt < 0 ? null : asCount(row[neutralAt]);
    const negative = negativeAt < 0 ? null : asCount(row[negativeAt]);
    const parts =
      positive == null && neutral == null && negative == null
        ? { positive: 0, neutral: count, negative: 0 }
        : {
            positive: positive ?? 0,
            negative: negative ?? 0,
            neutral: neutral ?? Math.max(0, count - (positive ?? 0) - (negative ?? 0)),
          };
    if (count === 0 && parts.positive + parts.neutral + parts.negative === 0) continue;
    out.push({ label, count, ...parts });
  }
  return out;
}

function readJourney(rows: Grid): JourneyStage[] {
  if (rows.length < 2) return [];
  const headers = rows[0]!.map((cell) => asText(cell));
  const labelAt = columnIndex(headers, ["label", "stage", "step", "name"]);
  const mentionsAt = columnIndex(headers, ["mentions", "count", "volume"]);
  const positiveAt = columnIndex(headers, ["positive", "pos"]);
  const neutralAt = columnIndex(headers, ["neutral", "neu"]);
  const negativeAt = columnIndex(headers, ["negative", "neg"]);
  if (labelAt < 0 || positiveAt < 0 || neutralAt < 0 || negativeAt < 0) return [];
  const out: JourneyStage[] = [];
  for (const row of rows.slice(1)) {
    const label = asText(row[labelAt]);
    const positive = asCount(row[positiveAt]) ?? 0;
    const neutral = asCount(row[neutralAt]) ?? 0;
    const negative = asCount(row[negativeAt]) ?? 0;
    const mentions = mentionsAt < 0 ? positive + neutral + negative : (asCount(row[mentionsAt]) ?? 0);
    if (!label || (mentions === 0 && positive + neutral + negative === 0)) continue;
    const sum = positive + neutral + negative;
    const percents =
      mentions > 0 && sum === mentions && sum !== 100
        ? largestRemainder([negative, neutral, positive], mentions)
        : [negative, neutral, positive];
    out.push({
      label,
      mentions,
      negative: percents[0] ?? 0,
      neutral: percents[1] ?? 0,
      positive: percents[2] ?? 0,
    });
  }
  return out;
}

function readMonths(rows: Grid): MonthStack[] {
  if (rows.length < 2) return [];
  const headers = rows[0]!.map((cell) => asText(cell));
  const labelAt = columnIndex(headers, ["month", "label", "period"]);
  const fulfilledAt = columnIndex(headers, ["fulfilled", "fulfilled claims"]);
  const openAt = columnIndex(headers, ["open", "approved open", "approved still open", "approved, still open"]);
  const notAt = columnIndex(headers, ["not approved", "notapproved", "declined"]);
  if (labelAt < 0 || (fulfilledAt < 0 && openAt < 0 && notAt < 0)) return [];
  const out: MonthStack[] = [];
  for (const row of rows.slice(1)) {
    const label = asText(row[labelAt]);
    if (!label) continue;
    const fulfilled = fulfilledAt < 0 ? 0 : (asCount(row[fulfilledAt]) ?? 0);
    const open = openAt < 0 ? 0 : (asCount(row[openAt]) ?? 0);
    const notApproved = notAt < 0 ? 0 : (asCount(row[notAt]) ?? 0);
    if (fulfilled + open + notApproved === 0) continue;
    out.push({ label, fulfilled, open, notApproved });
  }
  return out;
}

function namedSheet(workbook: XLSX.WorkBook, names: string[]): string | undefined {
  return workbook.SheetNames.find((name) => names.includes(norm(name)));
}

function attachSeries(workbook: XLSX.WorkBook, book: Book, note: string): string {
  const listeningName = namedSheet(workbook, ["listening", "sentiment", "mentions"]);
  const journeyName = namedSheet(workbook, ["journey", "stages"]);
  const monthsName = namedSheet(workbook, ["months", "month", "trend"]);
  const skipped: string[] = [];
  if (listeningName) {
    const rows = sheetRows(workbook, listeningName);
    book.listening = readListening(rows);
    if (rows.length > 1 && book.listening.length === 0) {
      skipped.push(`${listeningName} needs Topic and Mentions`);
    }
  }
  if (journeyName) {
    const rows = sheetRows(workbook, journeyName);
    book.journey = readJourney(rows);
    if (rows.length > 1 && book.journey.length === 0) {
      skipped.push(`${journeyName} needs Stage, Positive, Neutral, and Negative`);
    }
  }
  if (monthsName) {
    const rows = sheetRows(workbook, monthsName);
    book.months = readMonths(rows);
    if (rows.length > 1 && book.months.length === 0) {
      skipped.push(`${monthsName} needs Month and at least one claim column`);
    }
  }
  const found = [
    book.listening.length ? `${book.listening.length} listening rows` : "",
    book.journey.length ? `${book.journey.length} journey stages` : "",
    book.months.length ? `${book.months.length} months` : "",
  ].filter(Boolean);
  const extra = found.length ? ` Also read ${found.join(", ")}.` : "";
  const warn = skipped.length ? ` Skipped ${skipped.join("; ")}.` : "";
  return `${note}${extra}${warn}`;
}

function parseBookGrid(rows: Grid, fileName: string, sheetName: string): SheetRead {
  if (rows.length < 2) throw new Error("The sheet needs a header row and one book.");
  const book = emptyBook();
  const headers = rows[0]!.map((cell) => asText(cell));
  const known = headers.some((header) => ALIAS[norm(header)]);

  if (known) {
    const records = rows.slice(1).map((row) => rowObject(headers, row));
    let best = 0;
    let bestWeight = -1;
    records.forEach((record, index) => {
      const name = norm(record.client);
      const weight = score(record) + (name === "total" || name === "all" ? 3 : 0);
      if (weight > bestWeight) {
        best = index;
        bestWeight = weight;
      }
    });
    if (bestWeight <= 0) {
      throw new Error(
        "No numbers in that sheet. Use Active subscribers, Claims, Fulfilled claims, and Approved claims.",
      );
    }
    const chosen = records[best] ?? {};
    (Object.keys(chosen) as Field[]).forEach((field) => apply(book, field, chosen[field]));
    const extra = records.length > 1 ? " Used the fullest row. A program split comes later." : "";
    return {
      book,
      sheetName,
      usedRow: best + 2,
      note: `Read ${fileName}, sheet ${sheetName}, row ${best + 2}.${extra}`,
    };
  }

  let matched = 0;
  rows.forEach((row) => {
    const field = ALIAS[norm(row[0])];
    if (!field) return;
    apply(book, field, row[1]);
    matched += 1;
  });
  if (matched === 0) {
    throw new Error(
      "No known columns. Use Client, Active subscribers, Claims, Fulfilled claims, and Approved claims.",
    );
  }
  return {
    book,
    sheetName,
    usedRow: 0,
    note: `Read ${fileName}, sheet ${sheetName}, as a label and value list.`,
  };
}

function labelCell(value: unknown): string {
  if (typeof value === "number" && value > 20000 && value < 80000 && XLSX.SSF?.parse_date_code) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      return `${names[parsed.m - 1] ?? "Jan"}-${String(parsed.y).slice(2)}`;
    }
  }
  return asText(value);
}

function tracksFromWorkbook(workbook: XLSX.WorkBook): Track[] {
  const tracks: Track[] = [];
  for (const name of workbook.SheetNames) {
    const rows = sheetRows(workbook, name).map((row) => row.map((cell) => labelCell(cell)));
    const track = trackFromRows(rows, name);
    if (track) tracks.push(track);
  }
  return tracks;
}

export function readSheet(buffer: ArrayBuffer, fileName: string): SheetRead {
  const workbook = XLSX.read(buffer, { type: "array" });
  if (workbook.SheetNames.length === 0) throw new Error("That file has no sheets.");
  const tracks = tracksFromWorkbook(workbook);
  const namedBook = workbook.SheetNames.find((name) => norm(name) === "book");
  const sheetName =
    namedBook ??
    workbook.SheetNames.find((name) => !SERIES_NAMES.has(norm(name)) && isBookGrid(sheetRows(workbook, name))) ??
    workbook.SheetNames.find((name) => !SERIES_NAMES.has(norm(name)));
  let read: SheetRead | null = null;
  let failure: Error | null = null;
  if (sheetName) {
    try {
      read = parseBookGrid(sheetRows(workbook, sheetName), fileName, sheetName);
    } catch (error) {
      failure = error instanceof Error ? error : new Error("Could not read that sheet.");
    }
  }
  const counted = read != null && (read.book.active > 0 || read.book.claims > 0);
  if (counted && read) {
    read.note = attachSeries(workbook, read.book, read.note);
    return read;
  }
  if (tracks.length > 0) {
    const book = read?.book ?? emptyBook();
    const reading = readingFromTracks(tracks, book.client || fileName);
    book.reading = reading;
    book.grid = null;
    if (!book.client) book.client = reading.title;
    if (!book.period) book.period = reading.period;
    const names = tracks.map((track) => track.name).join(" and ");
    return {
      book,
      sheetName: read?.sheetName ?? tracks[0]!.name,
      usedRow: 0,
      note: `Read ${names} from ${fileName}. ${reading.intent}`,
    };
  }
  const charted = chartSheet(workbook, fileName);
  if (charted) {
    if (!charted.book.client && read?.book.client) charted.book.client = read.book.client;
    return charted;
  }
  if (failure) throw failure;
  if (!read) {
    throw new Error(
      "Add a Book sheet with Active subscribers, Claims, Fulfilled claims, and Approved claims.",
    );
  }
  read.note = attachSeries(workbook, read.book, read.note);
  return read;
}

function chartSheet(workbook: XLSX.WorkBook, fileName: string): SheetRead | null {
  const found: { name: string; weight: number; grid: NonNullable<Book["grid"]> }[] = [];
  for (const name of workbook.SheetNames) {
    const grid = gridFromMatrix(gridMatrix(workbook, name), name);
    if (!grid) continue;
    const weight = grid.columns.reduce((sum, column) => sum + column.values.filter((value) => value != null).length, 0);
    found.push({ name, weight, grid });
  }
  found.sort((left, right) => right.weight - left.weight);
  const best = found[0];
  if (!best) return null;
  const book = emptyBook();
  book.grid = best.grid;
  book.client = fileTitle(fileName) || best.grid.title;
  book.period = gridPeriod(best.grid);
  const others = found.length > 1 ? " Other sheets were not combined." : "";
  return {
    book,
    sheetName: best.name,
    usedRow: 0,
    note: `Read ${fileName}, sheet ${best.name}. ${best.grid.note}${others}`,
  };
}

function gridMatrix(workbook: XLSX.WorkBook, name: string): (string | number | null)[][] {
  return sheetRows(workbook, name).map((row) =>
    row.map((cell, index) => {
      if (index === 0) {
        const label = labelCell(cell);
        return label || null;
      }
      if (typeof cell === "number" && Number.isFinite(cell)) return cell;
      const text = asText(cell);
      return text || null;
    }),
  );
}

function fileTitle(fileName: string): string {
  const base = fileName.replace(/\.(xlsx|xls|csv)$/i, "").replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  if (!base || /^(sheet\d*|book)$/i.test(base)) return "";
  return base
    .split(" ")
    .map((word, index) => {
      if (index > 0 && word === word.toLowerCase()) return word;
      if (word.length <= 4 && word === word.toUpperCase() && /[A-Z]/.test(word)) return word;
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(" ")
    .slice(0, 80);
}

export function emptyBook(): Book {
  return {
    client: "",
    market: "",
    period: "",
    active: 0,
    claims: 0,
    fulfilled: 0,
    approved: 0,
    peerName: "",
    peerActive: 0,
    peerClaims: 0,
    peerFulfilled: 0,
    brief: "",
    recommendation: "",
    sourceName: "",
    sourceText: "",
    listening: [],
    journey: [],
    months: [],
    story: null,
    reading: null,
    sourceDeck: null,
    facts: null,
    grid: null,
    outline: null,
    logo: "",
  };
}

export function sampleSheetBlob(): Blob {
  const header = [
    "Client",
    "Market",
    "Period",
    "Active subscribers",
    "Claims",
    "Fulfilled claims",
    "Approved claims",
    "Peer",
    "Peer active",
    "Peer claims",
    "Peer fulfilled",
    "Recommendation",
    "Brief",
  ];
  const row = [
    SAMPLE_BOOK.client,
    SAMPLE_BOOK.market,
    SAMPLE_BOOK.period,
    SAMPLE_BOOK.active,
    SAMPLE_BOOK.claims,
    SAMPLE_BOOK.fulfilled,
    SAMPLE_BOOK.approved,
    SAMPLE_BOOK.peerName,
    SAMPLE_BOOK.peerActive,
    SAMPLE_BOOK.peerClaims,
    SAMPLE_BOOK.peerFulfilled,
    SAMPLE_BOOK.recommendation,
    SAMPLE_BOOK.brief,
  ];
  const sheet = XLSX.utils.aoa_to_sheet([header, row]);
  sheet["!cols"] = header.map((label) => ({ wch: Math.max(18, label.length + 2) }));
  const listening = XLSX.utils.aoa_to_sheet([
    ["Topic", "Mentions", "Positive", "Neutral", "Negative"],
    ...SAMPLE_LISTENING.map((item) => [item.label, item.count, item.positive, item.neutral, item.negative]),
  ]);
  const journey = XLSX.utils.aoa_to_sheet([
    ["Stage", "Mentions", "Positive", "Neutral", "Negative"],
    ...SAMPLE_JOURNEY.map((item) => [item.label, item.mentions, item.positive, item.neutral, item.negative]),
  ]);
  const months = XLSX.utils.aoa_to_sheet([
    ["Month", "Fulfilled", "Open", "Not approved"],
    ...SAMPLE_MONTHS.map((item) => [item.label, item.fulfilled, item.open, item.notApproved]),
  ]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Book");
  XLSX.utils.book_append_sheet(workbook, listening, "Listening");
  XLSX.utils.book_append_sheet(workbook, journey, "Journey");
  XLSX.utils.book_append_sheet(workbook, months, "Months");
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

const decode = (value: string) =>
  value
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'");

export async function readDeckText(buffer: ArrayBuffer): Promise<{ slides: number; text: string }> {
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => {
      const left = Number(a.match(/(\d+)\.xml$/i)?.[1] ?? 0);
      const right = Number(b.match(/(\d+)\.xml$/i)?.[1] ?? 0);
      return left - right;
    });
  const chunks: string[] = [];
  for (const name of names) {
    const xml = await zip.file(name)!.async("string");
    const bits = [...xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map((match) => decode(match[1] ?? ""));
    const line = bits.join(" ").replace(/\s+/g, " ").trim();
    if (line) chunks.push(line);
  }
  return { slides: names.length, text: chunks.join("\n\n") };
}

export async function readWordText(buffer: ArrayBuffer): Promise<string> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    throw new Error("Could not read that Word file. Save it as .docx.");
  }
  const file = zip.file("word/document.xml");
  if (!file) throw new Error("That Word file has no document body. Save it as .docx.");
  const xml = await file.async("string");
  const paragraphs = xml
    .split(/<\/w:p>/)
    .map((block) =>
      [...block.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
        .map((match) => decode(match[1] ?? ""))
        .join(""),
    )
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  return paragraphs.join("\n\n");
}

export async function readPlainText(buffer: ArrayBuffer): Promise<string> {
  return new TextDecoder().decode(buffer).replace(/\n{3,}/g, "\n\n").trim();
}
