import * as XLSX from "xlsx";
import { emptyBook, readSheet } from "./parse.ts";
import { gridProgress, projectNext } from "./grid.ts";
import { rateOutlook, seriesProgressList } from "./reading.ts";
import { slidesFor } from "./slides.tsx";
import type { Book } from "./book.ts";

const failures: string[] = [];
const expect = (label: string, actual: unknown, wanted: unknown) => {
  if (actual !== wanted) failures.push(`${label}: ${String(actual)} !== ${String(wanted)}`);
};
const includes = (label: string, actual: string, wanted: string) => {
  if (!actual.includes(wanted)) failures.push(`${label}: ${JSON.stringify(actual)} missing ${JSON.stringify(wanted)}`);
};

function bookBuffer(sheets: Record<string, (string | number | null)[][]>): ArrayBuffer {
  const workbook = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  }
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}

const repairs = readSheet(
  bookBuffer({
    Repairs: [
      ["Month", "Repairs", "Cost"],
      ["Jan-24", 100, 400],
      ["Feb-24", 110, 420],
      ["Mar-24", 130, 450],
      ["Apr-24", 150, null],
    ],
    Notes: [
      ["Month", "Extra"],
      ["Jan-24", 1],
      ["Feb-24", 2],
      ["Mar-24", 3],
    ],
  }),
  "repairs-by-month.xlsx",
);
const grid = repairs.book.grid;
expect("grid client", repairs.book.client, "Repairs by month");
expect("grid not a book", repairs.book.active, 0);
expect("grid rows", grid?.labels.join(","), "Jan-24,Feb-24,Mar-24,Apr-24");
expect("repairs kept", grid?.columns[0]?.values.join(","), "100,110,130,150");
expect("blank cost is not zero", grid?.columns[1]?.values[3], null);
expect("repairs projection", grid?.projection[0], 165);
expect("projection matches the line", grid?.projection[0], projectNext([100, 110, 130, 150]));
expect("projection label", grid?.projectionLabel, "May-24 proj");
expect("no sum of repairs", grid?.columns[0]?.values.includes(490), false);
includes("not added", repairs.note, "not added together");
includes("not combined", repairs.note, "Other sheets were not combined.");
const progress = grid ? gridProgress(grid) : null;
expect("progress title", progress?.title, "Progress");
expect("progress previous", progress?.left, 130);
expect("progress latest", progress?.right, 150);
expect("progress kind", progress?.kind, "pair");
const repairSlides = slidesFor(repairs.book);
expect("repair slide count", repairSlides.length, 4);
expect("first chart", repairSlides[1]?.id, "progress");
expect("cost slide", repairSlides[3]?.label, "Cost");

const years = readSheet(
  bookBuffer({
    Sheet1: [
      ["Year", "Policies"],
      [2023, 10],
      [2024, 20],
      [2025, 30],
    ],
  }),
  "policies.xlsx",
);
expect("year cadence", years.book.grid?.cadence, "year");
expect("year projection label", years.book.grid?.projectionLabel, "2026 proj");
expect("year projection", years.book.grid?.projection[0], 40);
expect("year not percent", years.book.grid?.columns[0]?.percent, false);

const rates = readSheet(
  bookBuffer({
    Rates: [
      ["Month", "Containment"],
      ["Jan-25", 80],
      ["Feb-25", 70],
      ["Mar-25", 60],
    ],
  }),
  "rates.xlsx",
);
const rateProgress = rates.book.grid ? gridProgress(rates.book.grid) : null;
expect("rate is percent", rates.book.grid?.columns[0]?.percent, true);
expect("rate direction", rateProgress?.title, "Not progress");
expect("rate latest", rateProgress?.left, 60);
expect("rate rest", rateProgress?.right, 40);
expect("rate kind", rateProgress?.kind, "share");

const book = readSheet(
  bookBuffer({
    Book: [
      ["Client", "Active subscribers", "Claims", "Fulfilled claims", "Approved claims"],
      ["Northline Care", 48200, 19280, 9640, 14460],
    ],
  }),
  "northline.xlsx",
);
expect("book active", book.book.active, 48200);
expect("book claims", book.book.claims, 19280);
expect("book is not a grid", book.book.grid, null);

const voice = readSheet(
  bookBuffer({
    Voice: [
      ["Month", "Total", "Contained", "%"],
      ["Mar-25", 11, 9, 82],
      ["Apr-25", 80000, 66000, 83],
      ["May-25", 81000, 64000, 79],
      ["Jun-25", 82000, 60000, 73],
      ["Jul-25", 83000, 58000, 70],
      ["Aug-25", 84000, 52000, 62],
      ["Sep-25", 85000, 48000, 56],
      ["Oct-25", 20000, 11000, 55],
    ],
  }),
  "voice.xlsx",
);
expect("voice stays a series", voice.book.reading?.tracks.length, 1);
expect("voice is not a grid", voice.book.grid, null);
const cards = voice.book.reading ? seriesProgressList(voice.book.reading) : [];
expect("voice latest month", cards[0]?.label, "Sep-25");
expect("voice contained", cards[0]?.contained, 48000);
expect("voice rest", cards[0]?.rest, 37000);
expect("voice direction", cards[0]?.direction, "Not progress");
const outlook = voice.book.reading ? rateOutlook(voice.book.reading.tracks[0]!) : null;
expect("voice outlook label", outlook?.label, "Nov-25 proj");
expect("voice outlook rate", outlook?.rate, 46);
const voiceSlides = slidesFor(voice.book).map((slide) => slide.id);
expect("voice first chart", voiceSlides[1], "progress");
includes("voice order", voiceSlides.join(","), "cover,progress,track,outlook");

const withCounts = {
  ...emptyBook(),
  active: 48200,
  claims: 19280,
  client: "Harbor review",
  sourceDeck: {
    title: "Harbor review",
    intent: "Read 2 slides.",
    slides: [
      { index: 1, title: "Harbor review", context: "Kept.", numbers: [], chart: null, extraCharts: 0, picture: false, images: [], table: null },
      { index: 2, title: "Next", context: "Also kept.", numbers: [], chart: null, extraCharts: 0, picture: false, images: [], table: null },
    ],
  },
} satisfies Book;
const kept = slidesFor(withCounts);
expect("pptx keeps both slides", kept.length, 2);
expect("pptx first", kept[0]?.id, "origin");
expect("pptx second", kept[1]?.label, "Next");

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log("grid checks passed");
