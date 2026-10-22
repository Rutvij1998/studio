import PptxGenJS from "pptxgenjs";
import { readDeckText } from "./parse.ts";
import { bulletLines, chartSentence, copyBlocks, isSource, leadCopy, piecesOf, presentKind, readSourceDeck, slideRead, sourceSummary } from "./sourceDeck.ts";

const failures: string[] = [];
const expect = (label: string, actual: unknown, wanted: unknown) => {
  if (actual !== wanted) failures.push(`${label}: ${String(actual)} !== ${String(wanted)}`);
};
const includes = (label: string, actual: string, wanted: string) => {
  if (!actual.includes(wanted)) failures.push(`${label}: ${JSON.stringify(actual)} missing ${JSON.stringify(wanted)}`);
};

const PNG = "image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

async function fixture(): Promise<ArrayBuffer> {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  pres.layout = "DECK";
  const cover = pres.addSlide();
  cover.addText("Harbor review", { x: 0.7, y: 0.6, w: 10, h: 0.8, fontSize: 36, fontFace: "Calibri" });
  cover.addText("Third quarter, kept as written.", { x: 0.7, y: 1.6, w: 10, h: 0.5, fontSize: 18, fontFace: "Calibri" });

  const columns = pres.addSlide();
  columns.addText("Containment by month", { x: 0.6, y: 0.4, w: 10, h: 0.5, fontSize: 28, fontFace: "Calibri" });
  columns.addChart(pres.charts.BAR, [
    { name: "Containment", labels: ["Apr", "May", "Jun"], values: [83, 80, 72] },
  ], { x: 0.6, y: 1.2, w: 11, h: 4.5, barDir: "col", barGrouping: "clustered", showValue: false });
  columns.addText("Full months only.", { x: 0.6, y: 6.4, w: 10, h: 0.4, fontSize: 14, fontFace: "Calibri" });

  const figures = pres.addSlide();
  figures.addText("The call", { x: 0.6, y: 0.4, w: 10, h: 0.5, fontSize: 28, fontFace: "Calibri" });
  figures.addText("We handled 1.9M IVA calls. Containment is above 70%. Authentication is above 90%.", {
    x: 0.6, y: 1.4, w: 11, h: 1.2, fontSize: 20, fontFace: "Calibri",
  });

  const prose = pres.addSlide();
  prose.addText("What changed", { x: 0.6, y: 0.4, w: 10, h: 0.5, fontSize: 28, fontFace: "Calibri" });
  prose.addText("The desk kept the same population and did not add a cause for the movement.", {
    x: 0.6, y: 1.4, w: 11, h: 1.2, fontSize: 18, fontFace: "Calibri",
  });

  const picture = pres.addSlide();
  picture.addImage({ data: PNG, x: 1, y: 1, w: 4, h: 3 });

  const donut = pres.addSlide();
  donut.addText("Mix", { x: 0.6, y: 0.4, w: 10, h: 0.5, fontSize: 28, fontFace: "Calibri" });
  donut.addChart(pres.charts.DOUGHNUT, [
    { name: "Channel", labels: ["Voice", "Chat"], values: [64, 36] },
  ], { x: 0.6, y: 1.2, w: 8, h: 4.5, showPercent: false });

  const line = pres.addSlide();
  line.addText("Pace", { x: 0.6, y: 0.4, w: 10, h: 0.5, fontSize: 28, fontFace: "Calibri" });
  line.addChart(pres.charts.LINE, [
    { name: "Pace", labels: ["Apr", "May", "Jun"], values: [10, 20, 30] },
  ], { x: 0.6, y: 1.2, w: 11, h: 4.5 });

  const stacked = pres.addSlide();
  stacked.addText("Opened", { x: 0.6, y: 0.4, w: 10, h: 0.5, fontSize: 28, fontFace: "Calibri" });
  stacked.addChart(pres.charts.BAR, [
    { name: "Open", labels: ["Apr", "May"], values: [4, 5] },
    { name: "Closed", labels: ["Apr", "May"], values: [6, 7] },
  ], { x: 0.6, y: 1.2, w: 11, h: 4.5, barDir: "col", barGrouping: "stacked" });

  const bars = pres.addSlide();
  bars.addText("Queues", { x: 0.6, y: 0.4, w: 10, h: 0.5, fontSize: 28, fontFace: "Calibri" });
  bars.addChart(pres.charts.BAR, [
    { name: "Waiting", labels: ["Desk", "Repair"], values: [12, 9] },
  ], { x: 0.6, y: 1.2, w: 11, h: 4.5, barDir: "bar", barGrouping: "clustered" });

  const out = await pres.write({ outputType: "arraybuffer" });
  return out as ArrayBuffer;
}

const buffer = await fixture();
const read = await readSourceDeck(buffer, "harbor-review.pptx");
const flat = await readDeckText(buffer);
expect("slides", read.deck.slides.length, 10);
includes("flat", flat.text, "Harbor review");
includes("intent", sourceSummary(read.deck), "Read 10 slides");
includes("intent charts", sourceSummary(read.deck), "5 charts");
includes("intent figures", sourceSummary(read.deck), "3 figures");
includes("intent picture", sourceSummary(read.deck), "1 picture");
includes("lead", sourceSummary(read.deck), "Lead: Harbor review.");
includes("kept", sourceSummary(read.deck), "Every slide is kept.");
includes("colors", sourceSummary(read.deck), "Charts are redrawn in the studio colors.");
expect("counts do not drop the deck", isSource({ sourceDeck: read.deck, active: 48200, claims: 19280 }), true);
expect("empty deck is not a source", isSource({ sourceDeck: { ...read.deck, slides: [] }, active: 0, claims: 0 }), false);

const [cover, opening, columns, figures, prose, picture, donut, line, stacked, bars] = read.deck.slides;
expect("cover kind", cover ? presentKind(cover) : "", "cover");
expect("cover title", cover?.title, "Harbor review");
expect("cover is only the title", cover?.context ?? "", "");
expect("cover has no pictures", cover?.images.length ?? -1, 0);
includes("opening kept", opening?.context ?? "", "Third quarter");
expect("cover chart", cover?.chart, null);

expect("columns kind", columns ? presentKind(columns) : "", "chart");
expect("columns chart", columns?.chart?.kind, "columns");
expect("columns cats", columns?.chart?.categories.join(","), "Apr,May,Jun");
expect("columns series", columns?.chart?.series[0]?.name, "Containment");
expect("columns values", columns?.chart?.series[0]?.values.join(","), "83,80,72");
includes("columns context", columns?.context ?? "", "Full months only");
includes("columns sentence", columns?.chart ? chartSentence(columns.chart) : "", "Columns of Containment");

expect("figures kind", figures ? presentKind(figures) : "", "numbers");
const values = (figures?.numbers ?? []).map((number) => number.value).join(",");
expect("figure values", values, "1900000,70,90");
const labels = (figures?.numbers ?? []).map((number) => number.label).join("|");
expect("figure labels", labels, "IVA calls|Containment|Authentication");
expect("display m", figures?.numbers[0]?.display, "1.9M");
expect("display pct", figures?.numbers[1]?.display, "70%");

expect("prose kind", prose ? presentKind(prose) : "", "prose");
includes("prose", prose?.context ?? "", "same population");
expect("picture kind", picture ? presentKind(picture) : "", "picture");
expect("picture kept", picture?.images.length ?? 0, 1);
includes("picture note", picture ? slideRead(picture) : "", "picture stays");
expect("donut", donut?.chart?.kind, "donut");
expect("donut cats", donut?.chart?.categories.join(","), "Voice,Chat");
expect("line", line?.chart?.kind, "line");
expect("line value", line?.chart?.series[0]?.values[2], 30);
expect("stacked", stacked?.chart?.kind, "stacked");
expect("stacked series", stacked?.chart?.series.map((series) => series.name).join(","), "Open,Closed");
expect("bars", bars?.chart?.kind, "bars");
expect("bars cat", bars?.chart?.categories.join(","), "Desk,Repair");

async function measured() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  pres.layout = "DECK";
  const cover = pres.addSlide();
  cover.addText("Jira Dashboard Creation", { x: 0.5, y: 0.4, w: 10, h: 0.5, fontSize: 28, fontFace: "Calibri" });
  cover.addText("09/29/2026", { x: 0.5, y: 1.2, w: 6, h: 0.4, fontSize: 18, fontFace: "Calibri" });

  const kpi = pres.addSlide();
  kpi.addText("BI Reporting Landscape Overview", { x: 0.5, y: 0.3, w: 12, h: 0.4, fontSize: 20, fontFace: "Calibri" });
  kpi.addText("The BI reporting landscape currently consists of 1,026 reports across DP and GTP, with 492 reports identified as active or currently used.", { x: 0.5, y: 0.8, w: 12, h: 0.5, fontSize: 14, fontFace: "Calibri" });
  kpi.addTable([
    [{ text: "KPI" }, { text: "Count" }],
    [{ text: "Total Reports" }, { text: "1,026" }],
    [{ text: "DP Reports" }, { text: "~996" }],
    [{ text: "GTP Reports" }, { text: "30" }],
    [{ text: "Active / Used Reports" }, { text: "492" }],
  ], { x: 0.5, y: 1.6, w: 8, h: 2.4, fontFace: "Calibri" });

  const months = pres.addSlide();
  months.addText("Calls by month", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 20, fontFace: "Calibri" });
  months.addTable([
    [{ text: "Month" }, { text: "Calls" }],
    [{ text: "Apr" }, { text: "10" }],
    [{ text: "May" }, { text: "20" }],
    [{ text: "Jun" }, { text: "30" }],
  ], { x: 0.5, y: 1, w: 6, h: 2, fontFace: "Calibri" });

  const years = pres.addSlide();
  years.addText("Desk by year", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 20, fontFace: "Calibri" });
  years.addTable([
    [{ text: "Queue" }, { text: "2024" }, { text: "2025" }],
    [{ text: "Desk" }, { text: "10" }, { text: "14" }],
    [{ text: "Repair" }, { text: "8" }, { text: "9" }],
  ], { x: 0.5, y: 1, w: 7, h: 2, fontFace: "Calibri" });

  const funnel = pres.addSlide();
  funnel.addText("Signup funnel", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 20, fontFace: "Calibri" });
  funnel.addTable([
    [{ text: "Stage" }, { text: "People" }],
    [{ text: "Visit" }, { text: "100" }],
    [{ text: "Start" }, { text: "60" }],
    [{ text: "Finish" }, { text: "20" }],
  ], { x: 0.5, y: 1, w: 6, h: 2, fontFace: "Calibri" });

  const share = pres.addSlide();
  share.addText("Channel share", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 20, fontFace: "Calibri" });
  share.addTable([
    [{ text: "Channel" }, { text: "Share" }],
    [{ text: "Voice" }, { text: "64" }],
    [{ text: "Chat" }, { text: "36" }],
  ], { x: 0.5, y: 1, w: 6, h: 1.6, fontFace: "Calibri" });

  const sources = pres.addSlide();
  sources.addText("Reports by Data Source", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 20, fontFace: "Calibri" });
  sources.addTable([
    [{ text: "Data Source" }, { text: "DP" }, { text: "GTP" }, { text: "Total" }],
    [{ text: "Unknown" }, { text: "281" }, { text: "0" }, { text: "281" }],
    [{ text: "EDW" }, { text: "310" }, { text: "0" }, { text: "310" }],
    [{ text: "eSecuritel" }, { text: "271" }, { text: "0" }, { text: "271" }],
  ], { x: 0.5, y: 1, w: 8, h: 2, fontFace: "Calibri" });

  const step = pres.addSlide();
  step.addText("Step 1: Open Dashboards", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 20, fontFace: "Calibri" });
  step.addText("Open the dashboard list from the menu. Choose a 2-column layout.", { x: 0.5, y: 1, w: 10, h: 0.4, fontSize: 14, fontFace: "Calibri" });

  const read = await readSourceDeck(await pres.write({ outputType: "arraybuffer" }) as ArrayBuffer, "measured.pptx");
  const [opened, landscape, byMonth, byYear, signup, channel, bySource, openedStep] = read.deck.slides;
  expect("date is not a figure", opened?.numbers.length ?? -1, 0);
  expect("opening kind", opened ? presentKind(opened) : "", "cover");
  expect("kpi kind", landscape?.chart?.kind ?? "", "bars");
  expect("kpi values", landscape?.chart?.series[0]?.values.join(","), "1026,996,30,492");
  expect("kpi tilde kept", landscape?.table?.rows[1]?.[1] ?? "", "~996");
  includes("kpi approx", landscape?.chart?.drawn ?? "", "approximate");
  includes("kpi sentence", landscape?.context ?? "", "1,026 reports");
  expect("month kind", byMonth?.chart?.kind ?? "", "line");
  includes("month over month", byMonth?.chart?.drawn ?? "", "Month over month");
  expect("year kind", byYear?.chart?.kind ?? "", "columns");
  includes("year over year", byYear?.chart?.drawn ?? "", "Year over year");
  expect("funnel kind", signup?.chart?.kind ?? "", "funnel");
  expect("funnel top", signup?.chart?.series[0]?.values[0], 100);
  expect("share kind", channel?.chart?.kind ?? "", "donut");
  expect("source series", bySource?.chart?.series.map((series) => series.name).join(","), "DP,GTP");
  includes("total stays", bySource?.chart?.drawn ?? "", "Total stays");
  expect("step is not a figure", openedStep?.numbers.length ?? -1, 0);
  expect("step kind", openedStep ? presentKind(openedStep) : "", "prose");
}

await measured();

async function crowd() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  pres.layout = "DECK";
  const cover = pres.addSlide();
  cover.addText("Weekly review", { x: 0.5, y: 0.4, w: 10, h: 0.5, fontSize: 28, fontFace: "Calibri" });
  cover.addText("Stuti Jain", { x: 0.5, y: 1.3, w: 8, h: 0.4, fontSize: 18, fontFace: "Calibri" });
  const pics = pres.addSlide();
  pics.addText("Arrange the gadgets", { x: 0.4, y: 0.3, w: 10, h: 0.4, fontSize: 22, fontFace: "Calibri" });
  pics.addText("Drag the first gadget.\nPlace the filter.\nSave the layout.\nShare the board.", {
    x: 0.4, y: 0.8, w: 10, h: 1, fontSize: 14, fontFace: "Calibri",
  });
  for (let index = 0; index < 6; index += 1) {
    pics.addImage({ data: PNG, x: 0.4 + (index % 3) * 1.4, y: 2.2 + Math.floor(index / 3) * 1.2, w: 1, h: 0.8 });
  }
  const read = await readSourceDeck(await pres.write({ outputType: "arraybuffer" }) as ArrayBuffer, "crowd.pptx");
  expect("crowd slides", read.deck.slides.length, 3);
  expect("presenter stays", read.deck.slides[0]?.context, "Stuti Jain");
  expect("presenter slide has no pictures", read.deck.slides[0]?.images.length ?? -1, 0);
  const first = read.deck.slides[1];
  const second = read.deck.slides[2];
  expect("first picture group", first?.images.length ?? 0, 3);
  expect("second picture group", second?.images.length ?? 0, 3);
  includes("flow", second?.context ?? "", "Continued.");
  includes("first agenda", first?.context ?? "", "Drag the first gadget.");
  includes("second agenda", second?.context ?? "", "Save the layout.");
}

await crowd();

async function listed() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  pres.layout = "DECK";
  const slide = pres.addSlide();
  slide.addText("Live Demo", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 28, fontFace: "Calibri" });
  slide.addText([
    { text: "Create a dashboard with one of these queries.", options: { breakLine: true } },
    { text: "Tickets with no parent", options: { bullet: true, breakLine: true } },
    { text: "Tickets with no FV", options: { bullet: true } },
  ], { x: 0.5, y: 1, w: 10, h: 1.4, fontSize: 16, fontFace: "Calibri" });
  const read = await readSourceDeck(await pres.write({ outputType: "arraybuffer" }) as ArrayBuffer, "listed.pptx");
  const follow = read.deck.slides[1];
  expect("list slides", read.deck.slides.length, 2);
  expect("list cover is only the title", read.deck.slides[0]?.context ?? "x", "");
  expect("list items", follow ? bulletLines(follow).join("|") : "", "Tickets with no parent|Tickets with no FV");
  includes("list lead", follow ? leadCopy(follow) : "", "Create a dashboard");
}

await listed();

const broken = copyBlocks("Blanket verification creates collateral damage\n• More friction for legitimate customers• Longer approval times• Higher support contacts");
const list = broken.find((block) => block.kind === "list");
expect("inline points", list && list.kind === "list" ? list.points.join("|") : "", "More friction for legitimate customers|Longer approval times|Higher support contacts");

async function breaks() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  pres.layout = "DECK";
  const slide = pres.addSlide();
  slide.addText("What We Should Not Do", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 28, fontFace: "Calibri" });
  slide.addText([
    { text: "Blanket verification creates collateral damage", options: { breakLine: true } },
    { text: "• More friction for legitimate customers", options: { breakLine: true } },
    { text: "• Longer approval times", options: { breakLine: true } },
    { text: "• Higher support contacts", options: { breakLine: true } },
  ], { x: 0.5, y: 1, w: 10, h: 2, fontSize: 16, fontFace: "Calibri" });
  slide.addText("2", { x: 0.5, y: 4, w: 2, h: 0.5, fontSize: 32, fontFace: "Calibri" });
  const read = await readSourceDeck(await pres.write({ outputType: "arraybuffer" }) as ArrayBuffer, "breaks.pptx");
  const follow = read.deck.slides[1];
  includes("line break kept", follow?.context ?? "", "\n• Longer approval times");
  const points = copyBlocks(follow?.context ?? "").find((block) => block.kind === "list");
  expect("broken lines", points && points.kind === "list" ? points.points.join("|") : "", "More friction for legitimate customers|Longer approval times|Higher support contacts");
}

await breaks();

async function cards() {
  const problem = piecesOf([
    "This is not just a fraud problem — it is a claims-quality and ownership problem.",
    "18% ↑ claim cost\nVolume is flat, so the business is paying more per claim.",
    "5+ day delays\nGenuine customers are waiting too long for verification.",
    "$2M savings target\nReduce cost without harming approvals or partner trust.",
    "Finance\nFraud + duplicates",
    "Support\nCustomer delays",
    "The gap is ownership\nWe need one operating system for claim risk, cost, and customer speed.",
    "1",
  ]);
  const stat = problem.find((piece) => piece.kind === "card" && piece.display === "18%");
  const five = problem.find((piece) => piece.kind === "card" && piece.display === "5+");
  const savings = problem.find((piece) => piece.kind === "card" && piece.display === "$2M");
  const finance = problem.find((piece) => piece.kind === "card" && piece.title === "Finance");
  expect("page number is not a card", problem.some((piece) => piece.kind === "card" && (piece.display === "1" || piece.title === "1")), false);
  expect("cost display", stat && stat.kind === "card" ? stat.label : "", "↑ claim cost");
  expect("delay display", five && five.kind === "card" ? five.label : "", "day delays");
  expect("savings display", savings && savings.kind === "card" ? savings.label : "", "savings target");
  expect("finance detail", finance && finance.kind === "card" ? finance.detail : "", "Fraud + duplicates");
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  pres.layout = "DECK";
  const slide = pres.addSlide();
  slide.addText("The Real Problem", { x: 0.5, y: 0.2, w: 10, h: 0.4, fontSize: 28, fontFace: "Calibri" });
  slide.addText("This is not just a fraud problem.", { x: 0.5, y: 0.7, w: 10, h: 0.3, fontSize: 14, fontFace: "Calibri" });
  slide.addText([
    { text: "18% ↑ claim cost", options: { breakLine: true } },
    { text: "Volume is flat, so the business is paying more per claim." },
  ], { x: 0.5, y: 1.1, w: 4, h: 0.8, fontSize: 14, fontFace: "Calibri" });
  slide.addText([
    { text: "Finance", options: { breakLine: true } },
    { text: "Fraud + duplicates" },
  ], { x: 5, y: 1.1, w: 4, h: 0.6, fontSize: 14, fontFace: "Calibri" });
  slide.addText([
    { text: "Data Science", options: { breakLine: true } },
    { text: "Model trust + explainability" },
  ], { x: 5, y: 2, w: 4, h: 0.6, fontSize: 14, fontFace: "Calibri" });
  slide.addText("1", { x: 12, y: 6.8, w: 0.4, h: 0.3, fontSize: 12, fontFace: "Calibri" });
  const read = await readSourceDeck(await pres.write({ outputType: "arraybuffer" }) as ArrayBuffer, "problem.pptx");
  const follow = read.deck.slides[1];
  const pieces = follow?.pieces ?? [];
  expect("cover is the title only", read.deck.slides[0]?.context ?? "x", "");
  expect("shapes stay apart", pieces.length, 4);
  expect("read keeps the stat", pieces.some((piece) => piece.kind === "card" && piece.display === "18%" && piece.label === "↑ claim cost"), true);
  expect("read keeps the tile", pieces.some((piece) => piece.kind === "card" && piece.title === "Finance" && piece.detail === "Fraud + duplicates"), true);
  expect("read drops the page number", pieces.some((piece) => piece.kind === "card" && piece.display === "1"), false);
  expect("department stays a card", pieces.some((piece) => piece.kind === "card" && piece.title === "Data Science" && piece.detail === "Model trust + explainability"), true);
}

await cards();

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log("source deck checks passed");
