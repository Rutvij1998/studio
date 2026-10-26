import PptxGenJS from "pptxgenjs";
import type { SourceSlide } from "./book.ts";
import { readDeckText } from "./parse.ts";
import { addSlidePictures, bulletLines, canFlip, cardFace, chartSentence, compareSplit, compositionOf, copyBlocks, deleteSlide, duplicateSlide, figurePx, heroAside, imageSize, isSource, isTimeTitle, leadCopy, mergeSlides, moveSlide, pictureLayout, pieceDensity, piecesOf, presentKind, readSourceDeck, sameFigure, showSlide, slideRead, slideShow, sourceSummary, tableFromSlide } from "./sourceDeck.ts";

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

const lanes = piecesOf([
  "GREEN",
  "Auto-approve",
  "Known customer; device on policy; first claim in period; IMEI match; no velocity.",
  "Minutes",
  "SLA under 4 hours",
  "CS says: “Approved, device and account match.”",
  "AMBER",
  "Assisted review",
  "Multi-device household; address or email drift; model disagrees with rules.",
]);
expect("lane count", lanes.length, 2);
expect("green lane", lanes[0]?.kind === "card" ? lanes[0].title : "", "GREEN\nAuto-approve");
expect("green keeps the rule", lanes[0]?.kind === "card" ? lanes[0].detail.includes("Known customer") && lanes[0].detail.includes("SLA under 4 hours") : false, true);
expect("amber lane", lanes[1]?.kind === "card" ? lanes[1].title : "", "AMBER\nAssisted review");

const ask = piecesOf([
  "$2M",
  "Annual claim-cost reduction",
  "COO target, via a 90-day Product proposal.",
]);
expect("ask is one card", ask.length, 1);
expect("ask display", ask[0]?.kind === "card" ? ask[0].display : "", "$2M");
expect("ask label", ask[0]?.kind === "card" ? ask[0].label : "", "Annual claim-cost reduction");
includes("ask detail", ask[0]?.kind === "card" ? ask[0].detail : "", "COO target");

const signals = piecesOf(["Customer history", "Device history", "Address match", "Payment method"]);
expect("signal chips stay apart", signals.length, 4);

const constraint = piecesOf([
  "Work within the constraints",
  "Existing platform takes new validation rules.\nExternal fraud intelligence needs extra investment.\nNo auto-decline in v1.",
]);
expect("constraint is one list", constraint.length, 1);
expect("constraint title", constraint[0]?.kind === "list" ? constraint[0].title : "", "Work within the constraints");
includes("constraint point", constraint[0]?.kind === "list" ? constraint[0].points.join(" ") : "", "No auto-decline");

const reasons = piecesOf([
  "Reason codes let CS explain every decision out loud",
  "Agents never see a raw model score.",
  "Second claim on this IMEI in 90 days",
  "Email is not on the enrolled account; this IMEI has been used with other emails",
  "Shipping address differs from the enrolled address",
  "Household has N enrolled devices; this is claim K this year, consistent with the policy",
]);
expect("reason lead stays", reasons[0]?.kind === "prose" ? reasons[0].text : "", "Reason codes let CS explain every decision out loud");
expect("reason sentence stays", reasons[1]?.kind === "prose" ? reasons[1].text : "", "Agents never see a raw model score.");
expect("reasons are one list", reasons[2]?.kind === "list" ? reasons[2].points.length : 0, 4);
includes("reason example", reasons[2]?.kind === "list" ? reasons[2].points[0] ?? "" : "", "Second claim");

const compared = compareSplit([
  { kind: "card", title: "Today · one pile\nEvery claim", detail: "", display: "", label: "" },
  { kind: "prose", text: "Time burns on honest customers." },
  { kind: "card", title: "Tomorrow · route by risk", detail: "", display: "", label: "" },
  { kind: "card", title: "GREEN · auto-approve", detail: "Minutes", display: "", label: "" },
]);
expect("today column", compared?.left.length ?? 0, 2);
expect("tomorrow column", compared?.right[1]?.kind === "card" ? compared.right[1].title : "", "GREEN · auto-approve");
const aside = heroAside([
  { kind: "card", title: "Annual claim-cost reduction", detail: "COO target.", display: "$2M", label: "Annual claim-cost reduction" },
  { kind: "card", title: "Hold the customer and partner outcomes", detail: "Hold CSAT.", display: "", label: "" },
  { kind: "list", title: "Work within the constraints", points: ["No auto-decline in v1."] },
]);
expect("hero sits aside", aside?.stat.display ?? "", "$2M");
expect("hero keeps the list", aside?.rest[1]?.kind === "list" ? aside.rest[1].title : "", "Work within the constraints");
expect("mixed units do not compare", sameFigure(["18%", "5+", "$2M"]), false);
expect("percents compare", sameFigure(["12%", "64%"]), true);
expect("larger value is larger type", figurePx("64%", ["12%", "64%"], 128) > figurePx("12%", ["12%", "64%"], 128), true);
expect("order does not pick the hero", figurePx("12%", ["12%", "64%"], 128), figurePx("12%", ["64%", "12%"], 128));
expect("long money steps down", figurePx("$12,450,000", ["$12,450,000", "$4M"], 128) < 70, true);
expect("short peer stays under the long one", figurePx("$4M", ["$12,450,000", "$4M"], 128) <= figurePx("$12,450,000", ["$12,450,000", "$4M"], 128), true);
expect("two rates cascade", compositionOf(piecesOf(["12% First pass", "64% Second pass"]), "Two rates"), "cascade");
expect("mixed stays metrics", compositionOf(piecesOf([
  "18% Leak\nOne owner is missing from the handoff and nobody closes it.",
  "5+ Teams\nFive groups touch the same claim before anyone owns the outcome.",
  "$2M Cost\nThe wait costs this much and the sentence is part of the card.",
]), "The mix"), "metrics");
expect("before is a contrast", compareSplit([
  { kind: "card", title: "Before", detail: "Five days.", display: "", label: "" },
  { kind: "card", title: "After", detail: "A lane.", display: "", label: "" },
]) != null, true);
expect("phase is a sequence", isTimeTitle("Phase 1"), true);
expect("staging is not a stage", isTimeTitle("Staging"), false);
const sentenceThen = [
  { kind: "prose" as const, text: "The wait is the cost." },
  { kind: "card" as const, title: "Saved this year", detail: "", display: "$2M", label: "Saved this year" },
];
expect("sentence then figure is a hero", heroAside(sentenceThen)?.stat.display ?? "", "$2M");
expect("that hero is the composition", compositionOf(sentenceThen, "One number"), "hero");

const headed = piecesOf([
  "Work within the constraints",
  "• Existing platform takes new validation rules.\n• External fraud intelligence needs extra investment.\n• No auto-decline in v1.",
]);
expect("heading stays on the list", headed.length, 1);
expect("list title", headed[0]?.kind === "list" ? headed[0].title : "", "Work within the constraints");
expect("list points", headed[0]?.kind === "list" ? headed[0].points.length : 0, 3);

const tensions = piecesOf([
  "The lanes resolve the four tensions",
  "Fraud and CX",
  "Speed where risk is low. Scrutiny where risk is high.",
]);
expect("subtitle stays itself", tensions[0]?.kind === "card" ? tensions[0].detail : "x", "");
expect("tension keeps its sentence", tensions[1]?.kind === "card" && tensions[1].title === "Fraud and CX" ? tensions[1].detail : "", "Speed where risk is low. Scrutiny where risk is high.");

async function honest() {
  const footer = "Likewize Device Protection | Claims cost and experience | 90-day Product proposal | Confidential";
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  pres.layout = "DECK";
  const cover = pres.addSlide();
  cover.addText("Likewize", { x: 0.6, y: 0.4, w: 8, h: 0.4, fontSize: 20, fontFace: "Calibri" });
  cover.addText("DEVICE PROTECTION · 90-DAY PRODUCT PROPOSAL", { x: 0.6, y: 0.9, w: 10, h: 0.4, fontSize: 14, fontFace: "Calibri" });
  cover.addText([
    { text: "Faster honest claims.", options: { breakLine: true } },
    { text: "Tighter fraud.", options: { breakLine: true } },
    { text: "$2M without the five-day wait." },
  ], { x: 0.6, y: 1.6, w: 10, h: 1.4, fontSize: 28, fontFace: "Calibri" });
  cover.addText("Risk-tiered claims on the existing platform.", { x: 0.6, y: 3.4, w: 10, h: 0.4, fontSize: 16, fontFace: "Calibri" });
  cover.addText("Retailer pilot SLA: 80% in 24h, 95% in 48h. External fraud intelligence parked.", { x: 0.6, y: 3.9, w: 10, h: 0.4, fontSize: 16, fontFace: "Calibri" });

  const cost = pres.addSlide();
  cost.addText("Cost split", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 22, fontFace: "Calibri" });
  cost.addText("Finance and Operations can both be right until the cost is split.", { x: 0.5, y: 0.8, w: 10, h: 0.4, fontSize: 14, fontFace: "Calibri" });
  cost.addTable([
    [{ text: "Driver" }, { text: "What to separate" }, { text: "Owner" }],
    [{ text: "Replacement ASP" }, { text: "Average fulfillment cost" }, { text: "Operations" }],
    [{ text: "Mix" }, { text: "Repair, replace, reimburse share" }, { text: "Operations" }],
  ], { x: 0.5, y: 1.4, w: 10, h: 1.6, fontFace: "Calibri" });
  cost.addText(footer, { x: 0.5, y: 6.6, w: 12, h: 0.3, fontSize: 10, fontFace: "Calibri" });

  for (const title of ["Choices", "Evidence"]) {
    const slide = pres.addSlide();
    slide.addText(title, { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 22, fontFace: "Calibri" });
    slide.addText(`${title} keeps its own sentence.`, { x: 0.5, y: 1, w: 10, h: 0.4, fontSize: 16, fontFace: "Calibri" });
    slide.addText(footer, { x: 0.5, y: 6.6, w: 12, h: 0.3, fontSize: 10, fontFace: "Calibri" });
  }

  const read = await readSourceDeck(await pres.write({ outputType: "arraybuffer" }) as ArrayBuffer, "honest.pptx");
  expect("subject is the cover", read.deck.slides[0]?.title, "Faster honest claims.\nTighter fraud.\n$2M without the five-day wait.");
  expect("cover stays empty", read.deck.slides[0]?.context ?? "x", "");
  includes("opening kept", read.deck.slides[1]?.context ?? "", "Risk-tiered claims");
  expect("opening eyebrow", read.deck.slides[1]?.eyebrow ?? "", "DEVICE PROTECTION · 90-DAY PRODUCT PROPOSAL");
  expect("kicker left the body", (read.deck.slides[1]?.context ?? "").includes("DEVICE PROTECTION"), false);
  const rates = (read.deck.slides[1]?.pieces ?? []).filter((piece) => piece.kind === "card" && piece.display.endsWith("%"));
  expect("pilot rates stay", rates.map((piece) => piece.kind === "card" ? `${piece.display} ${piece.label}` : "").join("|"), "80% Retailer pilot SLA in 24h|95% Retailer pilot SLA in 48h");
  expect("green face", cardFace("GREEN\nAuto-approve"), "green");
  expect("amber face", cardFace("AMBER"), "amber");
  expect("red face", cardFace("RED lane"), "red");
  expect("today face", cardFace("Today · one pile"), "ink");
  expect("finance stays plain", cardFace("Finance"), "");
  expect("day sequence", isTimeTitle("DAYS 1–30"), true);
  expect("finance is not a day", isTimeTitle("Finance"), false);
  expect("brand is not the title", read.deck.title.startsWith("Likewize"), false);
  const tableSlide = read.deck.slides.find((slide) => slide.title === "Cost split");
  expect("text table headers", tableSlide?.table?.headers.join("|") ?? "", "Driver|What to separate|Owner");
  expect("text table is not a chart", tableSlide?.chart ?? null, null);
  expect("text table rows", tableSlide?.table?.rows.length ?? 0, 2);
  const joined = read.deck.slides.map((slide) => `${slide.context}\n${(slide.pieces ?? []).map((piece) => piece.kind === "prose" ? piece.text : "").join("\n")}`).join("\n");
  expect("footer removed", joined.includes("Confidential"), false);
}

await honest();

expect("density light", pieceDensity([{ kind: "prose", text: "One clear point." }], "Hello"), "light");
const heavyPoints = Array.from({ length: 24 }, (_, index) => `Keep this written point ${"x".repeat(index + 1)}`);
expect("density over", pieceDensity([{ kind: "list", title: "Plan", points: heavyPoints }], "Too much"), "over");

async function pace() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  pres.layout = "DECK";
  const slide = pres.addSlide();
  slide.addText("One idea", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 28, fontFace: "Calibri" });
  slide.addText("A single sentence stays.", { x: 0.5, y: 1, w: 10, h: 0.4, fontSize: 16, fontFace: "Calibri" });
  const light = await readSourceDeck(await pres.write({ outputType: "arraybuffer" }) as ArrayBuffer, "light.pptx");
  const lightSlide = light.deck.slides[1];
  includes("light note", lightSlide ? slideRead(lightSlide) : "", "Left large");
  expect("light stays one slide", light.deck.slides.length, 2);

  const crowd = new PptxGenJS();
  crowd.defineLayout({ name: "DECK", width: 13.333, height: 7.5 });
  crowd.layout = "DECK";
  const heavy = crowd.addSlide();
  heavy.addText("Too much", { x: 0.5, y: 0.3, w: 10, h: 0.4, fontSize: 28, fontFace: "Calibri" });
  heavy.addText(
    heavyPoints.map((point) => ({ text: `• ${point}`, options: { breakLine: true } })),
    { x: 0.5, y: 1, w: 10, h: 5, fontSize: 16, fontFace: "Calibri" },
  );
  const read = await readSourceDeck(await crowd.write({ outputType: "arraybuffer" }) as ArrayBuffer, "heavy.pptx");
  expect("heavy split", read.deck.slides.length > 2, true);
  const noted = read.deck.slides.find((item) => item.pace?.includes("Split into"));
  includes("split note", noted ? slideRead(noted) : "", "Nothing was rewritten");
  includes("split count", noted ? slideRead(noted) : "", "24 points");
  const kept = read.deck.slides.flatMap((item) => (item.pieces ?? []).flatMap((piece) => (piece.kind === "list" ? piece.points : [])));
  expect("points kept", kept.length, 24);
  expect("first point", kept[0], heavyPoints[0]);
  expect("last point", kept[23], heavyPoints[23]);
  includes("deck split", sourceSummary(read.deck), "split so each one stays readable");
}

await pace();

function blank(title: string, extra: Partial<SourceSlide> = {}): SourceSlide {
  return {
    index: 1,
    title,
    context: "",
    numbers: [],
    chart: null,
    extraCharts: 0,
    picture: false,
    images: [],
    table: null,
    pieces: [],
    ...extra,
  };
}

function edits() {
  const title = blank("Cover", { context: "Ada Lovelace" });
  const one = blank("One", { pieces: [{ kind: "list", title: "", points: ["Alpha", "Beta"] }] });
  const two = blank("Two", { pieces: [{ kind: "list", title: "", points: ["Gamma"] }] });
  const deck = [title, one, two].map((slide, position) => ({ ...slide, index: position + 1 }));
  const stayed = moveSlide(deck, 1, 1);
  includes("title stays", stayed.note, "title slide stays first");
  expect("title unmoved", stayed.slides[0]?.title, "Cover");
  const moved = moveSlide(deck, 3, -1);
  expect("moved title", moved.slides[1]?.title, "Two");
  expect("moved neighbor", moved.slides[2]?.title, "One");
  const copied = duplicateSlide(deck, 2);
  expect("copy length", copied.slides.length, 4);
  expect("copy title", copied.slides[2]?.title, "One");
  const removed = deleteSlide(deck, 2);
  expect("deleted length", removed.slides.length, 2);
  expect("deleted kept", removed.slides[1]?.title, "Two");
  includes("delete title", deleteSlide(deck, 1).note, "title slide stays");
  const merged = mergeSlides(deck, 2);
  includes("merge note", merged.note, "Nothing was rewritten");
  const points = (merged.slides[1]?.pieces ?? []).flatMap((piece) => (piece.kind === "list" ? piece.points : []));
  expect("merged points", points.includes("Alpha") && points.includes("Gamma"), true);
  expect("merged title", merged.slides[1]?.title, "One");
  const chart = { kind: "columns" as const, drawn: "", categories: ["Apr", "May"], series: [{ name: "Count", values: [3, 5] }] };
  const table = { headers: ["", "Count"], rows: [["Apr", "3"], ["May", "~5"]] };
  const left = blank("Bars", { chart, table });
  const right = blank("More", { chart: { ...chart, categories: ["Jun"], series: [{ name: "Count", values: [8] }] } });
  const pair = [title, left, right].map((slide, position) => ({ ...slide, index: position + 1 }));
  const apart = mergeSlides(pair, 2);
  includes("two charts", apart.note, "Two charts stay");
  expect("charts remain", apart.slides.length, 3);
  expect("can flip", canFlip(left), true);
  expect("show both", slideShow(left), "both");
  expect("shown table", showSlide(pair, 2, "table").slides[1]?.show, "table");
  expect("tilde stays", tableFromSlide(left)?.rows[1]?.[1], "~5");
  const picture = { mime: "image/png" as const, dataUrl: `data:${PNG}` };
  const added = addSlidePictures(
    [title, blank("Pictures")].map((slide, position) => ({ ...slide, index: position + 1 })),
    2,
    Array.from({ length: 5 }, () => picture),
  );
  expect("pictures split", added.slides.length > 2, true);
  includes("pictures stay", added.note, "pictures stay");
  const full = blank("Full", {
    pieces: [{ kind: "list", title: "Plan", points: Array.from({ length: 24 }, (_, index) => `Keep this written point ${"x".repeat(index + 1)}`) }],
  });
  const followed = addSlidePictures([title, full].map((slide, position) => ({ ...slide, index: position + 1 })), 2, [picture]);
  includes("picture continues", followed.note, "continues");
  expect("words stay", (followed.slides[1]?.pieces ?? []).some((piece) => piece.kind === "list" && piece.points.length === 24), true);
  expect("picture slide", followed.slides[2]?.images.length, 1);
  expect("png size", imageSize(`data:${PNG}`)?.width, 1);
  const shot = (width: number, height: number) => ({ dataUrl: pngBox(width, height) });
  expect("one picture", pictureLayout([shot(800, 500)]), "solo");
  expect("wide strips stack", pictureLayout([shot(800, 80), shot(780, 90), shot(760, 90)]), "stack");
  expect("two shots sit together", pictureLayout([shot(900, 500), shot(700, 500)]), "pair");
  expect("tall shots sit in a row", pictureLayout([shot(300, 700), shot(280, 640), shot(360, 500)]), "row");
}

function pngBox(width: number, height: number): string {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52]);
  bytes[16] = (width >>> 24) & 255;
  bytes[17] = (width >>> 16) & 255;
  bytes[18] = (width >>> 8) & 255;
  bytes[19] = width & 255;
  bytes[20] = (height >>> 24) & 255;
  bytes[21] = (height >>> 16) & 255;
  bytes[22] = (height >>> 8) & 255;
  bytes[23] = height & 255;
  return `data:image/png;base64,${btoa(String.fromCharCode(...bytes))}`;
}

edits();

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log("source deck checks passed");
