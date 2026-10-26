/** A client scorecard. Every figure on the slides comes from this book. */

export type ListeningRow = {
  label: string;
  count: number;
  positive: number;
  neutral: number;
  negative: number;
};

export type JourneyStage = {
  label: string;
  mentions: number;
  /** Percents of this stage. They should sum to 100. */
  positive: number;
  neutral: number;
  negative: number;
};

export type MonthStack = {
  label: string;
  fulfilled: number;
  open: number;
  notApproved: number;
};

/** How a stated rate compares with the figure in the brief. */
export type ProofRelation = "above" | "below" | "least" | "exact";

export type Proof = {
  label: string;
  relation: ProofRelation;
  /** The percent written in the brief. Not a measured value beyond that figure. */
  value: number;
};

/** A results story. Figures are only what the brief stated. */
export type Story = {
  client: string;
  period: string;
  verb: string;
  unit: string;
  volume: number | null;
  proofs: Proof[];
  /** "about" when the brief said about, around, or roughly. */
  hedge?: string;
};

/** One stated figure. A rate on a volume is the rate, not a share of that volume. */
export type Fact = {
  kind: "volume" | "rate" | "count";
  label: string;
  value: number;
  relation: ProofRelation;
  unit: string;
  hedge: string;
};

/** Voice, chat, or one unnamed result. Counts are not added across groups. */
export type FactGroup = {
  name: string;
  volume: Fact | null;
  rates: Fact[];
  counts: Fact[];
};

/** Figures read from a messy brief, before a slide is chosen. */
export type FactRead = {
  client: string;
  period: string;
  verb: string;
  groups: FactGroup[];
};

/** One month, or one period, inside a series. */
export type TrackPoint = {
  label: string;
  total: number;
  part: number;
  /** Stated percent when the source gave one, otherwise part divided by total. */
  rate: number;
};

/** A named series, such as voice calls or chats, with only the rows that were supplied. */
export type Track = {
  name: string;
  unit: string;
  partLabel: string;
  points: TrackPoint[];
};

export type ScopeCard = {
  label: string;
  text: string;
};

/** A figure written on a source slide. `display` is the text as written. */
export type SourceNumber = {
  label: string;
  value: number;
  display: string;
  /** The sentence under a callout, as written on the slide. */
  note?: string;
};

/** One block from the source slide, in the order it was written. */
export type SlidePiece =
  | { kind: "prose"; text: string }
  | { kind: "card"; title: string; detail: string; display: string; label: string }
  | { kind: "list"; title: string; points: string[] };

export type SourceSeries = {
  name: string;
  values: number[];
};

/** A chart taken from the file. Nothing here is estimated. */
export type SourceChart = {
  kind: "columns" | "bars" | "line" | "donut" | "stacked" | "funnel";
  /** Set when the drawing differs from the file, such as a pie drawn as a donut. */
  drawn: string;
  categories: string[];
  series: SourceSeries[];
};

/** A table as written on the slide. Cells stay text, including a tilde. */
export type SourceTable = {
  headers: string[];
  rows: string[][];
};

/** A picture kept from the file. A chart that is only an image stays an image. */
export type SourceImage = {
  mime: "image/png" | "image/jpeg" | "image/gif";
  dataUrl: string;
};

/** One slide of an uploaded PowerPoint, read before any new slide is chosen. */
export type SourceSlide = {
  /** 1-based position in the file. */
  index: number;
  /** 1-based slide in the uploaded file. Stays put when one slide becomes two. */
  from?: number;
  title: string;
  /** A short section label above the title, such as DIAGNOSIS. Empty when the title already says it. */
  eyebrow?: string;
  context: string;
  numbers: SourceNumber[];
  chart: SourceChart | null;
  /** Other charts on the same slide, left undrawn. */
  extraCharts: number;
  /** True when the slide is a picture and almost no text was stored. */
  picture: boolean;
  /** Every raster picture on the slide, including a bar chart saved as an image. */
  images: SourceImage[];
  /** The number table on the slide, kept beside any chart drawn from it. */
  table: SourceTable | null;
  /** Cards and lists in source order. Empty when the slide is ordinary prose. */
  pieces?: SlidePiece[];
  /** Why this slide was kept large, kept at reading size, or continued. */
  pace?: string;
  /** Chart, table, or both. Empty means both when the file has both. */
  show?: "chart" | "table" | "both";
  /** A layout the director chose. Empty means the code chooses. */
  direction?: "poster" | "proof" | "versus" | "hero" | "path" | "metrics" | "cascade" | "stack";
};

/** A deck read slide by slide. */
export type SourceDeck = {
  title: string;
  /** What the file contained: how many slides, charts, and figures, and where it leads. */
  intent: string;
  slides: SourceSlide[];
};

/** One placeholder slide. The hint is a prompt to fill, not a result. */
export type OutlineCard = {
  kicker: string;
  title: string;
  hint: string;
};

/** A figure-free deck. The shape follows the ask, so a journey is not a scorecard. */
export type Outline = {
  client: string;
  subject: string;
  intent: string;
  shape: "journey" | "points";
  cards: OutlineCard[];
};

/** What a brief, a sheet, or a deck was actually about, before any slide is chosen. */
export type Reading = {
  title: string;
  period: string;
  /** What the figures are. */
  subject: string;
  /** What the deck is for. */
  intent: string;
  scope: ScopeCard[];
  /** A next step, when the source asked for one. Empty when the close was only a definition. */
  ask: string;
  tracks: Track[];
};

export type Book = {
  client: string;
  market: string;
  period: string;
  active: number;
  claims: number;
  fulfilled: number;
  approved: number;
  peerName: string;
  peerActive: number;
  peerClaims: number;
  peerFulfilled: number;
  brief: string;
  recommendation: string;
  /** File name of a decoded deck or document. */
  sourceName: string;
  /** Text taken from that file. Empty when nothing was decoded. */
  sourceText: string;
  /** Empty unless a Listening sheet was read. */
  listening: ListeningRow[];
  /** Empty unless a Journey sheet was read. */
  journey: JourneyStage[];
  /** Empty unless a Months sheet was read. */
  months: MonthStack[];
  /** Set when the brief is a results story rather than a claims book. */
  story: Story | null;
  /** Set when the source is a series, a scope, or both, rather than a claims book. */
  reading: Reading | null;
  /** Set when an uploaded PowerPoint was read slide by slide. */
  sourceDeck: SourceDeck | null;
  /** Set when the brief states more than one series, such as voice and chat. */
  facts: FactRead | null;
  /** A plain number sheet. Rows are charted as written and are not added together. */
  grid: SheetGrid | null;
  /** A figure-free deck of placeholder slides. Null when the brief already has figures. */
  outline: Outline | null;
  /** Data URL for a known client mark. Empty when the client has no approved logo. */
  logo: string;
};

/** One numeric column. Blank cells stay null and are never treated as zero. */
export type SheetColumn = {
  name: string;
  values: (number | null)[];
  /** True when the header is a rate, or the values are a 0 to 100 rate rather than a count. */
  percent: boolean;
};

/** A plain number sheet. Nothing here is a sum of the rows. */
export type SheetGrid = {
  title: string;
  labels: string[];
  columns: SheetColumn[];
  cadence: "month" | "year" | "rows";
  /** Next period label, such as "Nov-26 proj". Null when the sheet is not month or year. */
  projectionLabel: string | null;
  /** One step ahead, parallel to columns. Null when that column has fewer than 3 points. */
  projection: (number | null)[];
  note: string;
};

/** Illustrative topics. Not a real client. Parts sum to the count. */
export const SAMPLE_LISTENING: ListeningRow[] = [
  { label: "Claims desk", count: 820, negative: 410, neutral: 246, positive: 164 },
  { label: "Repair", count: 540, negative: 162, neutral: 216, positive: 162 },
  { label: "Trade-in", count: 310, negative: 62, neutral: 93, positive: 155 },
  { label: "App", count: 180, negative: 36, neutral: 54, positive: 90 },
];

/** Illustrative stages. Percents are the sample split, not computed from the book. */
export const SAMPLE_JOURNEY: JourneyStage[] = [
  { label: "File a claim", mentions: 420, negative: 22, neutral: 18, positive: 60 },
  { label: "Approve", mentions: 380, negative: 14, neutral: 20, positive: 66 },
  { label: "Fulfill", mentions: 300, negative: 28, neutral: 16, positive: 56 },
  { label: "Close", mentions: 260, negative: 8, neutral: 12, positive: 80 },
];

/** Illustrative months. Each column is that month only. */
export const SAMPLE_MONTHS: MonthStack[] = [
  { label: "Apr", fulfilled: 1200, open: 400, notApproved: 300 },
  { label: "May", fulfilled: 1400, open: 480, notApproved: 280 },
  { label: "Jun", fulfilled: 1500, open: 520, notApproved: 310 },
  { label: "Jul", fulfilled: 1600, open: 600, notApproved: 340 },
  { label: "Aug", fulfilled: 1700, open: 640, notApproved: 360 },
  { label: "Sep", fulfilled: 1800, open: 700, notApproved: 380 },
];

export const SAMPLE_BOOK: Book = {
  client: "Northline Care",
  market: "US",
  period: "Year to date",
  active: 48200,
  claims: 19280,
  fulfilled: 9640,
  approved: 14460,
  peerName: "Harbor Cover",
  peerActive: 51000,
  peerClaims: 12750,
  peerFulfilled: 6120,
  brief:
    "Client scorecard for Northline Care. Lead with the claim rate, show how the book splits, then set it next to Harbor Cover.",
  recommendation:
    "The book claims at 40% of active subscribers, and 10 points of that are approved and still open. Clear that queue before adding acquisition.",
  sourceName: "",
  sourceText: "",
  listening: SAMPLE_LISTENING,
  journey: SAMPLE_JOURNEY,
  months: SAMPLE_MONTHS,
  story: null,
  reading: null,
  sourceDeck: null,
  facts: null,
  grid: null,
  outline: null,
  logo: "",
};

export type Tone = "fulfilled" | "approved" | "open" | "rest";

export type Slice = {
  key: string;
  label: string;
  count: number;
  /** Share of active subscribers, already rounded so the claim parts sum to the headline rate. */
  percent: number;
  tone: Tone;
};

export type BookView = {
  active: number;
  claims: number;
  fulfilled: number;
  approvedOpen: number;
  notApproved: number;
  noClaim: number;
  /** Claims divided by active subscribers, rounded. Can exceed 100. */
  rate: number;
  over: boolean;
  fulfilledShare: number;
  openShare: number;
  notShare: number;
  noClaimShare: number;
  /** Arcs of the ring. Zero slices are already omitted. */
  slices: Slice[];
  /** True when the ring is the claim mix because claims exceed the active book. */
  ringIsClaimMix: boolean;
  warnings: string[];
  peerRate: number;
  peerFulfilledRate: number;
};

const clampCount = (value: number) => (Number.isFinite(value) ? Math.max(0, Math.round(value)) : 0);

/**
 * Largest-remainder rounding of each count as a percent of `whole`.
 * The rounded percents sum to the rounded total, which is the headline rate
 * when the counts are the claim parts and `whole` is the active book.
 * The total may exceed 100 when claims exceed the active book.
 */
export function largestRemainder(counts: number[], whole: number): number[] {
  if (counts.length === 0) return [];
  if (whole <= 0) return counts.map(() => 0);
  const exact = counts.map((count) => (count / whole) * 100);
  const target = Math.round(exact.reduce((sum, value) => sum + value, 0));
  if (target <= 0) return counts.map(() => 0);
  const floors = exact.map((value) => Math.floor(value));
  let left = target - floors.reduce((sum, value) => sum + value, 0);
  const order = exact
    .map((value, index) => ({ index, frac: value - Math.floor(value) }))
    .sort((a, b) => b.frac - a.frac || a.index - b.index);
  const out = [...floors];
  let guard = 0;
  while (left > 0 && guard < counts.length * 4) {
    out[order[guard % order.length]!.index] += 1;
    left -= 1;
    guard += 1;
  }
  return out;
}

export function rateOf(part: number, whole: number): number {
  if (whole <= 0 || part <= 0) return 0;
  return Math.round((part / whole) * 100);
}

export function viewBook(book: Book): BookView {
  const active = clampCount(book.active);
  const claims = clampCount(book.claims);
  const fulfilled = clampCount(book.fulfilled);
  const approved = clampCount(book.approved);
  const warnings: string[] = [];

  if (fulfilled > approved) {
    warnings.push(
      "Fulfilled is higher than approved. Approved and still open stays at zero until approved includes fulfilled.",
    );
  }
  if (approved > claims) {
    warnings.push(
      "Approved is higher than claims. Not approved stays at zero until claims include every approved claim.",
    );
  }
  if (active === 0) {
    warnings.push("Add active subscribers. The rate is claims divided by that book.");
  } else if (claims > active) {
    warnings.push(
      "Claims are higher than the active book. The rate can exceed 100%, and the bar stops at the end of the track.",
    );
  }

  const approvedOpen = Math.max(0, approved - fulfilled);
  const notApproved = Math.max(0, claims - approved);
  const noClaim = Math.max(0, active - claims);
  const rate = rateOf(claims, active);
  const over = rate > 100;

  const [fulfilledShare, openShare, notShare] = largestRemainder(
    [fulfilled, approvedOpen, notApproved],
    active,
  ) as [number, number, number];
  const noClaimShare = over ? 0 : Math.max(0, 100 - rate);

  const ringIsClaimMix = over;
  const slices: Slice[] = [
    { key: "fulfilled", label: "Fulfilled", count: fulfilled, percent: fulfilledShare, tone: "fulfilled" },
    {
      key: "open",
      label: "Approved, still open",
      count: approvedOpen,
      percent: openShare,
      tone: "approved",
    },
    { key: "not", label: "Not approved", count: notApproved, percent: notShare, tone: "open" },
  ];
  if (!ringIsClaimMix) {
    slices.push({
      key: "rest",
      label: "No claim",
      count: noClaim,
      percent: noClaimShare,
      tone: "rest",
    });
  }

  const peerActive = clampCount(book.peerActive);
  return {
    active,
    claims,
    fulfilled,
    approvedOpen,
    notApproved,
    noClaim,
    rate,
    over,
    fulfilledShare,
    openShare,
    notShare,
    noClaimShare,
    slices: slices.filter((slice) => slice.count > 0),
    ringIsClaimMix,
    warnings,
    peerRate: rateOf(book.peerClaims, peerActive),
    peerFulfilledRate: rateOf(book.peerFulfilled, peerActive),
  };
}

export type Conversation = {
  total: number;
  positive: number;
  neutral: number;
  negative: number;
  /** Positive plus neutral, as a percent of mentions. */
  health: number;
  positiveShare: number;
  neutralShare: number;
  negativeShare: number;
};

/** Health is the share of mentions that are not negative. Shares use largest remainder and sum to 100. */
export function conversationOf(rows: ListeningRow[]): Conversation {
  const positive = rows.reduce((sum, row) => sum + clampCount(row.positive), 0);
  const neutral = rows.reduce((sum, row) => sum + clampCount(row.neutral), 0);
  const negative = rows.reduce((sum, row) => sum + clampCount(row.negative), 0);
  const counted = rows.reduce((sum, row) => sum + clampCount(row.count), 0);
  const parts = positive + neutral + negative;
  const total = counted || parts;
  const base = parts || total;
  const [negativeShare, neutralShare, positiveShare] = largestRemainder(
    [negative, neutral, positive],
    base,
  ) as [number, number, number];
  return {
    total,
    positive,
    neutral,
    negative,
    health: base ? Math.round(((positive + neutral) / base) * 100) : 0,
    positiveShare: positiveShare ?? 0,
    neutralShare: neutralShare ?? 0,
    negativeShare: negativeShare ?? 0,
  };
}

export function formatCount(value: number): string {
  return clampCount(value).toLocaleString("en-US");
}

/** A decoded document earns its own slide when it is not already the cover line. */
export function hasBroughtIn(book: Book): boolean {
  const source = book.sourceText.trim();
  if (!source) return false;
  if (source === book.brief.trim() && source.length <= 220) return false;
  return true;
}

export function slideBrief(brief: string, max = 220): string {
  const clean = brief.replace(/\s+/g, " ").trim();
  if (clean.length <= max) return clean;
  return `${clean.slice(0, max - 1).trimEnd()}…`;
}
