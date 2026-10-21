import type { Book, Outline, OutlineCard } from "./book";
import { decodeBrief, isAudienceLine } from "./brief";
import { isFacts, readFacts } from "./facts";
import { isGrid } from "./grid";
import { isSeries, readMaterial } from "./reading";
import { isSource } from "./sourceDeck";
import { isResults, readStory } from "./story";

type Step = readonly [string, string];

type Journey = {
  name: string;
  match: RegExp;
  steps: readonly Step[];
};

/** Stage names follow the product journeys. The lines are prompts to fill, not results. */
const JOURNEYS: readonly Journey[] = [
  {
    name: "Trade-in",
    match: /\b(trade[-\s]?in|buyback)\b/i,
    steps: [
      ["Get a device value", "Point to add. Say where the customer sees a value, and that a quote can be saved."],
      ["Confirm the device", "Point to add. Say how the device, the condition, and eligibility are checked before the offer is locked."],
      ["Accept and return", "Point to add. Say which return path this program uses: ship, drop-off, or store."],
      ["Inspect the device", "Point to add. Say what is checked against the declared condition."],
      ["Confirm the final value", "Point to add. Say what happens when inspection changes the value, and who accepts it."],
      ["Pay and recover", "Point to add. Say how the customer is paid and where the device goes next."],
    ],
  },
  {
    name: "Upgrade",
    match: /\bupgrades?\b/i,
    steps: [
      ["Enroll and protect", "Point to add. Say how the customer enters the offer."],
      ["Support the current device", "Point to add. Say what support exists while they still have this device."],
      ["Check upgrade eligibility", "Point to add. Say which rules decide whether an upgrade is allowed."],
      ["Assess the device", "Point to add. Say how the current device is assessed."],
      ["Hand back and return", "Point to add. Say how the device comes back."],
      ["Move to the next device", "Point to add. Say what the customer leaves with."],
    ],
  },
  {
    name: "Warranty",
    match: /\b(extended warranty|warrant(?:y|ies))\b/i,
    steps: [
      ["Activate the warranty", "Point to add. Say how the warranty starts."],
      ["Access service", "Point to add. Say where the customer goes for help."],
      ["Report the issue", "Point to add. Say what the customer is asked to describe."],
      ["Check coverage", "Point to add. Say how coverage is confirmed."],
      ["Repair, replace, or reimburse", "Point to add. Say which outcomes this program offers."],
      ["Complete the case", "Point to add. Say how the customer knows it is finished."],
    ],
  },
  {
    name: "Home",
    match: /\b(home tech|home gadget|gadget protection)\b/i,
    steps: [
      ["Choose or receive the plan", "Point to add. Say how the plan is offered."],
      ["Register eligible products", "Point to add. Say which products can be covered."],
      ["Access support", "Point to add. Say where the customer starts a request."],
      ["Report the issue", "Point to add. Say what has to be described."],
      ["Repair, replace, or resolve", "Point to add. Say how the product is put right."],
      ["Manage the ongoing plan", "Point to add. Say what continues after the case."],
    ],
  },
  {
    name: "Account benefit",
    match: /\b(embedded finance|account benefit|card benefit)\b/i,
    steps: [
      ["Receive or select the benefit", "Point to add. Say how the benefit is offered."],
      ["Confirm eligibility", "Point to add. Say what must be true before it starts."],
      ["Access the branded service", "Point to add. Say where the customer uses it."],
      ["Submit a claim", "Point to add. Say how a claim is opened."],
      ["Resolve the claim", "Point to add. Say how it is closed."],
      ["Reconcile and report", "Point to add. Say what the partner needs to see afterward."],
    ],
  },
  {
    name: "Protection",
    match: /\b(device protection|protection plan|mobile protection|claims journey|protection)\b/i,
    steps: [
      ["Enroll the device", "Point to add. Say how a device enters the plan."],
      ["Access support", "Point to add. Say where the customer gets help."],
      ["Report what happened", "Point to add. Say what the claim needs to capture."],
      ["Review the claim", "Point to add. Say how coverage and eligibility are checked."],
      ["Repair, replace, or resolve", "Point to add. Say which outcomes are on the table."],
      ["Close and report", "Point to add. Say how the case ends and what is reported."],
    ],
  },
];

function hasFigures(text: string): boolean {
  if (readMaterial(text) || readFacts(text) || readStory(text)) return true;
  const brief = decodeBrief(text);
  if ((brief.active ?? 0) > 0 || (brief.claims ?? 0) > 0 || (brief.fulfilled ?? 0) > 0 || (brief.approved ?? 0) > 0) return true;
  if ((brief.listening?.length ?? 0) > 0 || (brief.journey?.length ?? 0) > 0 || (brief.months?.length ?? 0) > 0) return true;
  return false;
}

function clientOf(text: string): string {
  const named = decodeBrief(text).client?.trim() ?? "";
  if (!named || isAudienceLine(named)) return "";
  return named;
}

function journeyOf(text: string): Journey | null {
  let best: { at: number; journey: Journey } | null = null;
  for (const journey of JOURNEYS) {
    const found = journey.match.exec(text);
    if (!found || found.index == null) continue;
    if (!best || found.index < best.at) best = { at: found.index, journey };
  }
  return best?.journey ?? null;
}

function listedStages(text: string): string[] {
  const match = text.match(/\b(?:stages?|steps?)\s*:\s*([^\n.]{3,200})/i) ?? text.match(/\bjourney\s+of\s+([^\n.]{3,200})/i);
  if (!match?.[1]) return [];
  return match[1]
    .split(/,|\/|\band\b/i)
    .map((part) => part.trim().replace(/[.:;]+$/, ""))
    .filter((part) => part.length > 1 && part.length < 48)
    .slice(0, 8);
}

function subjectOf(text: string, journey: Journey | null): string {
  if (journey) return journey.name;
  const match = text.match(/\b(?:about|on|covering)\s+(?:a|an|the|our)?\s*([A-Za-z0-9][^.\n]{2,60})/i);
  const raw = match?.[1]?.split(/\s+(?:where|that|and|with|so)\b/i)[0]?.trim().replace(/[.,;:]+$/, "") ?? "";
  if (!raw) return "";
  return `${raw.charAt(0).toUpperCase()}${raw.slice(1)}`.slice(0, 48);
}

function cardsFrom(steps: readonly Step[], kicker: string): OutlineCard[] {
  return steps.map(([title, hint], index) => ({
    kicker: `${kicker} ${index + 1}`,
    title,
    hint,
  }));
}

function openJourney(subject: string): OutlineCard[] {
  const name = subject || "this journey";
  return [
    { kicker: "Step 1", title: `Open ${name}`, hint: "Point to add. Say what the customer meets first." },
    { kicker: "Step 2", title: "The choice", hint: "Point to add. Say what the customer decides." },
    { kicker: "Step 3", title: "The handoff", hint: "Point to add. Say where the work goes next." },
    { kicker: "Step 4", title: "What comes back", hint: "Point to add. Say what the customer hears at the end." },
  ];
}

function pointCards(subject: string): OutlineCard[] {
  const name = subject || "this deck";
  return [
    { kicker: "Point 1", title: "Why this matters", hint: `Point to add. Say why ${name} belongs in the room.` },
    { kicker: "Point 2", title: "What to show", hint: "Point to add. Leave the proof blank until a figure is real." },
    { kicker: "Point 3", title: "What to leave out", hint: "Point to add. Say what this deck should not claim." },
    { kicker: "Point 4", title: "The ask", hint: "Point to add. Say the one decision this deck should earn." },
  ];
}

/** A deck request with no figures. Known products keep their own stages. */
export function readOutline(text: string): Outline | null {
  const clean = text.replace(/\r/g, "").trim();
  if (clean.length < 24 || hasFigures(clean)) return null;
  if (/\b(scorecard|claim rate|active subscribers)\b/i.test(clean) && !/\bjourney\b/i.test(clean)) return null;
  const journeyWord = /\bjourney\b/i.test(clean);
  const asksDeck = /\b(deck|presentation|slides?)\b/i.test(clean);
  if (!asksDeck && !journeyWord) return null;
  const journey = journeyOf(clean);
  const listed = listedStages(clean);
  const subject = subjectOf(clean, listed.length ? null : journey);
  const client = clientOf(clean);
  if (!client && !journey && !subject && listed.length === 0) return null;
  const asksPoints = /\b(important point|points to mention|what to mention|needs to be mentioned)\b/i.test(clean);
  const named = subject || journey?.name || "This deck";
  const intent = asksPoints
    ? `The points to mention${journey ? ` on the ${journey.name.toLowerCase()} journey` : ""}.`
    : journeyWord
      ? `${named} journey. Add the point to mention at each step.`
      : `Placeholder slides for ${named}. Add the point to mention.`;
  let cards: OutlineCard[];
  let shape: Outline["shape"];
  if (listed.length > 0) {
    shape = "journey";
    cards = listed.map((title, index) => ({
      kicker: `Step ${index + 1}`,
      title,
      hint: "Point to add. Say what needs to be mentioned at this step.",
    }));
  } else if (journey) {
    shape = "journey";
    cards = cardsFrom(journey.steps, "Step");
  } else if (journeyWord) {
    shape = "journey";
    cards = openJourney(subject);
  } else {
    shape = "points";
    cards = pointCards(subject);
  }
  return { client, subject: subject || journey?.name || "", intent, shape, cards };
}

export function isOutline(book: Book): boolean {
  if (!book.outline?.cards.length) return false;
  if (isSource(book) || isSeries(book) || isGrid(book) || isFacts(book) || isResults(book)) return false;
  if (book.active > 0 || book.claims > 0) return false;
  return true;
}

export function outlineLine(outline: Outline): string {
  const subject = outline.subject || outline.client || "This deck";
  return `${subject}. ${outline.intent}`.replace(/\s+/g, " ").trim();
}

export function outlineSummary(outline: Outline): string {
  const who = outline.client || "This deck";
  const count = outline.cards.length;
  return `${who}: ${outline.intent} ${count} placeholder ${count === 1 ? "slide" : "slides"}. No figure was in the brief.`;
}
