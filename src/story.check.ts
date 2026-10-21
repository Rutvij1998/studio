import { emptyBook } from "./parse.ts";
import { isResults, readStory, storyLine, storySummary } from "./story.ts";

const failures: string[] = [];
const expect = (label: string, actual: unknown, wanted: unknown) => {
  if (actual !== wanted) failures.push(`${label}: ${String(actual)} !== ${String(wanted)}`);
};

const prompt =
  "Create a presentation showing that we handled 1.9M IVA Calls for Boost Mobile in this year where the containment rate is above 70% and Authentication is above 90%.";
const story = readStory(prompt);
if (!story) {
  failures.push("story was not read");
} else {
  expect("client", story.client, "Boost Mobile");
  expect("period", story.period, "This year");
  expect("verb", story.verb, "handled");
  expect("unit", story.unit, "IVA calls");
  expect("volume", story.volume, 1900000);
  expect("proofs", story.proofs.length, 2);
  expect("containment", story.proofs[0]?.label, "Containment");
  expect("containment relation", story.proofs[0]?.relation, "above");
  expect("containment value", story.proofs[0]?.value, 70);
  expect("auth", story.proofs[1]?.label, "Authentication");
  expect("auth value", story.proofs[1]?.value, 90);
  expect(
    "line",
    storyLine(story),
    "1.9 million IVA calls handled this year. Containment is above 70%. Authentication is above 90%.",
  );
  const book = { ...emptyBook(), client: story.client, period: story.period, brief: prompt, story };
  expect("results", isResults(book), true);
  expect("summary client", storySummary(story).includes("Boost Mobile"), true);
}

const exact = readStory("Presentation for Pinetop Shield. We served 12,400 repairs in 2025. Approval is 88%.");
expect("exact client", exact?.client, "Pinetop Shield");
expect("exact volume", exact?.volume, 12400);
expect("exact unit", exact?.unit, "Repairs");
expect("exact period", exact?.period, "2025");
expect("exact relation", exact?.proofs[0]?.relation, "exact");
expect("exact value", exact?.proofs[0]?.value, 88);

const scorecard = readStory("Scorecard for Northline Care. 48,200 active subscribers and 19,280 claims.");
expect("scorecard stays a book", scorecard, null);

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log("story checks passed");
