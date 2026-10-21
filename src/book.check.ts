import { SAMPLE_BOOK, conversationOf, largestRemainder, viewBook } from "./book.ts";

const sample = viewBook(SAMPLE_BOOK);
const failures: string[] = [];
const expect = (label: string, actual: unknown, wanted: unknown) => {
  if (actual !== wanted) failures.push(`${label}: ${String(actual)} !== ${String(wanted)}`);
};

expect("rate", sample.rate, 40);
expect("fulfilled", sample.fulfilledShare, 20);
expect("open", sample.openShare, 10);
expect("not", sample.notShare, 10);
expect("sum", sample.fulfilledShare + sample.openShare + sample.notShare, 40);
expect("quiet", sample.noClaimShare, 60);
expect("peer", sample.peerRate, 25);
expect("peer fulfilled", sample.peerFulfilledRate, 12);
expect("slices", sample.slices.length, 4);
expect("ring", sample.ringIsClaimMix, false);

const over = viewBook({ ...SAMPLE_BOOK, active: 1000, claims: 1500, fulfilled: 400, approved: 900 });
expect("over rate", over.rate, 150);
expect("over sum", over.fulfilledShare + over.openShare + over.notShare, 150);
expect("over ring", over.ringIsClaimMix, true);
expect("over slices", over.slices.some((slice) => slice.key === "rest"), false);

const rounded = largestRemainder([1, 1, 1], 3);
expect("thirds", rounded.join(","), "34,33,33");

const talk = conversationOf(SAMPLE_BOOK.listening);
expect("topics", SAMPLE_BOOK.listening.length, 4);
expect("stages", SAMPLE_BOOK.journey.length, 4);
expect("months", SAMPLE_BOOK.months.length, 6);
expect("mentions", talk.total, 1850);
expect("health", talk.health, 64);
expect("neg share", talk.negativeShare, 36);
expect("neu share", talk.neutralShare, 33);
expect("pos share", talk.positiveShare, 31);
for (const stage of SAMPLE_BOOK.journey) {
  expect(stage.label, stage.negative + stage.neutral + stage.positive, 100);
}

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log("book checks passed");
