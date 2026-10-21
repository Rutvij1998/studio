import { emptyBook } from "./parse.ts";
import { applyBrief, briefSummary, coverCopy, decodeBrief } from "./brief.ts";

const failures: string[] = [];
const expect = (label: string, actual: unknown, wanted: unknown) => {
  if (actual !== wanted) failures.push(`${label}: ${String(actual)} !== ${String(wanted)}`);
};

const full = decodeBrief(`Scorecard for Northline Care.
Market: US. Period: Year to date.
48,200 active subscribers, 19,280 claims, 9,640 fulfilled, 14,460 approved.
Versus Harbor Cover, 51,000 active subscribers, 12,750 claims, 6,120 fulfilled.
Show the mix as columns. Large type.
Recommendation: Clear the open queue before adding acquisition.
Listening
Claims desk: 410 negative, 246 neutral, 164 positive
Repair: 162 negative, 216 neutral, 162 positive`);

expect("client", full.client, "Northline Care");
expect("market", full.market, "US");
expect("period", full.period, "Year to date");
expect("active", full.active, 48200);
expect("claims", full.claims, 19280);
expect("fulfilled", full.fulfilled, 9640);
expect("approved", full.approved, 14460);
expect("peer", full.peerName, "Harbor Cover");
expect("peer active", full.peerActive, 51000);
expect("peer claims", full.peerClaims, 12750);
expect("peer fulfilled", full.peerFulfilled, 6120);
expect("mix", full.mix, "columns");
expect("type", full.type, "l");
expect("topics", full.listening?.length, 2);
expect("desk", full.listening?.[0]?.count, 820);
expect("rec", full.recommendation, "Clear the open queue before adding acquisition.");
expect(
  "cover drops controls",
  coverCopy("Scorecard for Pinetop Shield. Use a meter. Compact type.\nRecommendation: Clear the queue before adding acquisition."),
  "Scorecard for Pinetop Shield.",
);
expect(
  "cover drops series",
  coverCopy("Scorecard for Northline Care.\n48,200 active subscribers.\nListening\nClaims desk: 410 negative, 246 neutral, 164 positive"),
  "Scorecard for Northline Care. 48,200 active subscribers.",
);

const rated = decodeBrief("10,000 active subscribers and a 25% claim rate. 1,000 fulfilled. 1,800 approved. Use a meter. Compact type.");
const ratedBook = applyBrief(emptyBook(), rated, { counts: "fill", words: "fill" });
expect("derived claims", ratedBook.claims, 2500);
expect("derived fulfilled", ratedBook.fulfilled, 1000);
expect("meter", rated.mix, "meter");
expect("compact", rated.type, "s");

const sheet = applyBrief(
  { ...emptyBook(), active: 10000, claims: 1000, client: "Pinetop Shield" },
  full,
  { counts: "empty", words: "fill" },
);
expect("sheet claims stay", sheet.claims, 1000);
expect("sheet active stays", sheet.active, 10000);
expect("prompt names the client", sheet.client, "Northline Care");
expect("prompt fills listening", sheet.listening.length, 2);

const months = decodeBrief(`Show the claims by month.
Apr: 1200 fulfilled, 400 open, 300 not approved
May: 1400 fulfilled, 480 still open, 280 not approved`);
expect("month count", months.months?.length, 2);
expect("apr open", months.months?.[0]?.open, 400);
expect("may not", months.months?.[1]?.notApproved, 280);

const ask = decodeBrief("Show sentiment and the journey over time, versus a peer.");
expect("no topics", ask.listening, null);
expect("asked listening", ask.askedListening, true);
expect("asked journey", ask.askedJourney, true);
expect("asked months", ask.askedMonths, true);
const askNote = briefSummary({ countsFrom: "none", prompt: ask, document: null, book: emptyBook() });
expect("note mentions listening", askNote.includes("Listening was asked for"), true);
expect("note mentions months", askNote.includes("Months were asked for"), true);

const columns = decodeBrief("Sentiment as columns, and a donut of the mix.");
expect("listen columns", columns.listeningChart, "columns");
expect("donut wins last", columns.mix, "donut");

const asurion = decodeBrief("We closed 12,400 claims for Asurion. 6,200 were fulfilled. Approval is 9,100. The active book is 40k.");
expect("asurion client", asurion.client, "Asurion");
expect("asurion claims", asurion.claims, 12400);
expect("asurion fulfilled", asurion.fulfilled, 6200);
expect("asurion approved", asurion.approved, 9100);
expect("asurion active", asurion.active, 40000);

const notClaim = decodeBrief("10,000 active subscribers. Containment rate is 64%. Authentication rate is 91%.");
expect("containment is not a claim rate", notClaim.claimRate, undefined);
const notClaimBook = applyBrief(emptyBook(), notClaim, { counts: "fill", words: "fill" });
expect("containment does not invent claims", notClaimBook.claims, 0);
expect("active stays", notClaimBook.active, 10000);

const percentApproval = decodeBrief("Presentation for Pinetop Shield. Approval is 88%.");
expect("percent approval is not a count", percentApproval.approved, undefined);

if (failures.length) {
  console.error(failures.join("\n"));
  throw new Error("brief checks failed");
}
console.log("brief checks passed");
