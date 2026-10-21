import { applyBrief, decodeBrief } from "./brief.ts";
import { factLine, readFacts, storyFromFacts } from "./facts.ts";
import { emptyBook } from "./parse.ts";
import { readMaterial } from "./reading.ts";
import { readStory, storyLine } from "./story.ts";

const failures: string[] = [];
const expect = (label: string, actual: unknown, wanted: unknown) => {
  if (actual !== wanted) failures.push(`${label}: ${String(actual)} !== ${String(wanted)}`);
};

const canonical =
  "Create a presentation showing that we handled 1.9M IVA Calls for Boost Mobile in this year where the containment rate is above 70% and Authentication is above 90%.";
const canonicalFacts = readFacts(canonical);
const canonicalStory = canonicalFacts ? storyFromFacts(canonicalFacts) : null;
expect("canonical client", canonicalStory?.client, "Boost Mobile");
expect("canonical period", canonicalStory?.period, "This year");
expect("canonical verb", canonicalStory?.verb, "handled");
expect("canonical unit", canonicalStory?.unit, "IVA calls");
expect("canonical volume", canonicalStory?.volume, 1900000);
expect("canonical hedge", canonicalStory?.hedge, undefined);
expect("canonical proofs", canonicalStory?.proofs.length, 2);
expect("canonical containment", canonicalStory?.proofs[0]?.label, "Containment");
expect("canonical containment relation", canonicalStory?.proofs[0]?.relation, "above");
expect("canonical containment value", canonicalStory?.proofs[0]?.value, 70);
expect("canonical auth", canonicalStory?.proofs[1]?.label, "Authentication");
expect("canonical auth relation", canonicalStory?.proofs[1]?.relation, "above");
expect("canonical auth value", canonicalStory?.proofs[1]?.value, 90);
expect(
  "canonical line",
  canonicalStory ? storyLine(canonicalStory) : "",
  "1.9 million IVA calls handled this year. Containment is above 70%. Authentication is above 90%.",
);

const messy = readFacts(
  "Boost handled about 1.9 million IVA calls this year. Containment sat above 70 percent and auth was over 90%.",
);
const messyStory = messy ? storyFromFacts(messy) : null;
expect("messy client", messyStory?.client, "Boost");
expect("messy period", messyStory?.period, "This year");
expect("messy volume", messyStory?.volume, 1900000);
expect("messy hedge", messyStory?.hedge, "about");
expect("messy unit", messyStory?.unit, "IVA calls");
expect("messy containment", `${messyStory?.proofs[0]?.label} ${messyStory?.proofs[0]?.relation} ${messyStory?.proofs[0]?.value}`, "Containment above 70");
expect("messy auth", `${messyStory?.proofs[1]?.label} ${messyStory?.proofs[1]?.relation} ${messyStory?.proofs[1]?.value}`, "Authentication above 90");
expect("messy keeps about", messyStory ? storyLine(messyStory).startsWith("about 1.9 million") : false, true);

const split = readFacts("Make a deck for the SLT. Voice containment is 64% on 1.5M calls. Chat is 90% on 243k chats.");
expect("split client", split?.client, "SLT");
expect("split groups", split?.groups.length, 2);
expect("split voice", split?.groups[0]?.name, "Voice");
expect("split voice volume", split?.groups[0]?.volume?.value, 1500000);
expect("split voice unit", split?.groups[0]?.volume?.unit, "Calls");
expect("split voice rate", split?.groups[0]?.rates[0]?.label, "Containment");
expect("split voice rate value", split?.groups[0]?.rates[0]?.value, 64);
expect("split chat", split?.groups[1]?.name, "Chat");
expect("split chat volume", split?.groups[1]?.volume?.value, 243000);
expect("split chat unit", split?.groups[1]?.volume?.unit, "Chats");
expect("split chat rate", split?.groups[1]?.rates[0]?.label, "Containment");
expect("split chat rate value", split?.groups[1]?.rates[0]?.value, 90);
expect("split not a share of voice", split?.groups[0]?.volume?.value === 960000, false);
expect("split not a share of chat", split?.groups[1]?.volume?.value === 218700, false);
expect("split line keeps both", factLine(split!).includes("1.5 million") && factLine(split!).includes("243 thousand"), true);
expect("split does not add them", factLine(split!).includes("1,738,401") || factLine(split!).includes("1.7 million"), false);
expect("short text is not a month track", readMaterial("Make a deck for the SLT. Voice containment is 64% on 1.5M calls. Chat is 90% on 243k chats."), null);

const claims = readFacts("We closed 12,400 claims for Asurion. 6,200 were fulfilled. Approval is 9,100. The active book is 40k.");
expect("claims sentence is not a fact deck", claims, null);
const claimsBrief = decodeBrief("We closed 12,400 claims for Asurion. 6,200 were fulfilled. Approval is 9,100. The active book is 40k.");
const claimsBook = applyBrief(emptyBook(), claimsBrief, { counts: "fill", words: "fill" });
expect("claims book active", claimsBook.active, 40000);
expect("claims book claims", claimsBook.claims, 12400);
expect("claims book fulfilled", claimsBook.fulfilled, 6200);
expect("claims book approved", claimsBook.approved, 9100);

const chats = readFacts("Containment rate is 64%. Authentication rate is 91%. We handled 2.4 million chats for T-Mobile last quarter.");
const chatsStory = chats ? storyFromFacts(chats) : null;
expect("chats client", chatsStory?.client, "T-Mobile");
expect("chats period", chatsStory?.period, "Last quarter");
expect("chats volume", chatsStory?.volume, 2400000);
expect("chats unit", chatsStory?.unit, "Chats");
expect("chats containment", `${chatsStory?.proofs[0]?.label} ${chatsStory?.proofs[0]?.value}`, "Containment 64");
expect("chats auth", `${chatsStory?.proofs[1]?.label} ${chatsStory?.proofs[1]?.value}`, "Authentication 91");
expect("chats claim rate", decodeBrief("Containment rate is 64%. Authentication rate is 91%. We handled 2.4 million chats for T-Mobile last quarter.").claimRate, undefined);

const verizon = readFacts("IVA for Verizon: 800k calls, containment above 75%, authentication at least 88%.");
const verizonStory = verizon ? storyFromFacts(verizon) : null;
expect("verizon client", verizonStory?.client, "Verizon");
expect("verizon volume", verizonStory?.volume, 800000);
expect("verizon unit", verizonStory?.unit, "Calls");
expect("verizon containment", `${verizonStory?.proofs[0]?.relation} ${verizonStory?.proofs[0]?.value}`, "above 75");
expect("verizon auth", `${verizonStory?.proofs[1]?.label} ${verizonStory?.proofs[1]?.relation} ${verizonStory?.proofs[1]?.value}`, "Authentication least 88");

const repairs = readFacts("Presentation for Pinetop Shield. We served 12,400 repairs in 2025. Approval is 88%.");
const repairsStory = repairs ? storyFromFacts(repairs) : null;
expect("repairs client", repairsStory?.client, "Pinetop Shield");
expect("repairs volume", repairsStory?.volume, 12400);
expect("repairs unit", repairsStory?.unit, "Repairs");
expect("repairs period", repairsStory?.period, "2025");
expect("repairs approval", repairsStory?.proofs[0]?.value, 88);
expect("repairs agrees with story volume", readStory("Presentation for Pinetop Shield. We served 12,400 repairs in 2025. Approval is 88%.")?.volume, 12400);

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log("facts checks passed");
