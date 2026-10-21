import * as XLSX from "xlsx";
import { decodeBrief } from "./brief.ts";
import { readSheet } from "./parse.ts";
import { missedTable, rateOutlook, readMaterial, readingSummary, trackNote, trackTotals } from "./reading.ts";

const failures: string[] = [];
const expect = (label: string, actual: unknown, wanted: unknown) => {
  if (actual !== wanted) failures.push(`${label}: ${String(actual)} !== ${String(wanted)}`);
};

const prompt = `Presentation for SLT.
Period: March 2025 to October 2026.
We handled 1,495,401 IVA voice calls. Voice containment is 64%. Chat containment is 90%.
Recommendation: Voice includes only inbound calls on accounts where IVA is live that entered portal deflection, rapid transfer, or direct Auth. It excludes accounts without IVA, short calls that never reached those paths, and any call that reached an agent. Chat includes only API-initiated chats that entered the IVA chatbot, with the same exclusions. Containment means the contact did not reach an agent. October 2026 is the short month in the table, and March 2025 voice is 11 calls.

Voice
Mar-25: 11 calls, 11 contained, 100%
Apr-25: 12035 calls, 9967 contained, 83%
May-25: 85566 calls, 68031 contained, 80%
Jun-25: 75267 calls, 53921 contained, 72%
Jul-25: 74967 calls, 50794 contained, 68%
Aug-25: 80231 calls, 54746 contained, 68%
Sep-25: 95459 calls, 62966 contained, 66%
Oct-25: 104018 calls, 64379 contained, 62%
Nov-25: 95134 calls, 61281 contained, 64%
Dec-25: 102310 calls, 65987 contained, 64%
Jan-26: 104169 calls, 62843 contained, 60%
Feb-26: 83022 calls, 49035 contained, 59%
Mar-26: 90980 calls, 55862 contained, 61%
Apr-26: 83041 calls, 50833 contained, 61%
May-26: 73023 calls, 44138 contained, 60%
Jun-26: 76161 calls, 45549 contained, 60%
Jul-26: 78336 calls, 45946 contained, 59%
Aug-26: 72201 calls, 41973 contained, 58%
Sep-26: 76093 calls, 43008 contained, 57%
Oct-26: 33377 calls, 19706 contained, 59%

Chat
Mar-26: 58651 chats, 56370 contained, 96%
Apr-26: 74946 chats, 72171 contained, 96%
May-26: 33919 chats, 31319 contained, 92%
Jun-26: 15339 chats, 12130 contained, 79%
Jul-26: 16610 chats, 12925 contained, 78%
Aug-26: 16928 chats, 13384 contained, 79%
Sep-26: 18101 chats, 12934 contained, 71%
Oct-26: 8111 chats, 6277 contained, 77%`;

const reading = readMaterial(prompt);
if (!reading) failures.push("reading missing");
else {
  expect("title", reading.title, "SLT");
  expect("period", reading.period, "March 2025 to October 2026");
  expect("tracks", reading.tracks.length, 2);
  expect("voice points", reading.tracks[0]?.points.length, 20);
  expect("chat points", reading.tracks[1]?.points.length, 8);
  const voice = trackTotals(reading.tracks[0]!);
  const chat = trackTotals(reading.tracks[1]!);
  expect("voice total", voice.total, 1495401);
  expect("voice rate", voice.rate, 64);
  expect("chat total", chat.total, 242605);
  expect("chat rate", chat.rate, 90);
  expect("scope", reading.scope.length, 3);
  expect("ask", reading.ask, "");
  expect("intent", reading.intent.includes("included"), true);
  expect("subject", reading.subject.includes("no agent"), true);
  const note = trackNote(reading.tracks[0]!);
  expect("voice move", note.includes("83% in Apr-25"), true);
  expect("voice end", note.includes("57% in Sep-26"), true);
  expect("summary title", readingSummary(reading).startsWith("Title: SLT."), true);
}

expect("scorecard ignored", readMaterial("Scorecard for Northline Care. 48,200 active subscribers and 19,280 claims."), null);

const email = `Hi Martin,
IVA (Voice)
Included: Inbound voice calls on accounts where IVA is live.
Excluded: Accounts without IVA, and any call that connected to an agent.
Containment definition: Call did not reach an agent.
Mar-25 11 11 100
Apr-25 12035 9967 83
May-25 85566 68031 80
Chatbot
Included: API-initiated chats that went through the IVA chatbot flow.
Excluded: Chats that never entered the IVA path.
Mar-26 58651 56370 96
Apr-26 74946 72171 96
May-26 33919 31319 92
Jun-26 15339 12130 79`;
const mailed = readMaterial(email);
expect("email tracks", mailed?.tracks.length, 2);
expect("email voice", mailed?.tracks[0]?.name, "Voice");
expect("email chat", mailed?.tracks[1]?.name, "Chat");
expect("email scope", (mailed?.scope.length ?? 0) >= 3, true);

const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(
  book,
  XLSX.utils.aoa_to_sheet([
    ["MM_YYYY", "Total Calls", "Contained Calls", "%"],
    ["Mar-25", 11, 11, 100],
    ["Apr-25", 12035, 9967, 83],
    ["May-25", 85566, 68031, 80],
    ["Jun-25", 75267, 53921, 72],
  ]),
  "Voice",
);
XLSX.utils.book_append_sheet(
  book,
  XLSX.utils.aoa_to_sheet([
    ["MM_YYYY", "Total Chats", "Contained Chats", "%"],
    ["Mar-26", 58651, 56370, 96],
    ["Apr-26", 74946, 72171, 96],
    ["May-26", 33919, 31319, 92],
    ["Jun-26", 15339, 12130, 79],
  ]),
  "Chat",
);
const bytes = XLSX.write(book, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
const sheet = readSheet(bytes, "iva.xlsx");
expect("sheet tracks", sheet.book.reading?.tracks.length, 2);
expect("sheet voice", sheet.book.reading?.tracks[0]?.points.length, 4);
expect("sheet chat unit", sheet.book.reading?.tracks[1]?.unit, "chats");
expect("sheet title", sheet.book.client, "Iva");

const voiceCells = [
  ["Mar-25", "11", "11", "100"],
  ["Apr-25", "12035", "9967", "83"],
  ["May-25", "85566", "68031", "80"],
  ["Jun-25", "75267", "53921", "72"],
  ["Jul-25", "74967", "50794", "68"],
  ["Aug-25", "80231", "54746", "68"],
  ["Sep-25", "95459", "62966", "66"],
  ["Oct-25", "104018", "64379", "62"],
  ["Nov-25", "95134", "61281", "64"],
  ["Dec-25", "102310", "65987", "64"],
  ["Jan-26", "104169", "62843", "60"],
  ["Feb-26", "83022", "49035", "59"],
  ["Mar-26", "90980", "55862", "61"],
  ["Apr-26", "83041", "50833", "61"],
  ["May-26", "73023", "44138", "60"],
  ["Jun-26", "76161", "45549", "60"],
  ["Jul-26", "78336", "45946", "59"],
  ["Aug-26", "72201", "41973", "58"],
  ["Sep-26", "76093", "43008", "57"],
  ["Oct-26", "33377", "19706", "59"],
];
const chatCells = [
  ["Mar-26", "58651", "56370", "96"],
  ["Apr-26", "74946", "72171", "96"],
  ["May-26", "33919", "31319", "92"],
  ["Jun-26", "15339", "12130", "79"],
  ["Jul-26", "16610", "12925", "78"],
  ["Aug-26", "16928", "13384", "79"],
  ["Sep-26", "18101", "12934", "71"],
  ["Oct-26", "8111", "6277", "77"],
];
const stacked = `Can you make a presentation for
Hi Martin, Jen,
Here’s the updated view for the SLT slides, with a clear summary of what is included vs excluded so the numbers are defensible.
IVA (Voice)
Included: Inbound voice calls on accounts where IVA is live, that actually entered an IVA path (portal deflection, rapid transfer, or direct Auth).
Excluded: Accounts without IVA, short/incomplete interactions that never reached those paths, and any call that connected to an agent.
Containment definition: Call did not reach an agent.
IVA
MM_YYYY
Total Calls
Contained Calls
%
${voiceCells.map((row) => row.join("\n")).join("\n")}

Chatbot
Included: API-initiated chats that went through the IVA chatbot flow.
Excluded: Chats that never entered the IVA path and any chat that connected to an agent.
Containment definition: Chat did not reach an agent. (Same filters as the numbers previously shared.)
MM_YYYY
Total Chats
Contained Chats
%
${chatCells.map((row) => row.join("\n")).join("\n")}

Both use the same containment definition (no agent reached) and only the populations that actually went through IVA. Happy to add regional breakouts or adjust the narrative if needed.
Thanks,`;
const pasted = readMaterial(stacked);
if (!pasted) failures.push("stacked paste missing");
else {
  expect("stacked title", pasted.title, "SLT");
  expect("stacked period", pasted.period, "Mar-25 to Oct-26");
  expect("stacked tracks", pasted.tracks.length, 2);
  expect("stacked voice name", pasted.tracks[0]?.name, "Voice");
  expect("stacked chat name", pasted.tracks[1]?.name, "Chat");
  expect("stacked voice points", pasted.tracks[0]?.points.length, 20);
  expect("stacked chat points", pasted.tracks[1]?.points.length, 8);
  expect("stacked voice unit", pasted.tracks[0]?.unit, "calls");
  expect("stacked chat unit", pasted.tracks[1]?.unit, "chats");
  const voice = trackTotals(pasted.tracks[0]!);
  const chat = trackTotals(pasted.tracks[1]!);
  expect("stacked voice total", voice.total, 1495401);
  expect("stacked voice rate", voice.rate, 64);
  expect("stacked chat total", chat.total, 242605);
  expect("stacked chat rate", chat.rate, 90);
  expect("stacked scope", pasted.scope.map((card) => card.label).join(","), "Voice,Chat,Containment");
  expect("stacked voice scope", pasted.scope[0]?.text.includes("portal deflection"), true);
  expect("stacked intent", pasted.intent.includes("included"), true);
  expect("stacked ask", pasted.ask.includes("regional"), true);
  const voiceNote = trackNote(pasted.tracks[0]!);
  expect("stacked voice move", voiceNote.includes("83% in Apr-25"), true);
  expect("stacked voice end", voiceNote.includes("57% in Sep-26"), true);
  expect("stacked short month", voiceNote.includes("Oct-26"), true);
  const chatNote = trackNote(pasted.tracks[1]!);
  expect("stacked chat move", chatNote.includes("96% in Mar-26"), true);
  expect("stacked chat end", chatNote.includes("71% in Sep-26"), true);
  expect("stacked chat drop", chatNote.includes("Jun-26"), true);
  const voiceOutlook = rateOutlook(pasted.tracks[0]!);
  expect("stacked outlook", voiceOutlook?.label, "Nov-26 proj");
  expect("stacked outlook from", voiceOutlook?.from, "Apr-25");
  expect("stacked outlook to", voiceOutlook?.to, "Sep-26");
}
const greeted = decodeBrief(stacked);
expect("greeting client", greeted.client, undefined);
expect("greeting active", greeted.active, undefined);
expect("greeting claims", greeted.claims, undefined);
expect("stacked was read", missedTable(stacked), null);
expect(
  "two numbers stay a grid",
  readMaterial("Jan-24\n100\n400\nFeb-24\n110\n420\nMar-24\n130\n450"),
  null,
);
expect(
  "broken table is named",
  missedTable("MM_YYYY\nTotal Calls\nContained Calls\n%\nMar-25\nsee note")?.includes("not read"),
  true,
);
expect("scorecard is not a missed table", missedTable("Scorecard for Northline Care. 48,200 active subscribers and 19,280 claims."), null);

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log("reading checks passed");
