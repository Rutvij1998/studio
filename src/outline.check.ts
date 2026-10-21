import { markRequest } from "./mark.ts";
import { emptyBook } from "./parse.ts";
import { isOutline, readOutline } from "./outline.ts";
import { acceptShape, needsShape } from "./shape.ts";
import { pictureFromModel } from "./seen.ts";

const failures: string[] = [];
const expect = (name: string, actual: unknown, wanted: unknown) => {
  const left = JSON.stringify(actual);
  const right = JSON.stringify(wanted);
  if (left !== right) failures.push(`${name}: ${left} !== ${right}`);
};
const includes = (name: string, actual: string, wanted: string) => {
  if (!actual.includes(wanted)) failures.push(`${name}: missing ${wanted}`);
};

const walmart = "Hey, I need to create a deck for Walmart where we have a trade-in platform journey. Give me the important point that needs to be mentioned.";
const outline = readOutline(walmart);
expect("walmart client", outline?.client, "Walmart");
expect("walmart shape", outline?.shape, "journey");
expect("walmart steps", outline?.cards.length, 6);
expect("walmart first", outline?.cards[0]?.title, "Get a device value");
expect("walmart last", outline?.cards[5]?.title, "Pay and recover");
includes("walmart intent", outline?.intent ?? "", "points to mention");
includes("walmart hint", outline?.cards[0]?.hint ?? "", "Point to add");
expect("walmart is outline", isOutline({ ...emptyBook(), client: "Walmart", outline }), true);
expect("walmart needs no model", needsShape(walmart), false);

const training = "Make a deck for Asurion about store training.";
const points = readOutline(training);
expect("training client", points?.client, "Asurion");
expect("training shape", points?.shape, "points");
expect("training subject", points?.subject, "Store training");
expect("training count", points?.cards.length, 4);

const listed = readOutline("Deck for Verizon. Journey stages: Quote, Ship, and Pay.");
expect("listed count", listed?.cards.length, 3);
expect("listed first", listed?.cards[0]?.title, "Quote");
expect("listed client", listed?.client, "Verizon");

const boost = "Create a presentation showing that we handled 1.9M IVA Calls for Boost Mobile in this year where the containment rate is above 70% and Authentication is above 90%.";
expect("boost stays figures", readOutline(boost), null);
expect("boost needs no model", needsShape(boost), false);

const scorecard = "Scorecard for Northline Care. Show the mix as columns.";
expect("empty scorecard stays a scorecard", readOutline(scorecard), null);

expect("boost mark domain", markRequest("Boost Mobile")?.domain, "boostmobile.com");
expect("boost mark file", markRequest("Boost Mobile")?.file, null);
expect("walmart file", markRequest("Walmart")?.file, "walmart.svg");
expect("slt mark", markRequest("SLT"), null);
expect("sample mark", markRequest("Northline Care"), null);
expect("greeting mark", markRequest("Hi Martin"), null);

expect("changed figure", acceptShape("We handled 1.9 million calls.", "We handled 2 million calls."), null);
expect("kept figure", acceptShape("We handled 1.9 million calls.", "We handled 1.9 million calls."), "We handled 1.9 million calls.");

const seen = pictureFromModel('{"title":"Calls","kind":"pie","categories":["Jan-24","Feb-24","Mar-24"],"series":[{"name":"Calls","values":[10,20,30]}]}');
expect("picture kind", seen?.chart?.kind, "line");
includes("picture drawn", seen?.chart?.drawn ?? "", "Read from the picture");
expect("picture ignores kind", pictureFromModel("not json"), null);
const share = pictureFromModel('{"title":"Mix","kind":"columns","categories":["Kept","Open"],"series":[{"name":"Share","values":[64,36]}]}');
expect("picture donut", share?.chart?.kind, "donut");

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log("outline ok");
