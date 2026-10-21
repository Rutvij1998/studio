import { chartFit, comboPair } from "./fit.ts";
import { emptyBook } from "./parse.ts";
import { slidesFor } from "./slides.tsx";

const failures: string[] = [];
const expect = (name: string, actual: unknown, wanted: unknown) => {
  const left = JSON.stringify(actual);
  const right = JSON.stringify(wanted);
  if (left !== right) failures.push(`${name}: ${left} !== ${right}`);
};

expect("months are a line", chartFit(["Jan-24", "Feb-24", "Mar-24"], [{ name: "Repairs", values: [100, 110, 130] }]).kind, "line");
expect("share is a donut", chartFit(["Kept", "Open"], [{ name: "Share", values: [64, 36] }]).kind, "donut");
expect(
  "long labels are bars",
  chartFit(["Device eligibility review", "Payment"], [{ name: "Count", values: [12, 8] }]).kind,
  "bars",
);
expect("years are a line", chartFit(["2022", "2023", "2024"], [{ name: "Policies", values: [10, 20, 30] }]).kind, "line");
expect(
  "count and rate",
  comboPair(
    ["Jan-24", "Feb-24", "Mar-24"],
    [
      { name: "Calls", values: [100, 110, 130] },
      { name: "Containment", values: [80, 70, 60], percent: true },
    ],
  ),
  { count: 0, rate: 1 },
);
expect(
  "two counts are not a combo",
  comboPair(
    ["Jan-24", "Feb-24", "Mar-24"],
    [
      { name: "Repairs", values: [100, 110, 130] },
      { name: "Cost", values: [400, 420, 450] },
    ],
  ),
  null,
);

const together = slidesFor({
  ...emptyBook(),
  client: "Calls",
  grid: {
    title: "Calls",
    labels: ["Jan-24", "Feb-24", "Mar-24"],
    columns: [
      { name: "Calls", values: [100, 110, 130], percent: false },
      { name: "Containment", values: [80, 70, 60], percent: true },
    ],
    cadence: "month",
    projectionLabel: "Apr-24 proj",
    projection: [150, 50],
    note: "Rows are not added together.",
  },
});
expect("combo slide is added", together.some((slide) => slide.id === "combo"), true);
expect("each series stays", together.filter((slide) => slide.id === "grid").length, 2);

if (failures.length) {
  console.error(failures.join("\n"));
  process.exit(1);
}
console.log("fit ok");
