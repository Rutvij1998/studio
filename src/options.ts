export type MixChart = "donut" | "columns" | "meter";
export type ListeningChart = "rows" | "columns";
export type TypeSize = "s" | "m" | "l";

export type DeckOptions = {
  mix: MixChart;
  listening: ListeningChart;
  type: TypeSize;
};

export const DEFAULT_OPTIONS: DeckOptions = {
  mix: "donut",
  listening: "rows",
  type: "m",
};

export function deckFileName(client: string): string {
  const base = client
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
  return `${base || "scorecard"}-scorecard.pptx`;
}
