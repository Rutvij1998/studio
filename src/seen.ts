import type { SourceChart, SourceNumber } from "./book";
import { chartFit } from "./fit";

export type PictureRead = {
  title: string;
  chart: SourceChart | null;
  numbers: SourceNumber[];
};

function numberOf(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;
  const match = value.trim().match(/^(\d[\d,]*(?:\.\d+)?)(?:\s*(%|percent|million|billion|thousand|m|k|bn|b))?$/i);
  if (!match) return null;
  const raw = Number((match[1] ?? "").replace(/,/g, ""));
  if (!Number.isFinite(raw)) return null;
  const unit = (match[2] ?? "").toLowerCase();
  const mult = unit === "k" || unit === "thousand" ? 1e3 : unit === "m" || unit === "million" ? 1e6 : unit === "b" || unit === "bn" || unit === "billion" ? 1e9 : 1;
  return Math.round(raw * mult);
}

/** The model returns JSON. The chart type is chosen from the numbers, not from the model's guess. */
export function pictureFromModel(text: string): PictureRead | null {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start < 0 || end <= start) return null;
  let body: unknown;
  try {
    body = JSON.parse(text.slice(start, end + 1));
  } catch {
    return null;
  }
  if (!body || typeof body !== "object") return null;
  const record = body as {
    title?: unknown;
    categories?: unknown;
    series?: unknown;
    numbers?: unknown;
  };
  const title = typeof record.title === "string" ? record.title.trim().slice(0, 140) : "";
  const categories = Array.isArray(record.categories)
    ? record.categories.filter((item): item is string => typeof item === "string").map((item) => item.trim()).filter(Boolean).slice(0, 24)
    : [];
  const seriesIn = Array.isArray(record.series) ? record.series.slice(0, 6) : [];
  const series = seriesIn
    .map((item) => {
      if (!item || typeof item !== "object") return null;
      const row = item as { name?: unknown; values?: unknown };
      const values = Array.isArray(row.values) ? row.values.map(numberOf).slice(0, categories.length || 24) : [];
      const finite = values.filter((value): value is number => value != null);
      if (finite.length < 2) return null;
      return {
        name: typeof row.name === "string" && row.name.trim() ? row.name.trim().slice(0, 48) : "Value",
        values: finite,
      };
    })
    .filter((item): item is { name: string; values: number[] } => item != null);
  const width = Math.min(categories.length, series[0]?.values.length ?? 0);
  if (series.length && width >= 2) {
    const cats = categories.slice(0, width);
    const drawn = series.map((item) => ({ name: item.name, values: item.values.slice(0, width) }));
    const fit = chartFit(cats, drawn.map((item) => ({ ...item, percent: item.values.every((value) => value >= 0 && value <= 100) })));
    const chart: SourceChart = {
      kind: fit.kind === "line" ? "line" : fit.kind === "bars" ? "bars" : fit.kind === "donut" ? "donut" : "columns",
      drawn: `Read from the picture. ${fit.sentence}`,
      categories: cats,
      series: drawn,
    };
    return { title, chart, numbers: [] };
  }
  const numbers = Array.isArray(record.numbers)
    ? record.numbers
        .map((item) => {
          if (!item || typeof item !== "object") return null;
          const row = item as { label?: unknown; display?: unknown; value?: unknown };
          const value = numberOf(row.value ?? row.display);
          if (value == null) return null;
          const display = typeof row.display === "string" && row.display.trim() ? row.display.trim() : String(value);
          const label = typeof row.label === "string" ? row.label.trim().slice(0, 48) : "";
          return { label, value, display };
        })
        .filter((item): item is SourceNumber => item != null)
        .slice(0, 8)
    : [];
  if (!title && numbers.length === 0) return null;
  return { title, chart: null, numbers };
}
