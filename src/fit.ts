export type FitKind = "line" | "columns" | "bars" | "donut" | "combo";

export type FitSeries = {
  name: string;
  values: (number | null)[];
  percent?: boolean;
};

const MONTH =
  /^(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b/i;

/** Months or years in a row are a time series. A handful of names is not. */
export function isTimeLabels(labels: string[]): boolean {
  if (labels.length < 3) return false;
  const hits = labels.filter((label) => {
    const text = label.trim();
    return MONTH.test(text) || /^(19|20)\d{2}$/.test(text) || /\b(19|20)\d{2}\b/.test(text);
  }).length;
  return hits / labels.length >= 0.6;
}

function finite(values: (number | null)[]): number[] {
  return values.filter((value): value is number => value != null && Number.isFinite(value));
}

/** The chart that fits these points. A pie is drawn as a donut. */
export function chartFit(labels: string[], series: FitSeries[]): { kind: Exclude<FitKind, "combo">; sentence: string } {
  const lead = series[0];
  const values = lead ? finite(lead.values) : [];
  const time = isTimeLabels(labels);
  if (time) return { kind: "line", sentence: "Drawn as a line. The points run across months or years." };
  const sum = values.reduce((total, value) => total + Math.max(0, value), 0);
  const parts = values.length >= 2 && values.length <= 6 && values.every((value) => value >= 0);
  const whole = parts && sum >= 98 && sum <= 102;
  if (whole) return { kind: "donut", sentence: "Drawn as a donut. The values are parts of one whole." };
  const long = labels.some((label) => label.trim().length > 16);
  if (long) return { kind: "bars", sentence: "Drawn as bars. The labels are long." };
  return { kind: "columns", sentence: "Drawn as columns. The labels are categories." };
}

/** A count and a rate on the same months share one slide, each on its own axis. */
export function comboPair(labels: string[], series: FitSeries[]): { count: number; rate: number } | null {
  if (!isTimeLabels(labels)) return null;
  const rate = series.findIndex((item) => item.percent);
  const count = series.findIndex((item) => !item.percent && finite(item.values).length >= 3);
  if (rate < 0 || count < 0) return null;
  if (finite(series[rate]!.values).length < 3) return null;
  return { count, rate };
}
