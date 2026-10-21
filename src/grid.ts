import type { Book, SheetColumn, SheetGrid } from "./book";
import { chartFit } from "./fit";

export type { SheetColumn, SheetGrid };

export type GridProgress = {
  title: "Progress" | "Not progress";
  kind: "share" | "pair";
  column: string;
  leftLabel: string;
  left: number;
  rightLabel: string;
  right: number;
  note: string;
};

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const COUNT_WORD = /\b(calls?|chats?|repairs?|costs?|revenues?|subscribers?|claims?|volumes?|totals?|counts?|units?|tickets?|sessions?|contacts?|sales?|orders?|amounts?)\b/i;
const LABEL_HEADER = /\b(month|year|date|period|label|category|name|quarter|week)\b/i;

type Cell = string | number | null;

export function isGrid(book: Book): boolean {
  const grid = book.grid;
  return Boolean(grid && grid.columns.length > 0 && grid.labels.length > 1);
}

export function gridPeriod(grid: SheetGrid): string {
  if (grid.cadence === "rows" || grid.labels.length < 2) return "";
  const first = grid.labels[0] ?? "";
  const last = grid.labels[grid.labels.length - 1] ?? "";
  if (!first || !last) return "";
  return first === last ? first : `${first} to ${last}`;
}

export function gridLine(grid: SheetGrid): string {
  const names = grid.columns.map((column) => column.name);
  const shown = names.slice(0, 4).join(", ");
  const more = names.length > 4 ? `, and ${names.length - 4} more` : "";
  return `${grid.labels.length} rows. ${shown}${more}. Each row stays its own point.`;
}

/** Straight line through the points, evaluated at `x`. Null when fewer than 3 points. */
export function projectAt(xs: number[], ys: number[], x: number): number | null {
  if (xs.length !== ys.length || xs.length < 3) return null;
  const n = xs.length;
  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumXX = 0;
  for (let index = 0; index < n; index += 1) {
    const px = xs[index] ?? 0;
    const py = ys[index] ?? 0;
    sumX += px;
    sumY += py;
    sumXY += px * py;
    sumXX += px * px;
  }
  const denom = n * sumXX - sumX * sumX;
  if (denom === 0) return null;
  const slope = (n * sumXY - sumX * sumY) / denom;
  const intercept = (sumY - slope * sumX) / n;
  const next = intercept + slope * x;
  if (!Number.isFinite(next)) return null;
  const integers = ys.every((value) => Math.abs(value - Math.round(value)) < 1e-9);
  return integers ? Math.round(next) : Math.round(next * 10) / 10;
}

/** One step past the last row index. Blank cells stay out of the fit. */
export function projectNext(values: (number | null)[]): number | null {
  const xs: number[] = [];
  const ys: number[] = [];
  values.forEach((value, index) => {
    if (value == null || !Number.isFinite(value)) return;
    xs.push(index);
    ys.push(value);
  });
  return projectAt(xs, ys, values.length);
}

export function nextPeriodLabel(label: string, cadence: "month" | "year"): string {
  if (cadence === "year") {
    const year = Number(label.trim());
    if (Number.isInteger(year) && year >= 1900 && year <= 2099) return `${year + 1} proj`;
    return "Next proj";
  }
  const parsed = parseMonth(label);
  if (!parsed) return "Next proj";
  let month = parsed.month + 1;
  let year = parsed.year;
  if (month > 11) {
    month = 0;
    if (year != null) year += 1;
  }
  const name = MONTHS[month] ?? "Jan";
  if (year == null) return `${name} proj`;
  return `${name}-${String(year).slice(2)} proj`;
}

export function gridFromMatrix(matrix: Cell[][], title: string): SheetGrid | null {
  const rows = matrix
    .map((row) => row.map((cell) => (cell == null || cell === "" ? null : cell)))
    .filter((row) => row.some((cell) => cell != null && String(cell).trim() !== ""));
  if (rows.length < 2) return null;
  const headered = isHeaderRow(rows[0] ?? []);
  const header = headered ? rows[0]! : null;
  const body = (headered ? rows.slice(1) : rows).slice(0, 24);
  if (body.length < 2) return null;
  const trimmedRows = headered ? rows.length - 1 > 24 : rows.length > 24;
  const labelColumn = isLabelColumn(header?.[0], body);
  const labels = body.map((row, index) => {
    if (!labelColumn) return `Row ${index + 1}`;
    const text = cellText(row[0]);
    return text || `Row ${index + 1}`;
  });
  const start = labelColumn ? 1 : 0;
  const width = body.reduce((max, row) => Math.max(max, row.length), header?.length ?? 0);
  const columns: SheetColumn[] = [];
  let trimmedColumns = false;
  for (let index = start; index < width; index += 1) {
    const values = body.map((row) => asNumber(row[index] ?? null));
    if (values.filter((value) => value != null).length < 2) continue;
    if (columns.length >= 6) {
      trimmedColumns = true;
      break;
    }
    const name = columnName(header?.[index], columns.length);
    columns.push({ name, values, percent: isPercent(name, values) });
  }
  if (columns.length === 0 || labels.length < 2) return null;
  const cadence = cadenceOf(labels);
  const projection = cadence === "rows" ? columns.map(() => null) : columns.map((column) => projectNext(column.values));
  const hasProjection = projection.some((value) => value != null);
  const projectionLabel = hasProjection ? nextPeriodLabel(labels[labels.length - 1] ?? "", cadence === "year" ? "year" : "month") : null;
  const extra = [
    trimmedRows ? "First 24 rows." : "",
    trimmedColumns ? "First 6 columns." : "",
    projectionLabel ? "Projection is a straight line from the points in the sheet. It is not a result." : "",
  ].filter(Boolean);
  const note = `${labels.length} rows, ${columns.length} ${columns.length === 1 ? "column" : "columns"}. Each row stays its own point. Rows are not added together.${extra.length ? ` ${extra.join(" ")}` : ""}`;
  return {
    title: tidyTitle(title),
    labels,
    columns,
    cadence,
    projectionLabel,
    projection,
    note,
  };
}

export function gridProgress(grid: SheetGrid): GridProgress | null {
  const column = grid.columns.find((item) => item.percent) ?? grid.columns[0];
  if (!column) return null;
  const pairs = column.values.flatMap((value, index) => (value == null ? [] : [{ label: grid.labels[index] ?? `Row ${index + 1}`, value }]));
  const latest = pairs[pairs.length - 1];
  const previous = pairs.length > 1 ? pairs[pairs.length - 2] : null;
  if (!latest) return null;
  if (column.percent && latest.value <= 100) {
    const title = !previous || latest.value >= previous.value ? "Progress" : "Not progress";
    const rest = Math.max(0, Math.round((100 - latest.value) * 10) / 10);
    return {
      title,
      kind: "share",
      column: column.name,
      leftLabel: "Progress",
      left: latest.value,
      rightLabel: "Not progress",
      right: rest,
      note: `Latest row only, ${latest.label}. This is ${formatGridPoint(latest.value)} out of 100. Rows are not added together.`,
    };
  }
  if (!previous) return null;
  return {
    title: latest.value >= previous.value ? "Progress" : "Not progress",
    kind: "pair",
    column: column.name,
    leftLabel: previous.label,
    left: previous.value,
    rightLabel: latest.label,
    right: latest.value,
    note: `Last two rows of ${column.name}. Rows are not added together.`,
  };
}

export function columnPoints(grid: SheetGrid, index: number): { label: string; value: number; projected: boolean }[] {
  const column = grid.columns[index];
  if (!column) return [];
  const points: { label: string; value: number; projected: boolean }[] = [];
  column.values.forEach((value, row) => {
    if (value == null) return;
    points.push({ label: grid.labels[row] ?? `Row ${row + 1}`, value, projected: false });
  });
  const projected = grid.projection[index];
  if (grid.projectionLabel && projected != null) {
    points.push({ label: grid.projectionLabel, value: projected, projected: true });
  }
  return points;
}

export function gridChartNote(grid: SheetGrid, index: number): string {
  const column = grid.columns[index];
  if (!column) return "";
  const projected = grid.projection[index];
  const fit = chartFit(grid.labels, [{ name: column.name, values: column.values, percent: column.percent }]);
  const bits = [`${column.name} by row. Rows are not added together.`, fit.sentence];
  if (grid.projectionLabel && projected != null) {
    bits.push(`${grid.projectionLabel} is ${formatGridPoint(projected)}. Projection is a straight line from the points in the sheet. It is not a result.`);
    if (projected < 0) bits.push("The line falls below zero, so that column is not drawn above the axis.");
  }
  return bits.join(" ");
}

export function formatGridPoint(value: number): string {
  if (!Number.isFinite(value)) return "";
  if (Math.abs(value - Math.round(value)) < 0.05) return Math.round(value).toLocaleString("en-US");
  return value.toLocaleString("en-US", { maximumFractionDigits: 1 });
}

function isHeaderRow(row: Cell[]): boolean {
  if (row.length < 2) return false;
  const tail = row.slice(1);
  const tailNumbers = tail.filter((cell) => asNumber(cell) != null).length;
  const tailText = tail.filter((cell) => cellText(cell) && asNumber(cell) == null).length;
  return tailNumbers === 0 && tailText > 0;
}

function isLabelColumn(header: Cell | undefined, body: Cell[][]): boolean {
  if (LABEL_HEADER.test(cellText(header ?? null))) return true;
  if (body.length === 0) return false;
  const years = body.filter((row) => {
    const value = asNumber(row[0] ?? null);
    return value != null && value >= 1900 && value <= 2099 && Math.abs(value - Math.round(value)) < 1e-6;
  }).length;
  const text = body.filter((row) => asNumber(row[0] ?? null) == null && cellText(row[0] ?? null)).length;
  return text >= body.length * 0.6 || years >= body.length * 0.6;
}

function columnName(header: Cell | undefined, index: number): string {
  const text = cellText(header ?? null);
  if (text) return text.slice(0, 48);
  return index === 0 ? "Value" : `Series ${index + 1}`;
}

function isPercent(name: string, values: (number | null)[]): boolean {
  if (/%|percent|\brate\b/i.test(name)) return true;
  const nums = values.filter((value): value is number => value != null);
  if (nums.length < 2 || !nums.every((value) => value >= 0 && value <= 100)) return false;
  if (COUNT_WORD.test(name)) return false;
  if (/containment|authentication|approval/i.test(name)) return true;
  return nums.some((value) => Math.abs(value - Math.round(value)) > 0.001);
}

function cadenceOf(labels: string[]): "month" | "year" | "rows" {
  if (labels.length < 3) return "rows";
  const months = labels.filter((label) => parseMonth(label) != null).length;
  const years = labels.filter((label) => /^(19|20)\d{2}$/.test(label.trim())).length;
  if (months / labels.length >= 0.6) return "month";
  if (years / labels.length >= 0.6) return "year";
  return "rows";
}

function parseMonth(label: string): { month: number; year: number | null } | null {
  const named = label.trim().match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s\-/.]*(\d{2,4})?$/i);
  if (named) {
    const month = MONTHS.findIndex((name) => name.toLowerCase() === (named[1] ?? "").slice(0, 3).toLowerCase());
    if (month < 0) return null;
    if (!named[2]) return { month, year: null };
    const raw = Number(named[2]);
    const year = raw < 100 ? 2000 + raw : raw;
    return { month, year };
  }
  const iso = label.trim().match(/^(\d{4})[-/](\d{1,2})(?:[-/]\d{1,2})?$/);
  if (!iso) return null;
  const month = Number(iso[2]) - 1;
  if (month < 0 || month > 11) return null;
  return { month, year: Number(iso[1]) };
}

function asNumber(value: Cell): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;
  const cleaned = value.replace(/[$,]/g, "").replace(/%$/, "").trim();
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function cellText(value: Cell): string {
  if (value == null) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function tidyTitle(raw: string): string {
  const base = raw.replace(/\.(xlsx|xls|csv)$/i, "").replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  if (!base || /^(sheet\d*|book)$/i.test(base)) return "";
  return base
    .split(" ")
    .map((word, index) => {
      if (index > 0 && word === word.toLowerCase()) return word;
      if (word.length <= 4 && word === word.toUpperCase() && /[A-Z]/.test(word)) return word;
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(" ")
    .slice(0, 80);
}
