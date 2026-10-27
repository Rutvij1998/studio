import { type CSSProperties, type ReactElement, useEffect, useState } from "react";
import { coverCopy } from "./brief";
import type { ListeningChart, MixChart, TypeSize } from "./options";
import {
  conversationOf,
  formatCount,
  hasBroughtIn,
  type Book,
  type BookView,
  type JourneyStage,
  type ListeningRow,
  type Slice,
  type Tone,
} from "./book";
import { compactVolume, isResults, proofPhrase, storyLine, storyReadLines } from "./story";
import { factLine, factReadLines, isFacts } from "./facts";
import { chartFit, comboPair } from "./fit";
import { columnPoints, formatGridPoint, gridChartNote, gridLine, gridProgress, isGrid } from "./grid";
import { isOutline, outlineLine } from "./outline";
import { isSeries, rateOutlook, readingLine, seriesChartsSource, seriesProgressList, trackNote, trackTotals } from "./reading";
import { coverWave } from "./dotWave";
import { axisMax, axisSpan, bulletLines, cardFace, chartEyebrow, chartFoot, chartSentence, compareSplit, directedComposition, copyBlocks, displayPx, figureBody, figurePx, figureShare, fillBandIndex, formLabel, formatPoint, gridAcross, heroAside, imageSize, isContinued, isSource, isTimeTitle, leadCopy, pictureLayout, pieceAir, pieceBands, pieceFit, plotMode, posterPx, presentKind, proofPx, proofSplit, railLabel, rowDisplay, sameFigure, slideRole, slideShow, spareHeaders, tableFromSlide, tableHasNumbers } from "./sourceDeck";
import type { SheetGrid, SlidePiece, SourceChart, SourceSlide, SourceTable } from "./book";

export type SlideId =
  | "cover"
  | "book"
  | "mix"
  | "versus"
  | "listening"
  | "journey"
  | "months"
  | "source"
  | "next"
  | "result"
  | "scope"
  | "track"
  | "compare"
  | "origin"
  | "facts"
  | "progress"
  | "outlook"
  | "grid"
  | "combo"
  | "card"
  | "read";

export type SlideMeta = { id: SlideId; label: string; form: string; track?: number; origin?: number };

function clipLabel(text: string): string {
  const clean = text.trim();
  return clean.length > 32 ? `${clean.slice(0, 32).trimEnd()}…` : clean;
}

function gridSlides(book: Book): SlideMeta[] {
  const grid = book.grid;
  if (!grid) return [];
  const slides: SlideMeta[] = [];
  const progress = gridProgress(grid);
  if (progress) slides.push({ id: "progress", label: progress.title, form: "Latest" });
  grid.columns.forEach((column, index) => {
    slides.push({
      id: "grid",
      label: clipLabel(column.name) || `Column ${index + 1}`,
      form: grid.projectionLabel ? "Chart and projection" : "Chart",
      track: index,
    });
  });
  const pair = comboPair(
    grid.labels,
    grid.columns.map((column) => ({ name: column.name, values: column.values, percent: column.percent })),
  );
  if (pair) slides.push({ id: "combo", label: "Together", form: "Two axes" });
  return slides;
}

/** Chart slides appear only when that sheet was actually read. */
export function slidesFor(book: Book): SlideMeta[] {
  if (isSource(book) && book.sourceDeck) {
    const origin = book.sourceDeck.slides.map((slide) => ({
      id: "origin" as const,
      label: railLabel(slide),
      form: formLabel(slide),
      track: slide.index,
      origin: slide.from,
    }));
    return origin;
  }
  if (isSeries(book) && book.reading) {
    const reading = book.reading;
    const slides: SlideMeta[] = [{ id: "cover", label: "Cover", form: "Title" }];
    if (reading.scope.length > 0) slides.push({ id: "scope", label: "What counts", form: "Scope" });
    const progress = seriesProgressList(reading);
    if (progress.length > 0) {
      slides.push({ id: "progress", label: progress.length === 1 ? progress[0]!.direction : "Progress", form: "Latest month" });
    }
    reading.tracks.forEach((track, index) => {
      slides.push({ id: "track", label: track.name, form: "By month", track: index });
    });
    if (reading.tracks.length > 1) slides.push({ id: "compare", label: "Side by side", form: "Rates" });
    if (reading.tracks.some((track) => rateOutlook(track))) slides.push({ id: "outlook", label: "Outlook", form: "Projection" });
    if (hasBroughtIn(book) && !seriesChartsSource(book)) slides.push({ id: "source", label: "Brought in", form: "Decoded text" });
    if (book.recommendation.trim()) slides.push({ id: "next", label: "What to do", form: "Recommendation" });
    return slides;
  }
  if (isGrid(book)) {
    const slides: SlideMeta[] = [{ id: "cover", label: "Cover", form: "Title" }, ...gridSlides(book)];
    if (book.recommendation.trim()) slides.push({ id: "next", label: "What to do", form: "Recommendation" });
    return slides;
  }
  if (isFacts(book) && book.facts) {
    const slides: SlideMeta[] = [{ id: "cover", label: "Cover", form: "Title" }];
    book.facts.groups.forEach((group, index) => {
      slides.push({ id: "facts", label: group.name || "The result", form: "Figures", track: index });
    });
    if (factReadLines(book.facts).length > 0) slides.push({ id: "read", label: "What it says", form: "Read" });
    if (hasBroughtIn(book)) slides.push({ id: "source", label: "Brought in", form: "Decoded text" });
    if (book.recommendation.trim()) slides.push({ id: "next", label: "What to do", form: "Recommendation" });
    return slides;
  }
  if (isResults(book)) {
    const slides: SlideMeta[] = [
      { id: "cover", label: "Cover", form: "Title" },
      { id: "result", label: "The result", form: "Headline" },
    ];
    if (book.story && storyReadLines(book.story, book.period || book.story.period).length > 0) {
      slides.push({ id: "read", label: "What it says", form: "Read" });
    }
    if (hasBroughtIn(book)) slides.push({ id: "source", label: "Brought in", form: "Decoded text" });
    if (book.recommendation.trim()) slides.push({ id: "next", label: "What to do", form: "Recommendation" });
    return slides;
  }
  if (isOutline(book) && book.outline) {
    return [
      { id: "cover", label: "Cover", form: "Title" },
      ...book.outline.cards.map((card, index) => ({
        id: "card" as const,
        label: clipLabel(card.title) || `Slide ${index + 1}`,
        form: card.kicker,
        track: index,
      })),
    ];
  }
  const slides: SlideMeta[] = [
    { id: "cover", label: "Cover", form: "Title" },
    { id: "book", label: "Book", form: "Scorecard" },
    { id: "mix", label: "Mix", form: "Donut of the book" },
    { id: "versus", label: "Versus", form: "Head to head" },
  ];
  if (book.listening.length > 0) slides.push({ id: "listening", label: "Listening", form: "Sentiment" });
  if (book.journey.length > 0) slides.push({ id: "journey", label: "Journey", form: "Diverging" });
  if (book.months.length > 0) slides.push({ id: "months", label: "Months", form: "By month" });
  if (hasBroughtIn(book)) slides.push({ id: "source", label: "Brought in", form: "Decoded text" });
  slides.push({ id: "next", label: "What to do", form: "Recommendation" });
  return slides;
}

const TONE: Record<Tone, string> = {
  fulfilled: "fulfilled",
  approved: "approved",
  open: "open",
  rest: "rest",
};

function Donut({ slices, center, caption, label }: { slices: Slice[]; center: string; caption: string; label: string }) {
  const radius = 62;
  const turn = 2 * Math.PI * radius;
  const gap = slices.length > 1 ? 4 : 0;
  const budget = turn - gap * slices.length;
  const sum = slices.reduce((total, slice) => total + slice.count, 0) || 1;
  const minimum = 3;
  let lengths = slices.map((slice) => (slice.count / sum) * budget);
  lengths = lengths.map((length) => Math.max(length, minimum));
  const overflow = lengths.reduce((total, length) => total + length, 0) - budget;
  if (overflow > 0) {
    const flexible = lengths.reduce((total, length) => total + Math.max(0, length - minimum), 0) || 1;
    lengths = lengths.map((length) => length - overflow * (Math.max(0, length - minimum) / flexible));
  }
  let cursor = 0;
  return (
    <div className="donut">
      <svg viewBox="0 0 160 160" role="img" aria-label={label}>
        <g transform="rotate(-90 80 80)">
          {slices.map((slice, index) => {
            const length = lengths[index] ?? 0;
            const node = (
              <circle
                key={slice.key}
                className={TONE[slice.tone]}
                cx="80"
                cy="80"
                r={radius}
                strokeDasharray={`${length} ${turn - length}`}
                strokeDashoffset={-cursor}
              />
            );
            cursor += length + gap;
            return node;
          })}
        </g>
      </svg>
      <p>
        <strong data-wide={center.length > 4 ? "" : undefined}>{center}</strong>
        <span>{caption}</span>
      </p>
    </div>
  );
}

function Bar({ rate }: { rate: number }) {
  const width = Math.max(rate > 0 ? 2 : 0, Math.min(100, rate));
  return (
    <div className="meter" aria-hidden>
      <span style={{ width: `${width}%` }} />
    </div>
  );
}

function coverKicker(book: Book): string {
  if (isSource(book)) return "From the deck";
  if (isSeries(book)) return "Briefing";
  if (isGrid(book)) return "Sheet";
  if (isFacts(book)) return "Figures";
  if (isResults(book)) return "Results";
  if (isOutline(book)) return "Outline";
  return "Client scorecard";
}

const COVER_DOTS = coverWave();

function CoverWave() {
  return (
    <svg className="cover-wave" viewBox="0 0 1280 720" aria-hidden="true">
      {COVER_DOTS.map((dot, index) => (
        <circle
          key={index}
          cx={dot.x * 1280}
          cy={dot.y * 720}
          r={dot.r * 1280}
          fill={`#${dot.fill}`}
          opacity={dot.opacity}
          style={{ animationDelay: `${(dot.x * 0.7 + dot.y * 0.18).toFixed(3)}s` }}
        />
      ))}
    </svg>
  );
}

function figureParts(value: string): { prefix: string; digits: string; suffix: string } | null {
  const match = value.match(/^(\D{0,3})(\d[\d,]*)(\D{0,3})$/);
  if (!match?.[2]) return null;
  return { prefix: match[1] ?? "", digits: match[2], suffix: match[3] ?? "" };
}

function figureStart(value: string): string {
  const parts = figureParts(value);
  if (!parts) return value;
  return `${parts.prefix}0${parts.suffix}`;
}

/** Counts a figure that is already on the slide. The last frame is the source text. */
function Figure({ value }: { value: string }) {
  const reduce = typeof window !== "undefined" && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const [shown, setShown] = useState(reduce ? value : figureStart(value));
  useEffect(() => {
    if (reduce) {
      setShown(value);
      return;
    }
    const parts = figureParts(value);
    if (!parts) {
      setShown(value);
      return;
    }
    const target = Number(parts.digits.replace(/,/g, ""));
    if (!Number.isFinite(target)) {
      setShown(value);
      return;
    }
    const commas = parts.digits.includes(",");
    const start = performance.now();
    let frame = 0;
    const tick = (now: number) => {
      const t = Math.min(1, (now - start) / 720);
      if (t >= 1) {
        setShown(value);
        return;
      }
      const current = Math.round(target * (1 - (1 - t) ** 3));
      const text = commas ? current.toLocaleString("en-US") : String(current);
      setShown(`${parts.prefix}${text}${parts.suffix}`);
      frame = requestAnimationFrame(tick);
    };
    frame = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(frame);
  }, [reduce, value]);
  return <span className="figure" aria-hidden="true">{shown}</span>;
}

function CoverTitle({ title }: { title: string }) {
  if (!title.includes("\n")) return title;
  const lines = title.split("\n");
  const last = lines[lines.length - 1] ?? "";
  const lead = lines.slice(0, -1).join("\n");
  return (
    <>
      {lead}
      {"\n"}
      <span className="cover-grad">{last}</span>
    </>
  );
}

function Cover({ book, chapters }: { book: Book; chapters: string[] }) {
  const line = isOutline(book) && book.outline
    ? outlineLine(book.outline)
    : isSeries(book) && book.reading
    ? readingLine(book.reading)
    : isGrid(book) && book.grid
      ? gridLine(book.grid)
      : isFacts(book) && book.facts
        ? factLine(book.facts)
        : book.story && isResults(book)
          ? storyLine(book.story, book.period || book.story.period)
          : coverCopy(book.brief);
  const title = book.client || (isOutline(book) ? book.outline?.subject : "") || "Untitled book";
  const long = title.length > 16;
  return (
    <article className="slide cover">
      <CoverWave />
      {book.logo ? (
        <div className="cover-mark">
          <img src={book.logo} alt={`${book.client} mark`} />
        </div>
      ) : null}
      <p className="kicker">{coverKicker(book)}</p>
      <h1 data-long={long ? "" : undefined}>{title}</h1>
      {( [book.market, book.period].filter(Boolean).join(" · ") || (isFacts(book) || isResults(book) || isSeries(book) || isGrid(book) || isOutline(book) ? "" : "Market and period") ) && (
        <p className="meta">
          {[book.market, book.period].filter(Boolean).join(" · ") || "Market and period"}
        </p>
      )}
      <p className={`deck-line${line ? "" : " is-empty"}`}>
        {line || "Add a brief, drop a sheet, or drop an existing deck."}
      </p>
      {chapters.length > 0 && (
        <footer>
          {chapters.map((chapter) => (
            <span key={chapter}>{chapter}</span>
          ))}
        </footer>
      )}
    </article>
  );
}

function BookSlide({ book, view }: { book: Book; view: BookView }) {
  const parts = [
    { label: "Fulfilled", count: view.fulfilled, share: view.fulfilledShare, tone: "fulfilled" as const },
    { label: "Approved, still open", count: view.approvedOpen, share: view.openShare, tone: "approved" as const },
    { label: "Not approved", count: view.notApproved, share: view.notShare, tone: "open" as const },
  ];
  return (
    <article className="slide paper book">
      <header className="book-top">
        <p className="kicker">The book</p>
        <h2>{book.client || "Untitled book"}</h2>
      </header>
      <div className="rate-block">
        <p className="rate">
          {view.rate}
          <span>%</span>
        </p>
        <div>
          <p className="rate-caption">of active subscribers</p>
          <p className="rate-sub">
            {formatCount(view.claims)} claims on {formatCount(view.active)} active
          </p>
        </div>
      </div>
      <Bar rate={view.rate} />
      <p className="bar-note">
        {view.over
          ? "Claims are higher than the active book. The bar stops at the end of the track."
          : "Bar length is the claim rate, and it stops at 100%."}
      </p>
      <ul className="split">
        {parts.map((part) => (
          <li key={part.label}>
            <i className={part.tone} />
            <span>{part.label}</span>
            <strong>{formatCount(part.count)}</strong>
            <em>{part.share}% of active</em>
          </li>
        ))}
      </ul>
    </article>
  );
}

function Mix({ book, view, chart }: { book: Book; view: BookView; chart: MixChart }) {
  const label = view.slices
    .map((slice) => `${slice.label} ${formatCount(slice.count)}, ${slice.percent}% of active`)
    .join(". ");
  const note = view.ringIsClaimMix
    ? "Claims are larger than the active book, so this chart is the claim mix. The rate is still claims divided by active subscribers."
    : "Approved already includes fulfilled, so the open part is what is still waiting.";
  return (
    <article className={`slide mix${chart === "donut" ? "" : " is-flat"}`}>
      <header>
        <p className="kicker">The mix</p>
        <h2>{book.client}</h2>
        <p>{chart === "donut" ? (view.ringIsClaimMix ? "Claims are larger than the active book, so the ring is the claim mix. The center is still the rate of active subscribers." : "The ring is the whole book. Approved already includes fulfilled, so the open slice is what is still waiting.") : note}</p>
      </header>
      {chart === "donut" ? (
        <Donut
          slices={view.slices}
          center={`${view.rate}%`}
          caption="of active subs"
          label={`${book.client}: ${view.rate}% of active subscribers. ${label}`}
        />
      ) : chart === "columns" ? (
        <MixColumns slices={view.slices} label={label} />
      ) : (
        <MixMeter slices={view.slices} rate={view.rate} label={label} />
      )}
      <ul className="key">
        {view.slices.map((slice) => (
          <li key={slice.key}>
            <i className={slice.tone} />
            <span>{slice.label}</span>
            <b>{formatCount(slice.count)}</b>
            <em>{slice.percent}%</em>
          </li>
        ))}
      </ul>
    </article>
  );
}

function MixColumns({ slices, label }: { slices: Slice[]; label: string }) {
  const top = Math.max(1, ...slices.map((slice) => slice.count));
  return (
    <figure className="mix-chart" role="img" aria-label={label}>
      <div className="mix-plot">
        {slices.map((slice) => (
          <span key={slice.key} className="mix-col" style={{ "--h": slice.count / top } as CSSProperties}>
            <b>{slice.percent}%</b>
            <i className={slice.tone} />
          </span>
        ))}
      </div>
      <div className="mix-axis">
        {slices.map((slice) => (
          <span key={slice.key}>{slice.label}</span>
        ))}
      </div>
    </figure>
  );
}

function MixMeter({ slices, rate, label }: { slices: Slice[]; rate: number; label: string }) {
  return (
    <div className="mix-meter" role="img" aria-label={label}>
      <p>
        {rate}
        <span>%</span>
      </p>
      <div>
        <div className="mix-meter-bar">
          {slices.map((slice) =>
            slice.count > 0 ? <span key={slice.key} className={slice.tone} style={{ flexGrow: slice.count }} /> : null,
          )}
        </div>
      </div>
    </div>
  );
}

function Side({
  name,
  rate,
  lead,
}: {
  name: string;
  rate: number;
  lead: boolean;
}) {
  const width = Math.max(rate > 0 ? 2 : 0, Math.min(100, rate));
  return (
    <div className={`v-side${lead ? " is-lead" : ""}`}>
      <strong>{rate}%</strong>
      <span>{name || "Peer"}</span>
      <i aria-hidden>
        <b style={{ width: `${width}%` }} />
      </i>
    </div>
  );
}

function Versus({ book, view }: { book: Book; view: BookView }) {
  const rows = [
    {
      id: "rate",
      label: "Claim rate",
      note: "of active subscribers",
      left: view.rate,
      right: view.peerRate,
    },
    {
      id: "fulfilled",
      label: "Fulfilled",
      note: "of active subscribers",
      left: view.fulfilledShare,
      right: view.peerFulfilledRate,
    },
  ];
  return (
    <article className="slide paper versus">
      <header>
        <p className="kicker">Versus</p>
        <h2>
          {book.client}
          <span> and {book.peerName || "the comparison"}</span>
        </h2>
      </header>
      <ul>
        {rows.map((row) => {
          const lead = row.left === row.right ? "tie" : row.left > row.right ? "left" : "right";
          return (
            <li key={row.id}>
              <p>
                {row.label}
                <small>{row.note}</small>
              </p>
              <div className="v-pair">
                <Side name={book.client} rate={row.left} lead={lead !== "right"} />
                <Side name={book.peerName} rate={row.right} lead={lead !== "left"} />
              </div>
            </li>
          );
        })}
      </ul>
      <p className="fine">The larger figure is marked. A tie marks both. Bar length is the rate, and it stops at 100%.</p>
    </article>
  );
}

function FactSlide({ book, index }: { book: Book; index: number }) {
  const group = book.facts?.groups[index];
  if (!group) return null;
  const volume = group.volume;
  const rates = group.rates.slice(0, 4);
  const floor = volume && volume.relation !== "exact" ? (volume.relation === "least" ? "at least" : volume.relation) : "";
  const when = [volume?.hedge, floor, book.facts?.verb, book.period].filter(Boolean).join(" · ");
  const separate = (book.facts?.groups.length ?? 0) > 1;
  const hedged = rates.some((rate) => rate.relation !== "exact");
  return (
    <article className="slide paper result">
      <header className="book-top">
        <p className="kicker">{group.name || "The result"}</p>
        <h2>{book.client || "Untitled"}</h2>
      </header>
      {volume && volume.value > 0 && (
        <div className="rate-block">
          <p className="rate">{compactVolume(volume.value)}</p>
          <div>
            <p className="rate-caption">{volume.unit || "Count"}</p>
            {when && <p className="rate-sub">{when}</p>}
          </div>
        </div>
      )}
      {rates.length > 0 && (
        <ul className={`proofs proofs-${Math.min(rates.length, 4)}`}>
          {rates.map((rate) => (
            <li key={`${rate.label}-${rate.value}`}>
              <span>{rate.label}</span>
              <strong>{rate.hedge ? `${rate.hedge} ` : ""}{proofPhrase(rate)}</strong>
              <div className="meter" aria-hidden>
                <span style={{ width: `${Math.max(rate.value > 0 ? 2 : 0, Math.min(100, rate.value))}%` }} />
              </div>
            </li>
          ))}
        </ul>
      )}
      <p className="bar-note">
        {separate
          ? "This slide is only this series. Calls and chats are not added together."
          : hedged
            ? "Each bar stops at the figure in the brief. Above means at least that high."
            : "Each bar is the rate stated in the brief."}
      </p>
    </article>
  );
}

function Result({ book }: { book: Book }) {
  const story = book.story;
  if (!story) return null;
  const proofs = story.proofs.slice(0, 4);
  const hedged = proofs.some((proof) => proof.relation !== "exact");
  const when = [story.hedge, story.verb, book.period || story.period].filter(Boolean).join(" · ");
  return (
    <article className="slide paper result">
      <header className="book-top">
        <p className="kicker">The result</p>
        <h2>{book.client || story.client || "Untitled"}</h2>
      </header>
      {story.volume != null && story.volume > 0 && (
        <div className="rate-block">
          <p className="rate">{compactVolume(story.volume)}</p>
          <div>
            <p className="rate-caption">{story.unit || "Count"}</p>
            {when && <p className="rate-sub">{when}</p>}
          </div>
        </div>
      )}
      {proofs.length > 0 && (
        <ul className={`proofs proofs-${Math.min(proofs.length, 4)}`}>
          {proofs.map((proof) => (
            <li key={`${proof.label}-${proof.value}`}>
              <span>{proof.label}</span>
              <strong>{proofPhrase(proof)}</strong>
              <div className="meter" aria-hidden>
                <span style={{ width: `${Math.max(proof.value > 0 ? 2 : 0, Math.min(100, proof.value))}%` }} />
              </div>
            </li>
          ))}
        </ul>
      )}
      <p className="bar-note">
        {proofs.length === 0
          ? "The figure is the one written in the brief."
          : hedged
            ? "Each bar stops at the figure in the brief. Above means at least that high."
            : "Each bar is the rate stated in the brief."}
      </p>
    </article>
  );
}

function ScopeSlide({ book }: { book: Book }) {
  const cards = book.reading?.scope.slice(0, 4) ?? [];
  return (
    <article className="slide paper scope">
      <header className="book-top">
        <p className="kicker">What counts</p>
        <h2>{book.client || book.reading?.title || "The count"}</h2>
      </header>
      <ul className={`scope-list scope-${Math.min(cards.length, 4)}`}>
        {cards.map((card) => (
          <li key={card.label}>
            <span>{card.label}</span>
            <p data-long={card.text.length > 180 ? "" : undefined}>{card.text}</p>
          </li>
        ))}
      </ul>
    </article>
  );
}

function PointLine({
  points,
  max,
  label,
}: {
  points: { label: string; value: number; projected?: boolean }[];
  max: number;
  label: string;
}) {
  const width = 640;
  const height = 220;
  const pad = 18;
  const count = points.length;
  const ceiling = max > 0 ? max : 1;
  const xOf = (index: number) => (count <= 1 ? width / 2 : pad + (index * (width - pad * 2)) / (count - 1));
  const yOf = (value: number) => height - 16 - (Math.max(0, value) / ceiling) * (height - 36);
  const path = points.map((point, index) => `${index === 0 ? "M" : "L"}${xOf(index).toFixed(1)} ${yOf(point.value).toFixed(1)}`).join(" ");
  const step = count > 14 ? 2 : 1;
  return (
    <div className="trend-line-wrap">
      <svg className="trend-line" viewBox={`0 0 ${width} ${height}`} role="img" aria-label={label}>
        <line x1={pad} y1={yOf(ceiling / 2)} x2={width - pad} y2={yOf(ceiling / 2)} />
        <path d={path} />
        {points.map((point, index) => (
          <circle key={`${point.label}-${index}`} className={point.projected ? "is-proj" : ""} cx={xOf(index)} cy={yOf(point.value)} r={count > 16 ? 3 : 4.5} />
        ))}
      </svg>
      <div className="trend-line-labels">
        {points.map((point, index) => (
          <span key={`${point.label}-label-${index}`}>{index % step === 0 || index === count - 1 ? point.label : ""}</span>
        ))}
      </div>
    </div>
  );
}

function Trend({ book, index }: { book: Book; index: number }) {
  const track = book.reading?.tracks[index];
  if (!track) return null;
  const totals = trackTotals(track);
  const points = track.points.map((point) => ({ label: point.label, value: point.rate }));
  return (
    <article className="slide paper trend">
      <header className="trend-top">
        <div>
          <p className="kicker">{track.name}</p>
          <h2>Containment</h2>
        </div>
        <div className="trend-rate">
          <p className="rate">
            {totals.rate}
            <span>%</span>
          </p>
          <p className="rate-sub">
            {formatCount(totals.part)} {track.partLabel.toLowerCase()} of {formatCount(totals.total)} {track.unit}
          </p>
        </div>
      </header>
      <div className="trend-chart is-line">
        <div className="trend-axis" aria-hidden>
          <span>100</span>
          <span>50</span>
          <span>0</span>
        </div>
        <div className="trend-body">
          <PointLine points={points} max={100} label={trackNote(track)} />
        </div>
      </div>
      <p className="bar-note">{trackNote(track)}</p>
    </article>
  );
}

function Compare({ book }: { book: Book }) {
  const tracks = book.reading?.tracks.slice(0, 3) ?? [];
  return (
    <article className="slide paper compare">
      <header className="book-top">
        <p className="kicker">Side by side</p>
        <h2>{book.period || book.client || "The series"}</h2>
      </header>
      <ul className={`compare-list compare-${tracks.length}`}>
        {tracks.map((track) => {
          const totals = trackTotals(track);
          return (
            <li key={track.name}>
              <span>{track.name}</span>
              <strong>
                {totals.rate}
                <i>%</i>
              </strong>
              <em>
                {formatCount(totals.part)} of {formatCount(totals.total)} {track.unit}
              </em>
              <div className="meter" aria-hidden>
                <span style={{ width: `${Math.max(totals.rate > 0 ? 2 : 0, Math.min(100, totals.rate))}%` }} />
              </div>
            </li>
          );
        })}
      </ul>
      <p className="bar-note">Each rate is that series alone. Calls and chats stay in their own totals.</p>
    </article>
  );
}

function Next({ book, view }: { book: Book; view: BookView }) {
  const text = book.recommendation.trim();
  const size = text.length > 180 ? "s" : text.length > 110 ? "m" : "l";
  const seriesLead = isSeries(book) ? book.reading?.tracks[0] : undefined;
  const seriesTotals = seriesLead ? trackTotals(seriesLead) : null;
  const factVolume = isFacts(book) ? book.facts?.groups.find((group) => (group.volume?.value ?? 0) > 0)?.volume : undefined;
  const results = !seriesTotals && !factVolume && isResults(book) && book.story?.volume != null && book.story.volume > 0;
  const figure = seriesTotals ? String(seriesTotals.rate) : factVolume ? compactVolume(factVolume.value) : results ? compactVolume(book.story?.volume ?? 0) : String(view.rate);
  return (
    <article className="slide next">
      <div className="next-panel">
        <p className="kicker">What to do</p>
        <p className="who">{book.client || "This book"}</p>
        <p className="next-rate">
          {figure}
          {!results && !factVolume && <span>%</span>}
        </p>
        <p className="next-rate-note">
          {seriesTotals ? `${seriesLead?.name ?? "Series"} containment` : factVolume ? factVolume.unit || "Count" : results ? book.story?.unit || "Count" : "of active subscribers"}
        </p>
        <p className="meta">{[book.market, book.period].filter(Boolean).join(" · ")}</p>
      </div>
      <div className="next-line">
        <p data-size={size} className={text ? "" : "is-empty"}>
          {text || "Write the one line. The deck will not invent a recommendation."}
        </p>
      </div>
    </article>
  );
}

function Listening({ book, chart }: { book: Book; chart: ListeningChart }) {
  const talk = conversationOf(book.listening);
  const max = Math.max(1, ...book.listening.map((row) => row.count));
  return (
    <article className="slide listening">
      <div className="listen-hero">
        <p className="kicker">Health</p>
        <p className="listen-period">{book.period || book.client || "This book"}</p>
        <p className="health">
          {talk.health}
          <span>%</span>
        </p>
        <p className="health-note">of mentions are positive or neutral.</p>
        <p className="listen-count">
          {formatCount(talk.total)} {talk.total === 1 ? "mention" : "mentions"}
        </p>
        <SentimentMeter
          negative={talk.negative}
          neutral={talk.neutral}
          positive={talk.positive}
          negativeShare={talk.negativeShare}
          neutralShare={talk.neutralShare}
          positiveShare={talk.positiveShare}
        />
      </div>
      <div className="listen-sheet">
        <header>
          <p className="kicker">Listening</p>
          <h2>Who is talking</h2>
        </header>
        {chart === "columns" ? (
          <ListenColumns rows={book.listening} />
        ) : (
          <ul className="srows">
            {book.listening.map((row) => (
              <SentimentRow key={row.label} row={row} max={max} total={talk.total} />
            ))}
          </ul>
        )}
        <p className="fine">
          {chart === "columns"
            ? "Each column is one topic. Negative sits at the base. No topic is added."
            : "Row length is the loudest topic. Colour runs negative, neutral, then positive."}
        </p>
      </div>
    </article>
  );
}

function SentimentMeter({
  negative,
  neutral,
  positive,
  negativeShare,
  neutralShare,
  positiveShare,
}: {
  negative: number;
  neutral: number;
  positive: number;
  negativeShare: number;
  neutralShare: number;
  positiveShare: number;
}) {
  return (
    <div
      className="sent-meter"
      role="img"
      aria-label={`${negativeShare}% negative, ${neutralShare}% neutral, ${positiveShare}% positive`}
    >
      <div className="sent-meter-bar">
        {negative > 0 && <span className="neg" style={{ flexGrow: negative }} />}
        {neutral > 0 && <span className="neu" style={{ flexGrow: neutral }} />}
        {positive > 0 && <span className="pos" style={{ flexGrow: positive }} />}
      </div>
      <ul>
        <li className="neg">
          <strong>{negativeShare}%</strong> negative
        </li>
        <li className="neu">
          <strong>{neutralShare}%</strong> neutral
        </li>
        <li className="pos">
          <strong>{positiveShare}%</strong> positive
        </li>
      </ul>
    </div>
  );
}

function SentimentRow({ row, max, total }: { row: ListeningRow; max: number; total: number }) {
  const share = total > 0 ? Math.round((row.count / total) * 100) : 0;
  return (
    <li className="srow">
      <span className="srow-label">{row.label}</span>
      <span className="srow-count">{formatCount(row.count)}</span>
      <span className="srow-share">{share}%</span>
      <span className="srow-track">
        <span className="srow-fill" style={{ width: `${Math.max(1.5, (row.count / max) * 100)}%` }}>
          {row.negative > 0 && <i className="neg" style={{ flexGrow: row.negative }} />}
          {row.neutral > 0 && <i className="neu" style={{ flexGrow: row.neutral }} />}
          {row.positive > 0 && <i className="pos" style={{ flexGrow: row.positive }} />}
        </span>
      </span>
    </li>
  );
}

function Journey({ book }: { book: Book }) {
  return (
    <article className="slide paper journey">
      <header>
        <p className="kicker">Journey</p>
        <h2>Where the tone turns</h2>
      </header>
      <div className="j-scale" aria-hidden>
        <span />
        <span className="j-scale-track">
          <span>Negative</span>
          <span>Positive</span>
        </span>
        <span />
      </div>
      <ul>
        {book.journey.map((stage) => (
          <JourneyRow key={stage.label} stage={stage} />
        ))}
      </ul>
      <p className="fine">Percents are the sheet's own split. Neutral sits against the center line.</p>
    </article>
  );
}

function JourneyRow({ stage }: { stage: JourneyStage }) {
  const halfNeutral = stage.neutral / 2;
  return (
    <li className="jrow">
      <span className="jrow-name">
        {stage.label}
        <small>
          {formatCount(stage.mentions)} {stage.mentions === 1 ? "mention" : "mentions"}
        </small>
      </span>
      <span className="j-chart" aria-hidden>
        <span className="j-half is-left">
          <i className="neu" style={{ width: `${halfNeutral}%` }} />
          <i className="neg" style={{ width: `${stage.negative}%` }} />
        </span>
        <span className="j-half is-right">
          <i className="neu" style={{ width: `${halfNeutral}%` }} />
          <i className="pos" style={{ width: `${stage.positive}%` }} />
        </span>
      </span>
      <span className="jrow-figures">
        <span className="neg">{stage.negative}%</span>
        <span className="pos">{stage.positive}%</span>
      </span>
    </li>
  );
}

function Months({ book }: { book: Book }) {
  const stacks = book.months.map((month) => ({
    ...month,
    total: month.fulfilled + month.open + month.notApproved,
  }));
  const top = Math.max(1, ...stacks.map((month) => month.total));
  const caption = "Each column is that month's claims. The tallest month sets the scale.";
  return (
    <article className="slide paper months">
      <header className="book-top">
        <p className="kicker">Months</p>
        <h2>{book.client || "Claims by month"}</h2>
      </header>
      <figure className="months-chart">
        <div className="months-plot" role="img" aria-label={caption}>
          <div className="months-grid" aria-hidden>
            <span style={{ bottom: "100%" }}>
              <em>{formatCount(top)}</em>
            </span>
            <span style={{ bottom: "50%" }}>
              <em>{formatCount(Math.round(top / 2))}</em>
            </span>
            <span style={{ bottom: 0 }} />
          </div>
          <div className="months-cols" aria-hidden>
            {stacks.map((month) => (
              <span key={month.label} className="month-col" style={{ "--h": month.total / top } as CSSProperties}>
                {month.fulfilled > 0 && <i className="fulfilled" style={{ flexGrow: month.fulfilled }} />}
                {month.open > 0 && <i className="approved" style={{ flexGrow: month.open }} />}
                {month.notApproved > 0 && <i className="open" style={{ flexGrow: month.notApproved }} />}
              </span>
            ))}
          </div>
        </div>
        <div className="months-axis" aria-hidden>
          {stacks.map((month) => (
            <span key={month.label}>{month.label}</span>
          ))}
        </div>
        <figcaption>{caption}</figcaption>
      </figure>
      <ul className="month-key">
        <li className="fulfilled">Fulfilled</li>
        <li className="approved">Approved, still open</li>
        <li className="open">Not approved</li>
      </ul>
    </article>
  );
}

function ListenColumns({ rows }: { rows: Book["listening"] }) {
  const top = Math.max(1, ...rows.map((row) => row.count));
  return (
    <figure className="listen-cols">
      <div className="listen-plot" aria-hidden>
        {rows.map((row) => (
          <span key={row.label} className="month-col" style={{ "--h": row.count / top } as CSSProperties}>
            {row.positive > 0 && <i className="pos" style={{ flexGrow: row.positive }} />}
            {row.neutral > 0 && <i className="neu" style={{ flexGrow: row.neutral }} />}
            {row.negative > 0 && <i className="neg" style={{ flexGrow: row.negative }} />}
          </span>
        ))}
      </div>
      <div className="listen-axis" aria-hidden>
        {rows.map((row) => (
          <span key={row.label}>{row.label}</span>
        ))}
      </div>
    </figure>
  );
}

function BroughtIn({ book }: { book: Book }) {
  const text = book.sourceText.trim();
  const shown = text.length > 1600 ? `${text.slice(0, 1600).trimEnd()}…` : text;
  const fit = shown.length > 700 ? "s" : shown.length > 380 ? "m" : "l";
  return (
    <article className="slide paper brought">
      <header>
        <p className="kicker">Brought in</p>
        <h2>{book.sourceName || "Decoded text"}</h2>
      </header>
      <p className="brought-body" data-fit={fit}>
        {shown || "Nothing was decoded from a document."}
      </p>
    </article>
  );
}

function originSlide(book: Book, index: number): SourceSlide | null {
  return book.sourceDeck?.slides.find((slide) => slide.index === index) ?? book.sourceDeck?.slides[index - 1] ?? null;
}

function OriginLegend({ chart }: { chart: SourceChart }) {
  if (chart.series.length < 2) return null;
  return (
    <ul className="origin-key">
      {chart.series.slice(0, 4).map((series, index) => (
        <li key={`${series.name}-${index}`}>
          <i className={`s${index}`} />
          {series.name}
        </li>
      ))}
    </ul>
  );
}

function cellText(table: SourceTable | null, category: string, series: { name: string; values: number[] }, value: number) {
  return rowDisplay(table, category, [series], formatPoint(value));
}

function leadAt(categories: string[], values: number[]): number {
  const total = categories.findIndex((category) => /^total\b/i.test(category.trim()));
  if (total >= 0) return total;
  return values.reduce((best, value, index) => (value > (values[best] ?? 0) ? index : best), 0);
}

function ScoreChart({ chart, table }: { chart: SourceChart; table: SourceTable | null }) {
  const series = chart.series[0];
  if (!series) return null;
  const cats = chart.categories.slice(0, 8);
  const values = cats.map((_, index) => series.values[index] ?? 0);
  const max = Math.max(1, ...values.map((value) => Math.max(0, value)));
  const lead = leadAt(cats, values);
  const rest = cats.map((_, index) => index).filter((index) => index !== lead);
  const hero = cellText(table, cats[lead] ?? "", series, values[lead] ?? 0);
  const wide = rest.length > 3;
  return (
    <div className={`score-board${rest.length ? " has-rest" : ""}${wide ? " is-wide" : ""}`}>
      <div className="score-hero">
        <span className="score-name">{cats[lead]}</span>
        <b aria-label={hero} style={{ fontSize: displayPx(hero, wide ? 84 : 128) }}><Figure value={hero} /></b>
        <i className="score-beam"><em style={{ width: `${((Math.max(0, values[lead] ?? 0)) / max) * 100}%` }} /></i>
      </div>
      {rest.length > 0 ? (
        <div className={`score-constellation rest-${rest.length}`}>
          {rest.map((index) => {
            const value = values[index] ?? 0;
            const written = cellText(table, cats[index] ?? "", series, value);
            return (
              <div className="score-star" key={`${cats[index]}-${index}`}>
                <b aria-label={written} style={{ fontSize: displayPx(written, wide ? 22 : 32) }}><Figure value={written} /></b>
                <i className={value <= 0 ? "is-flat" : undefined} style={{ ["--h" as string]: value <= 0 ? 0 : Math.max(0, value) / max }} />
                <span>{cats[index]}</span>
              </div>
            );
          })}
        </div>
      ) : null}
    </div>
  );
}

function RankChart({ chart, table }: { chart: SourceChart; table: SourceTable | null }) {
  const series = chart.series[0];
  if (!series) return null;
  const cats = chart.categories.slice(0, 14);
  const values = cats.map((_, index) => series.values[index] ?? 0);
  const max = Math.max(1, ...values.map((value) => Math.max(0, value)));
  const lead = leadAt(cats, values);
  const rest = cats.map((_, index) => index).filter((index) => index !== lead);
  const hero = cellText(table, cats[lead] ?? "", series, values[lead] ?? 0);
  return (
    <div className="rank-signal" role="img" aria-label={chartSentence(chart)}>
      <div className="rank-beacon">
        <span>{cats[lead]}</span>
        <b aria-label={hero} style={{ fontSize: displayPx(hero, 72) }}><Figure value={hero} /></b>
      </div>
      <ol className="rank-chart">
        {rest.map((index) => {
          const value = values[index] ?? 0;
          const written = cellText(table, cats[index] ?? "", series, value);
          return (
            <li key={`${cats[index]}-${index}`}>
              <span>{cats[index]}</span>
              <i><em style={{ width: value <= 0 ? "0%" : `${(Math.max(0, value) / max) * 100}%` }} /></i>
              <b aria-label={written}><Figure value={written} /></b>
            </li>
          );
        })}
      </ol>
    </div>
  );
}

function MatrixChart({ chart, table }: { chart: SourceChart; table: SourceTable | null }) {
  const cats = chart.categories.slice(0, 12);
  const series = chart.series.slice(0, 4).map((item) => ({ ...item, values: cats.map((_, index) => item.values[index] ?? 0) }));
  const spare = spareHeaders(table, series);
  const maxes = series.map((item) => Math.max(1, ...item.values.map((value) => Math.max(0, value))));
  const columns = `minmax(148px, 1.15fr) repeat(${series.length + spare.length}, minmax(0, 1fr))`;
  return (
    <div className="matrix" role="img" aria-label={chartSentence(chart)}>
      <p>Each column has its own scale. The figure is the cell as written.</p>
      <div className="matrix-row is-head" style={{ gridTemplateColumns: columns }}>
        <span />
        {series.map((item) => <span key={item.name}>{item.name}</span>)}
        {spare.map((header) => <span key={header}>{header}</span>)}
      </div>
      {cats.map((cat, index) => {
        const row = table?.rows.find((item) => (item[0] ?? "").trim() === cat);
        return (
          <div className="matrix-row" key={`${cat}-${index}`} style={{ gridTemplateColumns: columns }}>
            <span className="matrix-name">{cat}</span>
            {series.map((item, seriesIndex) => {
              const peak = item.values.reduce((best, value, valueIndex) => (value > (item.values[best] ?? 0) ? valueIndex : best), 0);
              return (
                <span className="matrix-cell" key={item.name}>
                  <i><em className={`s${seriesIndex}`} style={{ width: `${(Math.max(0, item.values[index] ?? 0) / (maxes[seriesIndex] ?? 1)) * 100}%` }} /></i>
                  <b className={index === peak && (item.values[index] ?? 0) > 0 ? "is-peak" : undefined}>{cellText(table, cat, item, item.values[index] ?? 0)}</b>
                </span>
              );
            })}
            {spare.map((header) => {
              const column = table?.headers.findIndex((item) => item.trim() === header) ?? -1;
              return <b className="matrix-spare" key={header}>{column >= 0 ? row?.[column] ?? "" : ""}</b>;
            })}
          </div>
        );
      })}
    </div>
  );
}

function TrendChart({ chart }: { chart: SourceChart }) {
  const cats = chart.categories.slice(0, 16);
  const series = chart.series.slice(0, 4).map((item) => ({ ...item, values: cats.map((_, index) => item.values[index] ?? 0) }));
  const span = axisSpan(series.flatMap((item) => item.values));
  const width = 760;
  const height = 340;
  const padL = 58;
  const padR = 16;
  const padT = 24;
  const padB = 40;
  const innerW = width - padL - padR;
  const innerH = height - padT - padB;
  const xOf = (index: number) => (cats.length <= 1 ? padL + innerW / 2 : padL + (innerW * index) / (cats.length - 1));
  const yOf = (value: number) => padT + innerH - ((Math.max(span.min, value) - span.min) / (span.max - span.min || 1)) * innerH;
  const ticks = [span.max, span.min + (span.max - span.min) / 2, span.min];
  return (
    <figure className={`origin-trend${cats.length > 8 ? " is-dense" : ""}`}>
      <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label={chartSentence(chart)}>
        {ticks.map((tick) => (
          <g key={tick}>
            <line x1={padL} x2={width - padR} y1={yOf(tick)} y2={yOf(tick)} />
            <text className="axis" x={padL - 8} y={yOf(tick) + 4} textAnchor="end">{formatPoint(tick)}</text>
          </g>
        ))}
        {series.map((item, index) => {
          const line = item.values.map((value, point) => `${point === 0 ? "M" : "L"}${xOf(point).toFixed(1)} ${yOf(value).toFixed(1)}`).join(" ");
          const peak = item.values.reduce((best, value, point) => (value > (item.values[best] ?? -Infinity) ? point : best), 0);
          return (
            <g key={`${item.name}-${index}`} className={`s${index}`}>
              <path className="glow" d={line} />
              <path className="stroke" d={line} />
              {item.values.map((value, point) => <circle key={point} cx={xOf(point)} cy={yOf(value)} r={point === peak ? 8 : 5} />)}
              {series.length === 1 && cats.length <= 8 ? null : (
                <text className="point" x={xOf(peak)} y={yOf(item.values[peak] ?? 0) - 16} textAnchor="middle">{formatPoint(item.values[peak] ?? 0)}</text>
              )}
            </g>
          );
        })}
        {series.length === 1 && cats.length <= 8
          ? series[0]!.values.map((value, point) => (
            <text key={`value-${point}`} className="point" x={xOf(point)} y={yOf(value) - 18} textAnchor="middle">{formatPoint(value)}</text>
          ))
          : null}
        {cats.map((cat, index) => (
          <text key={`${cat}-${index}`} className="cat" x={xOf(index)} y={height - 12} textAnchor="middle">{cat}</text>
        ))}
      </svg>
    </figure>
  );
}

function OriginColumns({ chart, stacked }: { chart: SourceChart; stacked: boolean }) {
  const cats = chart.categories.slice(0, 12);
  const series = chart.series.slice(0, 4).map((item) => ({ ...item, values: item.values.slice(0, cats.length) }));
  const peaks = cats.map((_, index) => stacked
    ? series.reduce((total, item) => total + Math.max(0, item.values[index] ?? 0), 0)
    : Math.max(0, ...series.map((item) => item.values[index] ?? 0)));
  const max = axisMax(Math.max(0, ...peaks));
  const only = series.length === 1 ? series[0] : null;
  const peakIndex = only
    ? only.values.reduce((best, value, index) => (value > (only.values[best] ?? -Infinity) ? index : best), 0)
    : -1;
  return (
    <div className="origin-chart">
      <div className="origin-axis" aria-hidden>
        <span>{formatPoint(max)}</span>
        <span>{formatPoint(max / 2)}</span>
        <span>0</span>
      </div>
      <div className="origin-plot" role="img" aria-label={chartSentence(chart)}>
        {cats.map((cat, index) => (
          <div className="origin-cat" key={`${cat}-${index}`}>
            {stacked ? (
              <div className="origin-stack">
                {series.map((item, seriesIndex) => (
                  <i
                    key={`${item.name}-${seriesIndex}`}
                    className={`s${seriesIndex}`}
                    style={{ ["--h" as string]: Math.max(0, item.values[index] ?? 0) / max }}
                  />
                ))}
              </div>
            ) : (
              <div className="origin-bars">
                {series.map((item, seriesIndex) => {
                  const value = item.values[index] ?? 0;
                  const printed = formatPoint(value);
                  return (
                    <i key={`${item.name}-${seriesIndex}`} className={`s${seriesIndex}${only && index === peakIndex ? " is-peak" : ""}`} style={{ ["--h" as string]: Math.max(0, value) / max }}>
                      {printed.length <= 8 ? <em>{printed}</em> : null}
                    </i>
                  );
                })}
              </div>
            )}
            <b>{cat}</b>
          </div>
        ))}
      </div>
    </div>
  );
}

function OriginDonut({ chart }: { chart: SourceChart }) {
  const series = chart.series[0];
  const values = series?.values ?? [];
  const sum = values.reduce((total, value) => total + Math.max(0, value), 0) || 1;
  const tones: Tone[] = ["fulfilled", "approved", "open", "rest"];
  const slices: Slice[] = values.map((value, index) => ({
    key: `${index}`,
    label: chart.categories[index] || `Part ${index + 1}`,
    count: Math.max(0, value),
    percent: Math.round((Math.max(0, value) / sum) * 100),
    tone: tones[index % tones.length]!,
  })).filter((slice) => slice.count > 0);
  const top = [...slices].sort((left, right) => right.count - left.count)[0];
  return (
    <div className="origin-donut">
      <Donut slices={slices} center={top ? `${top.percent}%` : "0%"} caption={top?.label ?? series?.name ?? ""} label={chartSentence(chart)} />
      <ul className="origin-legend">
        {slices.map((slice) => (
          <li key={slice.key}>
            <i className={TONE[slice.tone]} />
            <span>{slice.label}</span>
            <b>{formatPoint(slice.count)}</b>
          </li>
        ))}
      </ul>
    </div>
  );
}

function KeptImages({ images, label, stage = false }: { images: SourceSlide["images"]; label: string; stage?: boolean }) {
  if (images.length === 0) return null;
  const shown = images.slice(0, 4);
  const layout = pictureLayout(shown);
  return (
    <div className={`kept-images kept-${shown.length}${stage ? ` is-${layout}` : ""}`}>
      {shown.map((image, index) => {
        if (!stage) return <img key={`${image.mime}-${index}`} src={image.dataUrl} alt={label} />;
        const size = imageSize(image.dataUrl);
        return (
          <div className="picture-cell" key={`${image.mime}-${index}`}>
            <div className="picture-fit" style={size ? { ["--r" as string]: String(size.width / size.height) } : undefined}>
              <img src={image.dataUrl} alt={label} />
            </div>
          </div>
        );
      })}
    </div>
  );
}

function TableCard({ table }: { table: SourceTable }) {
  const dense = table.rows.length > 6;
  const textual = !tableHasNumbers(table);
  return (
    <div className={`source-table${dense ? " is-dense" : ""}${textual ? " is-text" : ""}`}>
      <p>As written</p>
      <table>
        <thead>
          <tr>
            {table.headers.map((header, index) => (
              <th key={`${header}-${index}`}>{header}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {table.rows.map((row, index) => (
            <tr key={index}>
              {table.headers.map((_, cell) => (
                <td key={cell}>{row[cell] ?? ""}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function OriginFunnel({ chart, table }: { chart: SourceChart; table: SourceTable | null }) {
  const series = chart.series[0];
  const values = series?.values ?? [];
  const max = Math.max(1, ...values.map((value) => Math.max(0, value)));
  return (
    <ol className="funnel" role="img" aria-label={chartSentence(chart)}>
      {chart.categories.map((category, index) => {
        const value = values[index] ?? 0;
        const width = Math.max(34, (Math.max(0, value) / max) * 100);
        return (
          <li key={`${category}-${index}`}>
            <div style={{ width: `${width}%` }}>
              <span>{category}</span>
              <b>{rowDisplay(table, category, series ? [series] : [], formatPoint(value))}</b>
            </div>
          </li>
        );
      })}
    </ol>
  );
}

function ChartPlot({ chart, table }: { chart: SourceChart; table: SourceTable | null }) {
  const mode = plotMode(chart);
  if (mode === "funnel") return <OriginFunnel chart={chart} table={table} />;
  if (mode === "donut") return <OriginDonut chart={chart} />;
  if (mode === "score") return <ScoreChart chart={chart} table={table} />;
  if (mode === "rank") return <RankChart chart={chart} table={table} />;
  if (mode === "matrix") return <MatrixChart chart={chart} table={table} />;
  if (mode === "trend") return <TrendChart chart={chart} />;
  return <OriginColumns chart={chart} stacked={mode === "stacked"} />;
}

function OriginChart({ slide }: { slide: SourceSlide }) {
  const chart = slide.chart;
  if (!chart) return null;
  const mode = plotMode(chart);
  const inline = mode === "score" || mode === "rank" || mode === "matrix";
  const limit = mode === "rank" ? 14 : mode === "trend" || mode === "funnel" ? 16 : mode === "matrix" ? 12 : mode === "score" ? 8 : 12;
  const hidden = chart.categories.length > limit;
  const eyebrow = chartEyebrow(chart);
  const note = [
    chart.series.length > 1 && mode === "donut" ? "The first series is drawn." : "",
    hidden ? `First ${limit} of ${chart.categories.length} points.` : "",
    !inline && slide.table && slide.table.rows.length > chart.categories.length ? "Every row stays in the table." : "",
    slide.extraCharts > 0 ? "Another chart is on this slide and is not drawn." : "",
    chartFoot(chart),
  ].filter(Boolean).join(" ");
  const showLegend = mode === "trend" || mode === "columns" || mode === "stacked";
  const beside = slideShow(slide) === "chart" ? null : slide.table;
  return (
    <article className={`slide story is-${mode}${!inline && beside ? " has-table" : ""}`}>
      <header className="origin-top">
        <div>
          {eyebrow ? <p className="kicker">{eyebrow}</p> : null}
          <h2>{slide.title || "Chart"}</h2>
        </div>
        {showLegend ? <OriginLegend chart={chart} /> : null}
      </header>
      <CopyBlocks text={slide.context} tone="story" />
      <div className="story-stage">
        <div className={`story-plot${slide.images.length ? " has-kept" : ""}`}>
          <ChartPlot chart={chart} table={slide.table} />
          <KeptImages images={slide.images} label={slide.title || "Chart picture"} />
        </div>
        {!inline && beside ? <TableCard table={beside} /> : null}
      </div>
      {note && <p className="bar-note">{note}</p>}
    </article>
  );
}

function CopyBlocks({ text, tone }: { text: string; tone: "story" | "figure" }) {
  const blocks = copyBlocks(text);
  if (!blocks.length) return null;
  return (
    <div className={`copy-blocks is-${tone}`}>
      {blocks.map((block, index) => block.kind === "list" ? (
        <ul className={tone === "story" ? "story-points" : "bullet-list is-tight"} key={index}>
          {block.points.map((line, point) => (
            <li key={`${line}-${point}`}><i aria-hidden /><span>{line}</span></li>
          ))}
        </ul>
      ) : (
        <p className={tone === "story" ? "story-line" : "origin-context"} key={index}>{block.text}</p>
      ))}
    </div>
  );
}

function OriginNumbers({ slide }: { slide: SourceSlide }) {
  const numbers = slide.numbers.slice(0, 6);
  const body = figureBody(slide).trim();
  const busy = copyBlocks(body).some((block) => block.kind === "list") || body.length > 180;
  if (numbers.length === 1) {
    const number = numbers[0]!;
    const wide = number.display.length > 8;
    return (
      <article className={`slide paper origin${busy ? " has-copy" : ""}`}>
        <header className="book-top">
          <p className="kicker">Figure</p>
          <h2>{slide.title || number.label || "Figure"}</h2>
        </header>
        <div className="rate-block">
          <p className="rate" data-wide={wide ? "" : undefined}>{number.display}</p>
          {number.label && <p className="rate-caption">{number.label}</p>}
        </div>
        {body ? <CopyBlocks text={body} tone="figure" /> : null}
        {slide.images.length ? (
          <div className="picture-stage">
            <KeptImages images={slide.images} label={slide.title || "Picture"} stage />
          </div>
        ) : null}
      </article>
    );
  }
  const columns = numbers.length === 2 || numbers.length === 4 ? numbers.length : numbers.length === 3 ? 3 : 3;
  return (
    <article className={`slide paper origin${busy ? " has-copy" : ""}`}>
      <header className="book-top">
        <p className="kicker">Figures</p>
        <h2>{slide.title || "Figures"}</h2>
      </header>
      <ul className={`figure-grid figure-${columns}`}>
        {numbers.map((number, index) => (
          <li key={`${number.display}-${index}`}>
            <strong data-wide={number.display.length > 8 ? "" : undefined}>{number.display}</strong>
            {number.label && <span>{number.label}</span>}
          </li>
        ))}
      </ul>
      {body ? <CopyBlocks text={body} tone="figure" /> : null}
      {slide.images.length ? (
        <div className="picture-stage">
          <KeptImages images={slide.images} label={slide.title || "Picture"} stage />
        </div>
      ) : null}
    </article>
  );
}

function ProgressSlide({ book }: { book: Book }) {
  if (book.grid) {
    const view = gridProgress(book.grid);
    if (!view) return null;
    const share = view.kind === "share";
    return (
      <article className="slide paper compare">
        <header className="book-top">
          <p className="kicker">{view.title}</p>
          <h2>{view.column}</h2>
        </header>
        <ul className="compare-list compare-2">
          <li>
            <span>{share ? "Progress" : view.leftLabel}</span>
            <strong>
              {formatGridPoint(view.left)}
              {share ? <i>%</i> : null}
            </strong>
            <em>{share ? "Latest row" : "Previous row"}</em>
          </li>
          <li>
            <span>{share ? "Not progress" : view.rightLabel}</span>
            <strong>
              {formatGridPoint(view.right)}
              {share ? <i>%</i> : null}
            </strong>
            <em>{share ? "The rest of 100" : "Latest row"}</em>
          </li>
        </ul>
        <p className="bar-note">{view.note}</p>
      </article>
    );
  }
  const cards = book.reading ? seriesProgressList(book.reading) : [];
  if (cards.length === 0) return null;
  return (
    <article className="slide paper compare">
      <header className="book-top">
        <p className="kicker">{cards.length === 1 ? cards[0]!.direction : "Progress"}</p>
        <h2>Latest full month</h2>
      </header>
      <div className="progress-board" data-count={cards.length}>
        {cards.map((card) => (
          <div key={card.name}>
            <span>{card.name}</span>
            <p>
              {card.direction} · {card.label}
            </p>
            <div>
              <b>{formatCount(card.contained)}</b>
              <small>Contained</small>
            </div>
            <div>
              <b>{formatCount(card.rest)}</b>
              <small>Not contained</small>
            </div>
          </div>
        ))}
      </div>
      <p className="bar-note">Latest full month only. This is not a sum of the months. A short month and a tiny month are left out.</p>
    </article>
  );
}

function OutlookSlide({ book }: { book: Book }) {
  const rows = (book.reading?.tracks ?? []).map((track) => rateOutlook(track)).filter((row) => row != null);
  return (
    <article className="slide paper compare">
      <header className="book-top">
        <p className="kicker">Outlook</p>
        <h2>One period ahead</h2>
      </header>
      <ul className={`compare-list compare-${Math.min(rows.length, 3)}`}>
        {rows.map((row) => (
          <li key={row.name}>
            <span>{row.name}</span>
            <strong>
              {row.rate}
              <i>%</i>
            </strong>
            <em>{row.label}</em>
            <p>
              Line from {row.from} to {row.to}.
            </p>
          </li>
        ))}
      </ul>
      <p className="bar-note">Projection is a straight line from the full months. It is not a result. A short month and a tiny month are left out of the line.</p>
    </article>
  );
}

function gridDrawn(grid: SheetGrid, index: number): SourceChart | null {
  const column = grid.columns[index];
  if (!column) return null;
  const fit = chartFit(grid.labels, [{ name: column.name, values: column.values, percent: column.percent }]);
  const points = columnPoints(grid, index).filter((point) => (fit.kind === "donut" ? !point.projected : true));
  return {
    kind: fit.kind,
    drawn: fit.sentence,
    categories: points.map((point) => point.label),
    series: [{ name: column.name, values: points.map((point) => point.value) }],
  };
}

function GridSlide({ book, index }: { book: Book; index: number }) {
  const grid = book.grid;
  const column = grid?.columns[index];
  if (!grid || !column) return null;
  const points = columnPoints(grid, index);
  const peak = Math.max(0, ...points.map((point) => point.value));
  const max = column.percent && peak <= 100 ? 100 : axisMax(peak);
  const latest = [...points].reverse().find((point) => !point.projected);
  const drawn = gridDrawn(grid, index);
  const kind = drawn?.kind ?? "columns";
  return (
    <article className="slide paper trend">
      <header className="trend-top">
        <div>
          <p className="kicker">{grid.cadence === "rows" ? "By row" : grid.cadence === "year" ? "By year" : "By month"}</p>
          <h2>{column.name}</h2>
        </div>
        {latest && (
          <div className="trend-rate">
            <p className="rate">
              {formatGridPoint(latest.value)}
              {column.percent ? <span>%</span> : null}
            </p>
            <p className="rate-sub">Latest row, {latest.label}. Not a total.</p>
          </div>
        )}
      </header>
      {kind === "line" ? (
        <div className="trend-chart is-line">
          <div className="trend-axis" aria-hidden>
            <span>{formatGridPoint(max)}</span>
            <span>{formatGridPoint(max / 2)}</span>
            <span>0</span>
          </div>
          <div className="trend-body">
            <PointLine points={points} max={max} label={gridChartNote(grid, index)} />
          </div>
        </div>
      ) : kind === "bars" && drawn ? (
        <div className="grid-fit">
          <RankChart chart={drawn} table={null} />
        </div>
      ) : kind === "donut" && drawn ? (
        <div className="grid-fit">
          <OriginDonut chart={drawn} />
        </div>
      ) : (
        <div className="trend-chart">
          <div className="trend-axis" aria-hidden>
            <span>{formatGridPoint(max)}</span>
            <span>{formatGridPoint(max / 2)}</span>
            <span>0</span>
          </div>
          <div className="trend-body">
            <div className="trend-plot" role="img" aria-label={gridChartNote(grid, index)}>
              {points.map((point) => {
                const height = point.value <= 0 || max <= 0 ? 0 : Math.min(1, point.value / max);
                return (
                  <div key={`${point.label}-${point.projected ? "proj" : "row"}`} className="trend-col">
                    <div className="trend-bar">
                      <i
                        className={point.projected ? "is-proj" : height === 0 ? "is-zero" : ""}
                        style={{ ["--h" as string]: height }}
                      />
                    </div>
                    <b>
                      <span>{point.label}</span>
                    </b>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
      <p className="bar-note">{gridChartNote(grid, index)}</p>
    </article>
  );
}

function ComboSlide({ book }: { book: Book }) {
  const grid = book.grid;
  if (!grid) return null;
  const pair = comboPair(
    grid.labels,
    grid.columns.map((column) => ({ name: column.name, values: column.values, percent: column.percent })),
  );
  if (!pair) return null;
  const count = grid.columns[pair.count];
  const rate = grid.columns[pair.rate];
  if (!count || !rate) return null;
  const rows = grid.labels
    .map((label, index) => ({ label, count: count.values[index], rate: rate.values[index] }))
    .filter((row): row is { label: string; count: number; rate: number } => row.count != null && row.rate != null);
  const max = axisMax(Math.max(0, ...rows.map((row) => row.count)));
  const width = 640;
  const height = 240;
  const pad = 24;
  const slot = rows.length > 0 ? (width - pad * 2) / rows.length : 0;
  const barW = Math.min(26, slot * 0.42);
  const yCount = (value: number) => height - 28 - (Math.max(0, value) / (max || 1)) * (height - 48);
  const yRate = (value: number) => height - 28 - (Math.max(0, Math.min(100, value)) / 100) * (height - 48);
  const xOf = (index: number) => pad + slot * index + slot / 2;
  const path = rows.map((row, index) => `${index === 0 ? "M" : "L"}${xOf(index).toFixed(1)} ${yRate(row.rate).toFixed(1)}`).join(" ");
  return (
    <article className="slide paper trend">
      <header className="trend-top">
        <div>
          <p className="kicker">Together</p>
          <h2>
            {count.name} and {rate.name}
          </h2>
        </div>
      </header>
      <svg className="combo-plot" viewBox={`0 0 ${width} ${height}`} role="img" aria-label={`${count.name} as columns and ${rate.name} as a line. Rows are not added.`}>
        {rows.map((row, index) => {
          const y = yCount(row.count);
          return <rect key={row.label} x={xOf(index) - barW / 2} y={y} width={barW} height={Math.max(0, height - 28 - y)} rx="4" />;
        })}
        <path d={path} />
        {rows.map((row, index) => (
          <circle key={`${row.label}-rate`} cx={xOf(index)} cy={yRate(row.rate)} r="4.5" />
        ))}
      </svg>
      <div className="trend-line-labels">
        {rows.map((row) => (
          <span key={`${row.label}-name`}>{row.label}</span>
        ))}
      </div>
      <p className="bar-note">
        {count.name} is the column. {rate.name} is the line, on its own scale of 100. Rows are not added together.
      </p>
    </article>
  );
}

function CardSlide({ book, index }: { book: Book; index: number }) {
  const card = book.outline?.cards[index];
  if (!card) return null;
  return (
    <article className="slide card-slide">
      <header>
        <p className="kicker">{card.kicker}</p>
        <h2>{card.title}</h2>
      </header>
      <div className="card-point">
        <p>{card.hint.trim() || "Add the point to mention."}</p>
      </div>
      <p className="bar-note">Placeholder. No figure was in the brief.</p>
    </article>
  );
}

function ReadSlide({ book }: { book: Book }) {
  const lines = book.facts
    ? factReadLines(book.facts)
    : book.story
      ? storyReadLines(book.story, book.period || book.story.period)
      : [];
  return (
    <article className="slide paper scope">
      <header className="book-top">
        <p className="kicker">What it says</p>
        <h2>{book.client || "The brief"}</h2>
      </header>
      <ul className="read-list">
        {lines.map((line) => (
          <li key={line} data-long={line.length > 140 ? "" : undefined}>
            {line}
          </li>
        ))}
      </ul>
    </article>
  );
}

function OriginCover({ book, slide }: { book: Book; slide: SourceSlide }) {
  const title = slide.title.trim() || book.sourceDeck?.title || "Untitled";
  const presenter = slide.context.trim();
  return (
    <article className="slide cover">
      <CoverWave />
      {book.logo ? (
        <div className="cover-mark">
          <img src={book.logo} alt="" />
        </div>
      ) : null}
      <div className="cover-lockup">
        <h1 data-lines={title.includes("\n") ? "" : undefined} data-long={title.includes("\n") ? undefined : title.length > 42 ? "wide" : title.length > 22 ? "" : undefined}><CoverTitle title={title} /></h1>
        {presenter ? <p className="presenter">{presenter}</p> : null}
      </div>
    </article>
  );
}

function OriginProse({ slide }: { slide: SourceSlide }) {
  const text = slide.context.trim();
  const shown = text.length > 1600 ? `${text.slice(0, 1600).trimEnd()}…` : text;
  const fit = shown.length > 700 ? "s" : shown.length > 380 ? "m" : "l";
  const picture = presentKind(slide) === "picture";
  const images = slide.images;
  const both = images.length > 0 && !picture && Boolean(shown);
  return (
    <article className={`slide paper brought${both ? " has-both" : ""}${picture ? " is-picture" : ""}`}>
      <div>
        <header>
          <p className="kicker">{picture ? "Kept" : `Slide ${slide.index}`}</p>
          {(slide.title.trim() || !picture) && <h2>{slide.title.trim() || "Context"}</h2>}
        </header>
        {picture && images.length === 0 ? (
          <p className="brought-body">This slide is a picture. The file did not store a title, a figure, or a chart as text.</p>
        ) : null}
        {!picture ? (
          <p className="brought-body" data-fit={fit}>
            {shown || "This slide had no readable text."}
          </p>
        ) : null}
      </div>
      <KeptImages images={images} label={slide.title.trim() || "Picture from the deck"} />
    </article>
  );
}

function PieceStack({ pieces, stretch, figureScale }: { pieces: SlidePiece[]; stretch?: "cards" | "lists"; figureScale?: "air" | "read" | "plate" }) {
  const bands = pieceBands(pieces);
  const fillAt = stretch === "lists"
    ? bands.findIndex((band) => band.kind === "lists")
    : stretch === "cards"
      ? fillBandIndex(bands)
      : -1;
  return (
    <>
      {bands.map((band, index) => {
        const fill = index === fillAt;
        if (band.kind === "prose") {
          const callout = index === bands.length - 1 && index > 0 && band.text.length <= 240;
          return (
            <div className={callout ? "piece-callout" : undefined} key={index}>
              <CopyBlocks text={band.text} tone="figure" />
            </div>
          );
        }
        if (band.kind === "lists") {
          const timed = band.items.filter((item) => isTimeTitle(item.title)).length >= 2;
          return (
            <div className={`piece-columns cols-${band.items.length}${timed ? " is-time" : ""}${fill ? " is-fill" : ""}`} key={index}>
              {band.items.map((item, itemIndex) => (
                <section key={`${item.title}-${itemIndex}`}>
                  {timed ? <p className="step-index">{String(itemIndex + 1).padStart(2, "0")}</p> : null}
                  {item.title ? <h3>{item.title}</h3> : null}
                  <ul className="bullet-list is-tight">
                    {item.points.map((point, pointIndex) => (
                      <li key={`${point}-${pointIndex}`}><i aria-hidden /><span>{point}</span></li>
                    ))}
                  </ul>
                </section>
              ))}
            </div>
          );
        }
        const across = gridAcross(band.tone, band.items.length);
        const displays = band.tone === "stat" ? band.items.map((item) => item.display) : [];
        const shared = sameFigure(displays);
        const scaled = figureScale === "air" || figureScale === "read";
        const ceiling = figureScale === "air" ? (band.items.length > 2 ? 96 : 128) : figureScale === "read" ? 72 : 0;
        return (
          <div className={`piece-grid is-${band.tone}${fill ? " is-fill" : ""}`} data-across={across} key={index}>
            {band.items.map((item, itemIndex) => {
              const share = shared && item.display && (scaled || !item.detail.trim()) ? figureShare(item.display, displays) : null;
              const size = ceiling && item.display ? figurePx(item.display, displays, ceiling) : 0;
              return (
                <article
                  key={`${item.title}-${itemIndex}`}
                  data-face={cardFace(item.title) || undefined}
                  data-share={share != null ? "" : undefined}
                  style={share != null ? { ["--h" as string]: String(share) } : undefined}
                >
                  {item.display ? <b aria-label={item.display} style={size ? { fontSize: size } : undefined}><Figure value={item.display} /></b> : null}
                  <strong>{item.display ? item.label || item.title : item.title}</strong>
                  {item.detail ? <span>{item.detail}</span> : null}
                </article>
              );
            })}
          </div>
        );
      })}
    </>
  );
}

function PosterLine({ text, proof }: { text: string; proof?: boolean }) {
  const size = proof ? proofPx(text) : posterPx(text);
  return (
    <div className={`piece-poster${proof ? "" : " is-fill"}`} style={{ "--poster": `${size}px` } as CSSProperties}>
      <CopyBlocks text={text} tone="figure" />
    </div>
  );
}

function OriginSheet({ slide }: { slide: SourceSlide }) {
  const bullets = bulletLines(slide);
  const lead = leadCopy(slide);
  const continued = isContinued(slide);
  const pictures = slide.images.length > 0;
  const spokenOnly = Boolean(slide.spoken?.trim()) && !lead && bullets.length === 0 && !pictures;
  return (
    <article className={`slide sheet${pictures ? " has-pictures" : ""}${bullets.length ? " has-list" : ""}${spokenOnly ? " is-spoken" : ""}`}>
      {continued ? <p className="kicker">Continued</p> : null}
      {slide.eyebrow ? <p className="eyebrow">{slide.eyebrow}</p> : null}
      {slide.title.trim() ? <h2 data-lines={slide.title.includes("\n") ? "" : undefined}>{slide.title.trim()}</h2> : null}
      {lead ? <p className="sheet-copy">{lead}</p> : null}
      {bullets.length > 0 ? (
        <ul className={`bullet-list${bullets.length > 8 ? " is-tight" : ""}`}>
          {bullets.map((line, index) => (
            <li key={`${line}-${index}`}>
              <i aria-hidden />
              <span>{line}</span>
            </li>
          ))}
        </ul>
      ) : null}
      {pictures ? (
        <div className="picture-stage">
          <KeptImages images={slide.images} label={slide.title.trim() || "Picture from the deck"} stage />
        </div>
      ) : null}
    </article>
  );
}

function OriginPieces({ slide }: { slide: SourceSlide }) {
  const continued = isContinued(slide);
  const pieces = slide.pieces ?? [];
  const composition = directedComposition(slide);
  const compare = composition === "versus" ? compareSplit(pieces) : null;
  const aside = composition === "hero" ? heroAside(pieces) : null;
  const proof = composition === "proof" ? proofSplit(pieces) : null;
  const poster = composition === "poster" && pieces[0]?.kind === "prose" ? pieces[0].text : "";
  const air = pieceAir(slide) && slide.images.length === 0;
  const stretch = !air ? undefined : composition === "path" ? "lists" : composition === "stack" || composition === "metrics" || composition === "cascade" ? "cards" : undefined;
  return (
    <article className={`slide piece-slide is-${composition}${air ? " is-air" : ""}`} data-fit={pieceFit(slide)} data-composition={composition}>
      {continued ? <p className="kicker">Continued</p> : null}
      {slide.eyebrow ? <p className="eyebrow">{slide.eyebrow}</p> : null}
      {slide.title.trim() ? <h2>{slide.title.trim()}</h2> : null}
      <div className={air || poster || proof || aside ? "piece-body is-stretch" : "piece-body"}>
        {poster ? <PosterLine text={poster} /> : null}
        {proof ? (
          <div className="piece-proof is-fill">
            <PosterLine text={proof.text} proof />
            <article data-face={cardFace(proof.card.title) || undefined}>
              {proof.card.display ? <b aria-label={proof.card.display}><Figure value={proof.card.display} /></b> : null}
              <strong>{proof.card.display ? proof.card.label || proof.card.title : proof.card.title}</strong>
              {proof.card.detail ? <span>{proof.card.detail}</span> : null}
            </article>
          </div>
        ) : null}
        {compare ? (
          <>
            {compare.before.length ? <PieceStack pieces={compare.before} /> : null}
            <div className="piece-split is-versus">
              <div><PieceStack pieces={compare.left} /></div>
              <div><PieceStack pieces={compare.right} /></div>
            </div>
          </>
        ) : null}
        {aside ? (
          <div className="piece-aside is-fill">
            <article>
              <b aria-label={aside.stat.display} style={{ fontSize: displayPx(aside.stat.display, air ? 120 : pieceFit(slide) === "roomy" ? 72 : 56) }}><Figure value={aside.stat.display} /></b>
              <strong>{aside.stat.label || aside.stat.title}</strong>
              {aside.stat.detail ? <span>{aside.stat.detail}</span> : null}
            </article>
            <div><PieceStack pieces={aside.rest} /></div>
          </div>
        ) : null}
        {!poster && !proof && !compare && !aside ? (
        <PieceStack
          pieces={pieces}
          stretch={stretch}
          figureScale={composition === "cascade" ? (air ? "air" : "read") : composition === "metrics" ? "plate" : undefined}
        />
      ) : null}
      </div>
      {slide.images.length ? (
        <div className="picture-stage">
          <KeptImages images={slide.images} label={slide.title.trim() || "Picture"} stage />
        </div>
      ) : null}
    </article>
  );
}

function OriginTable({ slide }: { slide: SourceSlide }) {
  const table = tableFromSlide(slide);
  const pieces = slide.pieces ?? [];
  const lead = pieces.length
    ? pieces.filter((piece) => piece.kind === "prose").map((piece) => piece.text).join("\n")
    : slide.context;
  const after = pieces.filter((piece) => piece.kind !== "prose");
  return (
    <article className="slide paper origin is-table">
      {isContinued(slide) ? <p className="kicker">Continued</p> : null}
      <header className="book-top">
        {slide.eyebrow ? <p className="eyebrow">{slide.eyebrow}</p> : <p className="kicker">Table</p>}
        <h2>{slide.title.trim() || "Table"}</h2>
      </header>
      {lead.trim() ? <CopyBlocks text={lead} tone="figure" /> : null}
      {table ? <TableCard table={table} /> : null}
      {after.length ? <PieceStack pieces={after} /> : null}
      {slide.images.length ? (
        <div className="picture-stage">
          <KeptImages images={slide.images} label={slide.title || "Picture"} stage />
        </div>
      ) : null}
    </article>
  );
}

function Origin({ book, index }: { book: Book; index: number }) {
  const slide = originSlide(book, index);
  if (!slide) {
    return (
      <article className="slide sheet">
        <p className="kicker">Slide</p>
        <h2>Missing slide</h2>
      </article>
    );
  }
  const role = slideRole(slide);
  const structured = (slide.pieces ?? []).some((piece) => piece.kind === "card" || piece.kind === "list");
  if (role === "cover") return <OriginCover book={book} slide={slide} />;
  if (slide.table && !slide.chart) return <OriginTable slide={slide} />;
  if (role !== "chart" && structured) return <OriginPieces slide={slide} />;
  if (role === "chart" && slideShow(slide) === "table") return <OriginTable slide={slide} />;
  if (role === "chart") return <OriginChart slide={slide} />;
  if (role === "figures") return <OriginNumbers slide={slide} />;
  if (role === "pictures" && presentKind(slide) === "picture" && slide.images.length === 0) return <OriginProse slide={slide} />;
  return <OriginSheet slide={slide} />;
}

export function SlideFrame({
  book,
  view,
  id,
  track = 0,
  options,
}: {
  book: Book;
  view: BookView;
  id: SlideId;
  track?: number;
  options: { mix: MixChart; listening: ListeningChart; type: TypeSize };
}) {
  const chapters = slidesFor(book)
    .filter((slide) => slide.id !== "cover")
    .map((slide) => slide.label);
  let node: ReactElement;
  if (id === "book") node = <BookSlide book={book} view={view} />;
  else if (id === "mix") node = <Mix book={book} view={view} chart={options.mix} />;
  else if (id === "versus") node = <Versus book={book} view={view} />;
  else if (id === "listening") node = <Listening book={book} chart={options.listening} />;
  else if (id === "journey") node = <Journey book={book} />;
  else if (id === "months") node = <Months book={book} />;
  else if (id === "source") node = <BroughtIn book={book} />;
  else if (id === "result") node = <Result book={book} />;
  else if (id === "facts") node = <FactSlide book={book} index={track} />;
  else if (id === "scope") node = <ScopeSlide book={book} />;
  else if (id === "track") node = <Trend book={book} index={track} />;
  else if (id === "compare") node = <Compare book={book} />;
  else if (id === "origin") node = <Origin book={book} index={track} />;
  else if (id === "progress") node = <ProgressSlide book={book} />;
  else if (id === "outlook") node = <OutlookSlide book={book} />;
  else if (id === "grid") node = <GridSlide book={book} index={track} />;
  else if (id === "combo") node = <ComboSlide book={book} />;
  else if (id === "card") node = <CardSlide book={book} index={track} />;
  else if (id === "read") node = <ReadSlide book={book} />;
  else if (id === "next") node = <Next book={book} view={view} />;
  else node = <Cover book={book} chapters={chapters} />;
  return <div data-size={options.type}>{node}</div>;
}
