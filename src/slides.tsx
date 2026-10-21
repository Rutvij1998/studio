import { type CSSProperties, type ReactElement } from "react";
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
import { axisMax, bulletLines, chartKicker, chartSentence, formLabel, formatPoint, isContinued, isSource, leadCopy, presentKind, railLabel, rowDisplay, slideRole } from "./sourceDeck";
import type { SheetGrid, SourceChart, SourceSlide, SourceTable } from "./book";

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

export type SlideMeta = { id: SlideId; label: string; form: string; track?: number };

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

function OriginColumns({ chart, stacked }: { chart: SourceChart; stacked: boolean }) {
  const cats = chart.categories.slice(0, 12);
  const series = chart.series.slice(0, 4).map((item) => ({ ...item, values: item.values.slice(0, 12) }));
  const peaks = cats.map((_, index) => stacked
    ? series.reduce((total, item) => total + Math.max(0, item.values[index] ?? 0), 0)
    : Math.max(0, ...series.map((item) => item.values[index] ?? 0)));
  const max = axisMax(Math.max(0, ...peaks));
  const single = series.length === 1;
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
                    <i key={`${item.name}-${seriesIndex}`} className={`s${seriesIndex}`} style={{ ["--h" as string]: Math.max(0, value) / max }}>
                      {single && printed.length <= 6 ? <em>{printed}</em> : null}
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

function OriginBars({ chart, table }: { chart: SourceChart; table?: SourceTable | null }) {
  const cats = chart.categories.slice(0, 14);
  const series = chart.series.slice(0, 3).map((item) => ({ ...item, values: item.values.slice(0, 12) }));
  const max = axisMax(Math.max(0, ...series.flatMap((item) => item.values.map((value) => Math.max(0, value)))));
  return (
    <ul className="origin-rows" role="img" aria-label={chartSentence(chart)}>
      {cats.map((cat, index) => (
        <li key={`${cat}-${index}`}>
          <span>{cat}</span>
          <div>
            {series.map((item, seriesIndex) => (
              <i key={`${item.name}-${seriesIndex}`} className={`s${seriesIndex}`} style={{ width: `${(Math.max(0, item.values[index] ?? 0) / max) * 100}%` }} />
            ))}
          </div>
          <b>{rowDisplay(table ?? null, cat, series, series.map((item) => formatPoint(item.values[index] ?? 0)).join(" · "))}</b>
        </li>
      ))}
    </ul>
  );
}

function OriginLine({ chart }: { chart: SourceChart }) {
  const cats = chart.categories.slice(0, 16);
  const series = chart.series.slice(0, 4).map((item) => ({ ...item, values: item.values.slice(0, 16) }));
  const max = axisMax(Math.max(0, ...series.flatMap((item) => item.values.map((value) => Math.max(0, value)))));
  const width = 640;
  const height = 220;
  const pad = 36;
  const step = cats.length > 1 ? (width - pad * 2) / (cats.length - 1) : 0;
  const xOf = (point: number) => (cats.length === 1 ? width / 2 : pad + point * step);
  const yOf = (value: number) => height - (Math.max(0, value) / max) * (height - 28) - 8;
  return (
    <figure className="origin-line">
      <div className="origin-axis" aria-hidden>
        <span>{formatPoint(max)}</span>
        <span>{formatPoint(max / 2)}</span>
        <span>0</span>
      </div>
      <div className="origin-line-body">
        <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label={chartSentence(chart)}>
          <line x1={pad} y1={yOf(max / 2)} x2={width - pad} y2={yOf(max / 2)} />
          {series.map((item, index) => {
            const path = item.values.map((value, point) => {
              return `${point === 0 ? "M" : "L"}${xOf(point).toFixed(1)} ${yOf(value).toFixed(1)}`;
            }).join(" ");
            return <path key={`${item.name}-${index}`} className={`s${index}`} d={path} />;
          })}
          {series.length === 1 && cats.length <= 8 && series[0]!.values.map((value, point) => {
            const x = xOf(point);
            const y = yOf(value);
            return (
              <g key={`point-${point}`}>
                <circle className="s0" cx={x} cy={y} r="4.5" />
                <text x={x} y={y - 12} textAnchor="middle">{formatPoint(value)}</text>
              </g>
            );
          })}
        </svg>
        <div className="origin-labels">
          {cats.map((cat, index) => (
            <span key={`${cat}-${index}`}>{cat}</span>
          ))}
        </div>
      </div>
    </figure>
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

function KeptImages({ images, label }: { images: SourceSlide["images"]; label: string }) {
  if (images.length === 0) return null;
  const shown = images.slice(0, 4);
  return (
    <div className={`kept-images kept-${shown.length}`}>
      {shown.map((image, index) => (
        <img key={`${image.mime}-${index}`} src={image.dataUrl} alt={label} />
      ))}
    </div>
  );
}

function TableCard({ table }: { table: SourceTable }) {
  const dense = table.rows.length > 6;
  return (
    <div className={`source-table${dense ? " is-dense" : ""}`}>
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
  const kind = chart.kind;
  if (kind === "funnel") return <OriginFunnel chart={chart} table={table} />;
  if (kind === "donut") return <OriginDonut chart={chart} />;
  if (kind === "bars") return <OriginBars chart={chart} table={table} />;
  if (kind === "line") return <OriginLine chart={chart} />;
  return <OriginColumns chart={chart} stacked={kind === "stacked"} />;
}

function OriginChart({ slide }: { slide: SourceSlide }) {
  const chart = slide.chart;
  if (!chart) return null;
  const kind = chart.kind;
  const limit = kind === "bars" ? 14 : kind === "line" || kind === "funnel" ? 16 : 12;
  const hidden = chart.categories.length > limit;
  const note = [
    chart.series.length > 1 && kind === "donut" ? "The first series is drawn." : "",
    hidden ? `First ${limit} of ${chart.categories.length} points.` : "",
    slide.table && slide.table.rows.length > chart.categories.length ? `Every row stays in the table.` : "",
    slide.extraCharts > 0 ? "Another chart is on this slide and is not drawn." : "",
    chart.drawn,
  ].filter(Boolean).join(" ");
  return (
    <article className={`slide story${slide.table ? " has-table" : ""}`}>
      <header className="origin-top">
        <div>
          <p className="kicker">{chartKicker(chart)}</p>
          <h2>{slide.title || "Chart"}</h2>
        </div>
        <OriginLegend chart={chart} />
      </header>
      {slide.context.trim() && <p className="story-line">{slide.context.trim()}</p>}
      <div className="story-stage">
        <div className={`story-plot${slide.images.length ? " has-kept" : ""}${chart.categories.length > 6 ? " is-tight" : ""}`}>
          <ChartPlot chart={chart} table={slide.table} />
          <KeptImages images={slide.images} label={slide.title || "Chart picture"} />
        </div>
        {slide.table && <TableCard table={slide.table} />}
      </div>
      {note && <p className="bar-note">{note}</p>}
    </article>
  );
}

function OriginNumbers({ slide }: { slide: SourceSlide }) {
  const numbers = slide.numbers.slice(0, 6);
  const context = slide.context.trim();
  const shown = context.length > 280 ? `${context.slice(0, 280).trimEnd()}…` : context;
  if (numbers.length === 1) {
    const number = numbers[0]!;
    const wide = number.display.length > 8;
    return (
      <article className="slide paper origin">
        <header className="book-top">
          <p className="kicker">Figure</p>
          <h2>{slide.title || number.label || "Figure"}</h2>
        </header>
        <div className="rate-block">
          <p className="rate" data-wide={wide ? "" : undefined}>{number.display}</p>
          {number.label && <p className="rate-caption">{number.label}</p>}
        </div>
        {shown && <p className="origin-context">{shown}</p>}
        <KeptImages images={slide.images} label={slide.title || "Picture"} />
      </article>
    );
  }
  const columns = numbers.length === 2 || numbers.length === 4 ? numbers.length : numbers.length === 3 ? 3 : 3;
  return (
    <article className="slide paper origin">
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
      {shown && <p className="bar-note">{shown}</p>}
      <KeptImages images={slide.images} label={slide.title || "Picture"} />
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
          <OriginBars chart={drawn} />
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
      {book.logo ? (
        <div className="cover-mark">
          <img src={book.logo} alt="" />
        </div>
      ) : null}
      <div className="cover-lockup">
        <h1 data-long={title.length > 42 ? "wide" : title.length > 22 ? "" : undefined}>{title}</h1>
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

function OriginSheet({ slide }: { slide: SourceSlide }) {
  const bullets = bulletLines(slide);
  const lead = leadCopy(slide);
  const continued = isContinued(slide);
  const pictures = slide.images.length > 0;
  return (
    <article className={`slide sheet${pictures ? " has-pictures" : ""}${bullets.length ? " has-list" : ""}`}>
      {continued ? <p className="kicker">Continued</p> : null}
      {slide.title.trim() ? <h2>{slide.title.trim()}</h2> : null}
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
        <div className="sheet-stage">
          <KeptImages images={slide.images} label={slide.title.trim() || "Picture from the deck"} />
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
  if (role === "cover") return <OriginCover book={book} slide={slide} />;
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
