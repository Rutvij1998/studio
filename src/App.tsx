import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { viewBook, type Book, type FactGroup, type Proof, type SourceDeck, type SourceImage, type SourceSlide, type Story } from "./book";
import { deckBlob } from "./deck";
import { DEFAULT_OPTIONS, type DeckOptions } from "./options";
import { emptyBook } from "./parse";
import { slidesFor, SlideFrame, type SlideId } from "./slides";
import { storyReadLines } from "./story";
import { factReadLines } from "./facts";
import { gridChartNote, gridProgress } from "./grid";
import { rateOutlook, seriesProgressList, trackNote } from "./reading";
import { markRequest } from "./mark";
import { addSlidePictures, canFlip, chartSentence, deleteSlide, duplicateSlide, formatPoint, mergeSlides, moveSlide, readSourceDeck, removeSlidePicture, replaceSlides, showSlide, slideRead, slideShow, sourceSummary, type SlideEdit } from "./sourceDeck";
import { gradeBrief, type GradeMark } from "./critic";
import { acceptPlan, applyPlan, deckFacts } from "./director";
import { GUIDE_BEATS, GUIDE_MS, HomeBeats, HomePreview } from "./homeGuide";

const W = 1280;
const H = 720;

type Held = { id: string; file: File };

function clip(text: string): string {
  return text.length > 4000 ? `${text.slice(0, 4000).trimEnd()}…` : text;
}

const originalCache = new Map<string, Promise<string[]>>();
let directTicket = 0;

function loadOriginal(file: File): Promise<string[]> {
  const key = `${file.name}:${file.size}:${file.lastModified}`;
  const hit = originalCache.get(key);
  if (hit) return hit;
  const pending = file.arrayBuffer().then(async (body) => {
    const response = await fetch("/api/original", {
      method: "POST",
      headers: { "content-type": "application/octet-stream" },
      body,
    });
    if (!response.ok) throw new Error("original");
    const payload = await response.json() as { slides?: unknown };
    const pages = Array.isArray(payload.slides)
      ? payload.slides.filter((item): item is string => typeof item === "string" && item.startsWith("data:image/"))
      : [];
    if (pages.length === 0) throw new Error("original");
    return pages;
  });
  originalCache.set(key, pending);
  pending.catch(() => originalCache.delete(key));
  return pending;
}

async function attachMark(book: Book): Promise<Book> {
  if (!book.client || !markRequest(book.client)) return { ...book, logo: "" };
  try {
    const response = await fetch("/api/mark", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ name: book.client }),
    });
    if (!response.ok) return book;
    const body = (await response.json()) as { dataUrl?: string };
    return body.dataUrl ? { ...book, logo: body.dataUrl } : book;
  } catch {
    return book;
  }
}

export function App() {
  const [mode, setMode] = useState<"home" | "studio">("home");
  const [files, setFiles] = useState<Held[]>([]);
  const [busy, setBusy] = useState(false);
  const [busyLabel, setBusyLabel] = useState("Restyling");
  const [book, setBook] = useState<Book>(emptyBook());
  const [options, setOptions] = useState<DeckOptions>(DEFAULT_OPTIONS);
  const [index, setIndex] = useState(0);
  const [note, setNote] = useState("Drop a PowerPoint. Every slide is kept.");
  const [understood, setUnderstood] = useState("");
  const [issue, setIssue] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [presenting, setPresenting] = useState(false);
  const [comparing, setComparing] = useState(false);
  const [originals, setOriginals] = useState<string[]>([]);
  const [originalsReady, setOriginalsReady] = useState(false);
  const [deckName, setDeckName] = useState<string | null>(null);
  const [reviewOpen, setReviewOpen] = useState(false);
  const [spoken, setSpoken] = useState<{ summary: string; notes: Record<number, string> }>({ summary: "", notes: {} });
  const comparingRef = useRef(false);
  const reviewRef = useRef(false);
  const originalTicket = useRef(0);
  const toggleCompareRef = useRef<() => void>(() => {});
  comparingRef.current = comparing;
  reviewRef.current = reviewOpen;
  toggleCompareRef.current = () => {
    if (!comparing && originals.length === 0) setOriginalsReady(false);
    setComparing((current) => !current);
  };
  const view = useMemo(() => viewBook(book), [book]);
  const slides = useMemo(() => slidesFor(book), [book]);
  const slide = slides[index] ?? slides[0];
  const review = useMemo(() => gradeBrief(book.sourceDeck?.slides ?? []), [book.sourceDeck]);
  const reviewMarks = useMemo(() => {
    const marks = new Map<number, GradeMark>();
    for (const mark of review.grade.slides) marks.set(mark.index, { ...mark, note: spoken.notes[mark.index] || mark.note });
    for (const [key, note] of Object.entries(spoken.notes)) {
      const at = Number(key);
      if (!note || marks.has(at) || !review.brief.some((row) => row.index === at)) continue;
      const title = review.brief.find((row) => row.index === at)?.title || "Slide";
      marks.set(at, { index: at, title, severity: "sparse", note });
    }
    return [...marks.values()].sort((left, right) => left.index - right.index);
  }, [review, spoken]);

  useEffect(() => {
    setIndex((current) => Math.min(current, Math.max(0, slides.length - 1)));
  }, [slides.length]);

  useEffect(() => {
    if (mode !== "studio") return;
    const onKey = (event: KeyboardEvent) => {
      const tag = (event.target as HTMLElement | null)?.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
      if (presenting && event.key === "Escape") {
        event.preventDefault();
        if (reviewRef.current) setReviewOpen(false);
        else if (comparingRef.current) setComparing(false);
        else setPresenting(false);
        return;
      }
      if (presenting && !event.repeat && (event.key === "c" || event.key === "C") && !event.metaKey && !event.ctrlKey && !event.altKey) {
        event.preventDefault();
        toggleCompareRef.current();
        return;
      }
      if (presenting && !event.repeat && (event.key === "g" || event.key === "G") && !event.metaKey && !event.ctrlKey && !event.altKey) {
        event.preventDefault();
        setReviewOpen((current) => !current);
        return;
      }
      const forward = event.key === "ArrowRight" || event.key === "ArrowDown" || event.key === "PageDown" || (presenting && event.key === " ");
      const back = event.key === "ArrowLeft" || event.key === "ArrowUp" || event.key === "PageUp";
      if (forward) {
        event.preventDefault();
        setIndex((current) => Math.min(slides.length - 1, current + 1));
      } else if (back) {
        event.preventDefault();
        setIndex((current) => Math.max(0, current - 1));
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [mode, slides.length, presenting]);

  useEffect(() => {
    const file = files[0]?.file;
    if (!comparing || originalsReady || !file) return;
    const ticket = originalTicket.current;
    let cancel = false;
    void loadOriginal(file).then((pages) => {
      if (cancel || originalTicket.current !== ticket) return;
      setOriginals(pages);
      setOriginalsReady(true);
    }).catch(() => {
      if (cancel || originalTicket.current !== ticket) return;
      setOriginals([]);
      setOriginalsReady(true);
    });
    return () => { cancel = true; };
  }, [comparing, originalsReady, files]);

  useEffect(() => {
    setSpoken({ summary: "", notes: {} });
    setReviewOpen(false);
    const brief = review.brief;
    if (!brief.length) return;
    let cancel = false;
    void fetch("/api/grade", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ slides: brief }),
    }).then(async (response) => {
      if (!response.ok || cancel) return;
      const body = await response.json() as { summary?: unknown; notes?: unknown };
      if (cancel) return;
      const notes: Record<number, string> = {};
      if (Array.isArray(body.notes)) {
        for (const item of body.notes) {
          if (!item || typeof item !== "object") continue;
          const row = item as { index?: unknown; note?: unknown };
          if (typeof row.index === "number" && typeof row.note === "string") notes[row.index] = row.note;
        }
      }
      setSpoken({ summary: typeof body.summary === "string" ? body.summary : "", notes });
    }).catch(() => undefined);
    return () => { cancel = true; };
  }, [review]);

  const patch = (partial: Partial<Book>) => setBook((current) => ({ ...current, ...partial }));

  const restyle = async (file: File) => {
    const lower = file.name.toLowerCase();
    if (lower.endsWith(".ppt") && !lower.endsWith(".pptx")) {
      setIssue(`${file.name} is an older PowerPoint file. Save it as .pptx.`);
      return;
    }
    if (!lower.endsWith(".pptx")) {
      setIssue("Add a PowerPoint. This only restyles a .pptx.");
      return;
    }
    setBusy(true);
    setBusyLabel("Reading the deck");
    setIssue(null);
    setComparing(false);
    setOriginals([]);
    setOriginalsReady(false);
    originalTicket.current += 1;
    const ticket = ++directTicket;
    try {
      const read = await readSourceDeck(await file.arrayBuffer(), file.name);
      if (directTicket !== ticket) return;
      if (read.deck.slides.length === 0) {
        setIssue(`${file.name} has no slides.`);
        return;
      }
      const original = read.deck.slides;
      let slides = original;
      let directed = false;
      setBusyLabel("Choosing the layout");
      try {
        const controller = new AbortController();
        const timer = window.setTimeout(() => controller.abort(), 35000);
        const response = await fetch("/api/direct", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify(deckFacts(original)),
          signal: controller.signal,
        });
        window.clearTimeout(timer);
        if (response.ok && directTicket === ticket) {
          const body = await response.json() as unknown;
          const facts = deckFacts(original);
          const incoming = body && typeof body === "object" ? body as { drop?: unknown; keep?: unknown; slides?: unknown } : {};
          const drop = [...facts.repeated, ...(Array.isArray(incoming.drop) ? incoming.drop : [])];
          const applied = applyPlan(original, acceptPlan(original, { ...incoming, drop }));
          slides = applied.slides;
          directed = applied.changed;
        }
      } catch {
        const applied = applyPlan(original, acceptPlan(original, { drop: deckFacts(original).repeated, keep: [], slides: [] }));
        slides = applied.slides;
        directed = applied.changed;
      }
      if (directTicket !== ticket) return;
      const title = read.deck.title;
      const next = await attachMark({
        ...emptyBook(),
        client: title,
        sourceName: file.name,
        sourceText: clip(read.text),
        sourceDeck: { ...read.deck, title, slides },
      });
      if (directTicket !== ticket) return;
      const pictures = slides.reduce((total, item) => total + item.images.length, 0);
      const directedNote = directed ? "Director kept every figure and adjusted the layout." : "";
      setBook(next);
      setDeckName(file.name);
      setFiles([{ id: `${file.name}:${file.size}`, file }]);
      setNote([read.note, pictures ? `${pictures} ${pictures === 1 ? "picture stays" : "pictures stay"} on the slides.` : "", directedNote].filter(Boolean).join(" "));
      setUnderstood(`${sourceSummary(next.sourceDeck ?? read.deck)}${next.logo ? ` The ${next.client} mark is on the cover.` : ""}`);
      setIndex(0);
      setMode("studio");
      setPresenting(true);
    } catch (error) {
      if (directTicket === ticket) setIssue(error instanceof Error ? error.message : "Could not read that deck.");
    } finally {
      if (directTicket === ticket) {
        setBusy(false);
        setBusyLabel("Restyling");
      }
    }
  };

  const addFiles = (list: File[]) => {
    const file = list.find((item) => item.name.toLowerCase().endsWith(".pptx")) ?? list[0];
    if (!file) return;
    void restyle(file);
  };

  const downloadDeck = async () => {
    setSaving(true);
    setIssue(null);
    try {
      const { blob, fileName } = await deckBlob(book, options);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = fileName;
      link.click();
      URL.revokeObjectURL(url);
      setNote(`Saved ${fileName}.`);
    } catch (error) {
      setIssue(error instanceof Error ? error.message : "Could not write the PowerPoint.");
    } finally {
      setSaving(false);
    }
  };

  if (mode === "home" || !slide) {
    return (
      <Home
        files={files}
        addFiles={addFiles}
        clearFile={() => setFiles([])}
        issue={issue}
        busy={busy}
        busyLabel={busyLabel}
      />
    );
  }

  return (
    <div className="app">
      <header className="top">
        <div className="brand">
          <span className="mark" aria-hidden />
          <strong>Studio</strong>
        </div>
        <div className="doc">
          <small>Restyle</small>
          <strong>{book.client || deckName || "Untitled deck"}</strong>
        </div>
        <div className="top-side">
          <button type="button" className="ghost" onClick={() => { setPresenting(false); setComparing(false); setOriginals([]); setOriginalsReady(false); originalTicket.current += 1; directTicket += 1; setBook(emptyBook()); setFiles([]); setDeckName(null); setUnderstood(""); setNote("Drop a PowerPoint. Every slide is kept."); setMode("home"); }}>
            New deck
          </button>
          <button type="button" className="ghost" onClick={() => setPresenting(true)} disabled={slides.length === 0}>
            Present
          </button>
          <button type="button" className="download" aria-label="Download PowerPoint" onClick={() => void downloadDeck()} disabled={saving}>
            {saving ? "Saving" : "Download"}
          </button>
          <p className="top-count">
            {index + 1} / {slides.length}
          </p>
        </div>
      </header>

      <div className="workspace">
        <nav className="rail" aria-label="Slides">
          {slides.map((item, itemIndex) => (
            <button
              key={`${item.id}-${item.track ?? item.label}`}
              type="button"
              className={itemIndex === index ? "is-on" : ""}
              aria-current={itemIndex === index ? "true" : undefined}
              onClick={() => setIndex(itemIndex)}
            >
              <span>{itemIndex + 1}</span>
              <span>
                <b>{item.label}</b>
                <small>{item.form}</small>
              </span>
            </button>
          ))}
        </nav>

        <Stage
          label={`Slide ${index + 1} of ${slides.length}, ${slide.label}. ${slide.form}.`}
          caption={slide.label}
          detail={slide.id === "mix" ? (options.mix === "columns" ? "Columns" : options.mix === "meter" ? "Meter" : "Donut") : slide.id === "listening" ? (options.listening === "columns" ? "Columns" : "Rows") : slide.form}
          index={index}
          count={slides.length}
          onPrev={() => setIndex((current) => Math.max(0, current - 1))}
          onNext={() => setIndex((current) => Math.min(slides.length - 1, current + 1))}
        >
          <SlideFrame key={`${slide.id}-${slide.track ?? 0}`} book={book} view={view} id={slide.id} track={slide.track ?? 0} options={options} />
        </Stage>

        <aside className="panel">
          <div className="panel-body">
            <p className="panel-kicker">{slide.form}</p>
            <h2>{slide.label}</h2>
            {slide.id === "origin" && book.sourceDeck ? (
              <SlideTools
                deck={book.sourceDeck}
                track={slide.track ?? 1}
                onChange={(next, focus, message) => {
                  patch({ sourceDeck: next });
                  setIndex(focus);
                  setNote(message);
                  setUnderstood(`${sourceSummary(next)}${book.logo ? ` The ${book.client} mark is on the cover.` : ""}`);
                }}
              />
            ) : null}
            <Look options={options} setOptions={setOptions} listening={false} mix={false} />
            {understood && <p className="chart-note">{understood}</p>}
            <Editor slideId={slide.id} trackIndex={slide.track ?? 0} book={book} rate={view.rate} peerRate={view.peerRate} patch={patch} />
            {view.warnings.length > 0 && (slide.id === "book" || slide.id === "mix") && (
              <ul className="warnings">
                {view.warnings.map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            )}
          </div>
          <div className="panel-sources">
            <Drop
              title="Replace the deck"
              detail="PowerPoint only"
              accept=".pptx,.ppt"
              fileName={deckName}
              onFile={(file) => void restyle(file)}
            />
            <p className="note" role="status">
              {note}
            </p>
            {issue && <p className="issue">{issue}</p>}
          </div>
        </aside>
      </div>
      {presenting && slide ? (
        <PresentDeck
          index={index}
          count={slides.length}
          title={book.client || deckName || "Untitled deck"}
          saving={saving}
          comparing={comparing}
          origin={slide.origin ?? 0}
          sameOrigin={(slide.origin ?? 0) > 0 && index > 0 && slides[index - 1]?.origin === slide.origin}
          originalUrl={(slide.origin ?? 0) > 0 ? originals[(slide.origin ?? 1) - 1] ?? "" : ""}
          originalReady={originalsReady}
          onClose={() => setPresenting(false)}
          onPick={(next) => setIndex(Math.max(0, Math.min(slides.length - 1, next)))}
          onDownload={() => void downloadDeck()}
          onCompare={() => toggleCompareRef.current()}
          grade={review.grade.grade}
          summary={spoken.summary || review.grade.summary}
          marks={reviewMarks}
          reviewOpen={reviewOpen}
          onReview={() => setReviewOpen((current) => !current)}
        >
          <SlideFrame key={`${slide.id}-${slide.track ?? 0}`} book={book} view={view} id={slide.id} track={slide.track ?? 0} options={options} />
        </PresentDeck>
      ) : null}
    </div>
  );
}

function PresentDeck({
  children,
  index,
  count,
  title,
  saving,
  comparing,
  origin,
  sameOrigin,
  originalUrl,
  originalReady,
  onClose,
  onPick,
  onDownload,
  onCompare,
  grade,
  summary,
  marks,
  reviewOpen,
  onReview,
}: {
  children: ReactNode;
  index: number;
  count: number;
  title: string;
  saving: boolean;
  comparing: boolean;
  origin: number;
  sameOrigin: boolean;
  originalUrl: string;
  originalReady: boolean;
  onClose: () => void;
  onPick: (index: number) => void;
  onDownload: () => void;
  onCompare: () => void;
  grade: string;
  summary: string;
  marks: GradeMark[];
  reviewOpen: boolean;
  onReview: () => void;
}) {
  const [scale, setScale] = useState(1);
  useEffect(() => {
    const measure = () => {
      if (!comparing) {
        setScale(Math.max(0.2, Math.min(window.innerWidth / W, (window.innerHeight - 76) / H)));
        return;
      }
      const width = (window.innerWidth - 64 - 36) / (W * 2);
      const height = (window.innerHeight - 76 - 52) / H;
      setScale(Math.max(0.05, Math.min(width, height)));
    };
    measure();
    window.addEventListener("resize", measure);
    return () => window.removeEventListener("resize", measure);
  }, [comparing]);
  const frame = (node: ReactNode) => (
    <div className="present-frame" style={{ width: W * scale, height: H * scale }}>
      <div className="fit-inner" style={{ transform: `scale(${scale})` }}>{node}</div>
    </div>
  );
  const yours = !origin ? "" : sameOrigin ? "Same slide in your file" : `Slide ${origin} in your file`;
  return (
    <div className={comparing ? "present is-comparing" : "present"} role="dialog" aria-modal="true" aria-label={comparing ? "Comparison" : "Presentation"} data-origin={origin || undefined}>
      {comparing ? (
        <div className="compare-stage">
          <div className="compare-col">
            <p className="compare-kicker">Your slide<small>{yours}</small></p>
            {frame(
              <article className="slide compare-source">
                {originalUrl ? (
                  <img src={originalUrl} alt="" />
                ) : (
                  <p>{originalReady ? "Your slide could not be opened." : "Opening your slide"}</p>
                )}
              </article>,
            )}
          </div>
          <div className="compare-col">
            <p className="compare-kicker is-studio">Studio<small>{`Slide ${index + 1} of ${count}`}</small></p>
            {frame(children)}
          </div>
        </div>
      ) : frame(children)}
      <footer className="present-bar">
        <p>{title}</p>
        <div className="present-dots" role="tablist" aria-label="Slides">
          {Array.from({ length: count }, (_, item) => (
            <button
              key={item}
              type="button"
              role="tab"
              aria-label={`Slide ${item + 1}`}
              aria-selected={item === index}
              onClick={() => onPick(item)}
            />
          ))}
        </div>
        {reviewOpen ? (
          <aside className="review-sheet" aria-label="Grade">
            <p className="review-grade">{grade}</p>
            <p>{summary}</p>
            {marks.length > 0 ? (
              <ul>
                {marks.map((mark) => (
                  <li key={mark.index}>
                    <button type="button" onClick={() => onPick(mark.index - 1)}>
                      <b>{`Slide ${mark.index}`}</b>
                      <strong>{mark.title}</strong>
                      <small>{mark.note}</small>
                    </button>
                  </li>
                ))}
              </ul>
            ) : (
              <p className="review-clear">Nothing to change.</p>
            )}
          </aside>
        ) : null}
        <div className="present-actions">
          <button type="button" className={grade === "A" || grade === "B" ? "is-grade" : "is-grade is-low"} aria-pressed={reviewOpen} aria-label={`Grade ${grade}`} onClick={onReview}>{grade}</button>
          <button type="button" aria-pressed={comparing} onClick={onCompare}>Compare</button>
          <button type="button" onClick={onClose}>Edit</button>
          <button type="button" className="is-download" onClick={onDownload} disabled={saving}>
            {saving ? "Saving" : "Download"}
          </button>
        </div>
      </footer>
    </div>
  );
}

function readPicture(file: File): Promise<SourceImage | null> {
  const mime = file.type === "image/png" ? "image/png" as const : file.type === "image/jpeg" ? "image/jpeg" as const : null;
  if (!mime) return Promise.resolve(null);
  return file.arrayBuffer().then((buffer) => {
    const bytes = new Uint8Array(buffer);
    let binary = "";
    for (let cursor = 0; cursor < bytes.length; cursor += 0x8000) binary += String.fromCharCode(...bytes.subarray(cursor, cursor + 0x8000));
    return { mime, dataUrl: `data:${mime};base64,${btoa(binary)}` };
  });
}

function SlideTools({
  deck,
  track,
  onChange,
}: {
  deck: SourceDeck;
  track: number;
  onChange: (deck: SourceDeck, index: number, note: string) => void;
}) {
  const slide = deck.slides.find((item) => item.index === track) ?? deck.slides[Math.max(0, track - 1)];
  if (!slide) return null;
  const at = deck.slides.findIndex((item) => item.index === slide.index);
  const apply = (edit: SlideEdit) => onChange(replaceSlides(deck, edit.slides), edit.index, edit.note);
  const title = at <= 0;
  const last = at >= deck.slides.length - 1;
  return (
    <div className="slide-tools">
      <div className="slide-actions" role="group" aria-label="Arrange">
        <button type="button" onClick={() => apply(moveSlide(deck.slides, slide.index, -1))} disabled={title || at <= 1}>Up</button>
        <button type="button" onClick={() => apply(moveSlide(deck.slides, slide.index, 1))} disabled={title || last}>Down</button>
        <button type="button" onClick={() => apply(duplicateSlide(deck.slides, slide.index))}>Duplicate</button>
        <button
          type="button"
          onClick={() => {
            if (!window.confirm("Delete this slide?")) return;
            apply(deleteSlide(deck.slides, slide.index));
          }}
          disabled={title}
        >
          Delete
        </button>
        <button type="button" onClick={() => apply(mergeSlides(deck.slides, slide.index))} disabled={title || last}>Merge with next</button>
      </div>
      {canFlip(slide) ? (
        <div className="show-switch" role="group" aria-label="Show the numbers">
          {(["chart", "table", "both"] as const).map((mode) => (
            <button key={mode} type="button" aria-pressed={slideShow(slide) === mode} onClick={() => apply(showSlide(deck.slides, slide.index, mode))}>
              {mode === "both" ? "Both" : mode === "chart" ? "Chart" : "Table"}
            </button>
          ))}
        </div>
      ) : null}
      <div className="picture-row">
        <label className="file-add">
          Add a picture
          <input
            type="file"
            accept="image/png,image/jpeg,.png,.jpg,.jpeg"
            onChange={(event) => {
              const file = event.target.files?.[0];
              event.target.value = "";
              if (!file) return;
              void readPicture(file).then((image) => {
                if (!image) {
                  onChange(deck, at, "Use a PNG or JPEG.");
                  return;
                }
                apply(addSlidePictures(deck.slides, slide.index, [image]));
              });
            }}
          />
        </label>
        {slide.images.length > 0 ? (
          <ul>
            {slide.images.map((image, imageIndex) => (
              <li key={`${image.mime}-${imageIndex}`}>
                <img src={image.dataUrl} alt="" />
                <button type="button" onClick={() => apply(removeSlidePicture(deck.slides, slide.index, imageIndex))}>Remove</button>
              </li>
            ))}
          </ul>
        ) : null}
      </div>
    </div>
  );
}

function panelNote(book: Book, slideId: SlideId, trackIndex: number): string {
  if (slideId === "progress") {
    if (book.grid) return gridProgress(book.grid)?.note ?? book.grid.note;
    const cards = book.reading ? seriesProgressList(book.reading) : [];
    if (cards.length === 0) return "Latest full month only.";
    return cards
      .map((card) => `${card.name} ${card.label}: ${card.contained.toLocaleString("en-US")} contained, ${card.rest.toLocaleString("en-US")} not contained.`)
      .join(" ");
  }
  if (slideId === "outlook" && book.reading) {
    const rows = book.reading.tracks.map((track) => rateOutlook(track)).filter((row) => row != null);
    if (rows.length === 0) return "Not enough full months for a projection.";
    return `${rows.map((row) => `${row.name} ${row.label} is ${row.rate}%.`).join(" ")} Projection is a straight line from the full months. It is not a result.`;
  }
  if (slideId === "grid" && book.grid) return gridChartNote(book.grid, trackIndex);
  if (slideId === "read") {
    const lines = book.facts ? factReadLines(book.facts) : book.story ? storyReadLines(book.story, book.period || book.story.period) : [];
    return lines.join(" ");
  }
  return "";
}

function Editor({
  slideId,
  trackIndex,
  book,
  rate,
  peerRate,
  patch,
}: {
  slideId: SlideId;
  trackIndex: number;
  book: Book;
  rate: number;
  peerRate: number;
  patch: (partial: Partial<Book>) => void;
}) {
  if (slideId === "origin" && book.sourceDeck) {
    const source = book.sourceDeck.slides.find((item) => item.index === trackIndex) ?? book.sourceDeck.slides[0];
    if (!source) return null;
    const write = (partial: Partial<SourceSlide>) => {
      const slides = book.sourceDeck!.slides.map((item) => (item.index === source.index ? { ...item, ...partial } : item));
      const title = source.index === book.sourceDeck!.slides[0]?.index && partial.title != null
        ? partial.title || book.sourceDeck!.title
        : book.client;
      patch({
        client: title,
        sourceDeck: { ...book.sourceDeck!, title: source.index === 1 ? (partial.title || book.sourceDeck!.title) : book.sourceDeck!.title, slides },
      });
    };
    const chart = source.chart;
    return (
      <div className="breakdown">
        <p className="chart-note">{slideRead(source)}</p>
        <div className="fields">
          <TextField label="Title" value={source.title} onChange={(title) => write({ title })} wide />
          <label className="field is-wide">
            <span>Context</span>
            <textarea aria-label="Context" rows={4} value={source.context} onChange={(event) => write({ context: event.target.value })} placeholder="No other words on this slide." />
          </label>
        </div>
        <div>
          <p className="chart-note">Figures</p>
          {source.numbers.length > 0 ? (
            <ul className="num-chips">
              {source.numbers.map((number, numberIndex) => (
                <li key={`${number.display}-${numberIndex}`}>
                  <b>{number.display}</b>
                  {number.label && <span>{number.label}</span>}
                </li>
              ))}
            </ul>
          ) : (
            <p className="chart-note">No figure was written on this slide.</p>
          )}
        </div>
        <div>
          <p className="chart-note">Chart</p>
          {chart ? (
            <>
              <p className="chart-note">{chartSentence(chart)}</p>
              <ul className="point-list">
                {chart.categories.slice(0, 8).map((category, index) => (
                  <li key={`${category}-${index}`}>
                    <span>{category}</span>
                    <b>{chart.series.map((series) => formatPoint(series.values[index] ?? 0)).join(" · ")}</b>
                  </li>
                ))}
              </ul>
              {chart.categories.length > 8 && <p className="chart-note">First 8 of {chart.categories.length} points.</p>}
            </>
          ) : (
            <p className="chart-note">{source.images.length ? "No chart data was stored. The picture stays on the slide." : "No chart was stored on this slide."}</p>
          )}
          {source.images.length > 0 && (
            <p className="chart-note">{source.images.length} {source.images.length === 1 ? "picture stays" : "pictures stay"} on this slide.</p>
          )}
        </div>
      </div>
    );
  }

  if (slideId === "card" && book.outline) {
    const card = book.outline.cards[trackIndex];
    if (!card) return null;
    const write = (partial: Partial<typeof card>) => {
      const cards = book.outline!.cards.map((item, itemIndex) => (itemIndex === trackIndex ? { ...item, ...partial } : item));
      patch({ outline: { ...book.outline!, cards } });
    };
    return (
      <div className="fields">
        <TextField label="Title" value={card.title} onChange={(title) => write({ title })} wide />
        <label className="field is-wide">
          <span>Point to mention</span>
          <textarea aria-label="Point to mention" rows={5} value={card.hint} onChange={(event) => write({ hint: event.target.value })} />
        </label>
      </div>
    );
  }

  if (slideId === "cover") {
    return (
      <div className="fields">
        <TextField label="Client" value={book.client} onChange={(client) => patch({ client })} />
        <TextField label="Market" value={book.market} onChange={(market) => patch({ market })} />
        <TextField label="Period" value={book.period} onChange={(period) => patch({ period })} wide />
        <label className="field is-wide">
          <span>Brief</span>
          <textarea aria-label="Brief" rows={5} value={book.brief} onChange={(event) => patch({ brief: event.target.value })} />
        </label>
      </div>
    );
  }

  if (slideId === "book" || slideId === "mix") {
    return (
      <>
        <p className="readout-line">
          <b>{rate}%</b> of active subscribers
        </p>
        <div className="fields">
          <NumberField label="Active subscribers" value={book.active} onChange={(active) => patch({ active })} />
          <NumberField label="Claims" value={book.claims} onChange={(claims) => patch({ claims })} />
          <NumberField label="Fulfilled" value={book.fulfilled} onChange={(fulfilled) => patch({ fulfilled })} />
          <NumberField label="Approved" value={book.approved} onChange={(approved) => patch({ approved })} />
        </div>
      </>
    );
  }

  if (slideId === "listening") {
    return (
      <p className="chart-note">
        {book.listening.length} topics from the Listening sheet. Row length follows the largest count. Colour runs negative, neutral, positive. Health is the share that is not negative.
      </p>
    );
  }

  if (slideId === "journey") {
    return (
      <p className="chart-note">
        {book.journey.length} stages from the Journey sheet. The percents are the sheet's own split, drawn out from the center line.
      </p>
    );
  }

  if (slideId === "months") {
    return (
      <p className="chart-note">
        {book.months.length} months from the Months sheet. Each column is that month's claims. A book sheet on its own leaves this slide off.
      </p>
    );
  }

  if (slideId === "progress" || slideId === "outlook" || slideId === "grid" || slideId === "read") {
    return <p className="chart-note">{panelNote(book, slideId, trackIndex)}</p>;
  }

  if ((slideId === "scope" || slideId === "track" || slideId === "compare") && book.reading) {
    const track = book.reading.tracks[trackIndex];
    return (
      <p className="chart-note">
        {slideId === "track" && track
          ? trackNote(track)
          : slideId === "scope"
            ? `${book.reading.scope.length} populations from the brief: ${book.reading.scope.map((card) => card.label).join(", ")}.`
            : "Each rate is that series alone. Calls and chats stay in their own totals."}
      </p>
    );
  }

  if (slideId === "facts" && book.facts) {
    const facts = book.facts;
    const group = facts.groups[trackIndex];
    if (!group) return null;
    const writeGroup = (partial: Partial<FactGroup>) => {
      const groups = facts.groups.map((item, itemIndex) => (itemIndex === trackIndex ? { ...item, ...partial } : item));
      patch({ facts: { ...facts, groups } });
    };
    return (
      <div className="fields">
        {group.volume && (
          <>
            <NumberField
              label="Volume"
              value={group.volume.value}
              onChange={(value) => writeGroup({ volume: { ...group.volume!, value } })}
            />
            <TextField label="What was counted" value={group.volume.unit} onChange={(unit) => writeGroup({ volume: { ...group.volume!, unit, label: unit } })} />
          </>
        )}
        {group.rates.map((rate, index) => (
          <ProofFields
            key={`${index}-${rate.label}`}
            proof={{ label: rate.label, relation: rate.relation, value: rate.value }}
            onChange={(partial) => {
              const rates = group.rates.map((item, itemIndex) => (itemIndex === index ? { ...item, ...partial } : item));
              writeGroup({ rates });
            }}
          />
        ))}
      </div>
    );
  }

  if (slideId === "result" && book.story) {
    const story = book.story;
    const setStory = (partial: Partial<Story>) => patch({ story: { ...story, ...partial } });
    const setProof = (index: number, partial: Partial<Proof>) => {
      const proofs = story.proofs.map((proof, proofIndex) => (proofIndex === index ? { ...proof, ...partial } : proof));
      setStory({ proofs });
    };
    return (
      <>
        <div className="fields">
          <NumberField
            label="Volume"
            value={story.volume ?? 0}
            onChange={(volume) => setStory({ volume })}
          />
          <TextField label="What was counted" value={story.unit} onChange={(unit) => setStory({ unit })} />
          {story.hedge ? <TextField label="As written" value={story.hedge} onChange={(hedge) => setStory({ hedge })} /> : null}
          <TextField label="Verb" value={story.verb} onChange={(verb) => setStory({ verb })} />
          {story.proofs.map((proof, index) => (
            <ProofFields key={`${index}-${proof.label}`} proof={proof} onChange={(partial) => setProof(index, partial)} />
          ))}
        </div>
      </>
    );
  }

  if (slideId === "source") {
    return (
      <div className="fields">
        <TextField label="Document" value={book.sourceName} onChange={(sourceName) => patch({ sourceName })} wide />
        <label className="field is-wide">
          <span>Decoded text</span>
          <textarea
            aria-label="Decoded text"
            rows={10}
            value={book.sourceText}
            onChange={(event) => patch({ sourceText: event.target.value })}
          />
        </label>
      </div>
    );
  }

  if (slideId === "versus") {
    return (
      <>
        <div className="readout">
          <p>
            <b>{rate}%</b>
            <span>{book.client || "This book"}</span>
          </p>
          <p>
            <b>{peerRate}%</b>
            <span>{book.peerName || "Peer"}</span>
          </p>
        </div>
        <div className="fields">
          <TextField label="Peer" value={book.peerName} onChange={(peerName) => patch({ peerName })} wide />
          <NumberField label="Peer active" value={book.peerActive} onChange={(peerActive) => patch({ peerActive })} />
          <NumberField label="Peer claims" value={book.peerClaims} onChange={(peerClaims) => patch({ peerClaims })} />
          <NumberField label="Peer fulfilled" value={book.peerFulfilled} onChange={(peerFulfilled) => patch({ peerFulfilled })} wide />
        </div>
      </>
    );
  }

  return (
    <label className="field is-wide">
      <span>Recommendation</span>
      <textarea
        aria-label="Recommendation"
        rows={8}
        value={book.recommendation}
        onChange={(event) => patch({ recommendation: event.target.value })}
      />
    </label>
  );
}

function Stage({
  children,
  label,
  caption,
  detail,
  index,
  count,
  onPrev,
  onNext,
}: {
  children: ReactNode;
  label: string;
  caption: string;
  detail: string;
  index: number;
  count: number;
  onPrev: () => void;
  onNext: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [scale, setScale] = useState(0.45);

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const measure = () => {
      const style = getComputedStyle(el);
      const padX = parseFloat(style.paddingLeft) + parseFloat(style.paddingRight);
      const padY = parseFloat(style.paddingTop) + parseFloat(style.paddingBottom);
      const width = el.clientWidth - padX;
      const height = el.clientHeight - padY - 76;
      const fromWidth = width / W;
      const fromHeight = height / H;
      const narrow = window.matchMedia("(max-width: 1080px)").matches;
      const next = narrow ? fromWidth : Math.min(fromWidth, fromHeight);
      setScale(next > 0.2 ? next : 0.2);
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  return (
    <div className="stage-wrap" ref={ref}>
      <div className="stage-stack">
        <div className="fit" style={{ width: W * scale, height: H * scale }} role="img" aria-label={label}>
          <div className="fit-inner" style={{ transform: `scale(${scale})` }}>
            {children}
          </div>
        </div>
        <div className="scrub">
          <button type="button" aria-label="Previous slide" onClick={onPrev} disabled={index === 0}>
            <Chevron />
          </button>
          <p>
            {caption}
            <span>{detail}</span>
          </p>
          <button type="button" aria-label="Next slide" onClick={onNext} disabled={index >= count - 1}>
            <Chevron flip />
          </button>
        </div>
      </div>
    </div>
  );
}

function Chevron({ flip = false }: { flip?: boolean }) {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden style={flip ? { transform: "scaleX(-1)" } : undefined}>
      <path d="M10 3.5 5.5 8 10 12.5" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function ProofFields({ proof, onChange }: { proof: Proof; onChange: (partial: Partial<Proof>) => void }) {
  return (
    <>
      <TextField label="Mark" value={proof.label} onChange={(label) => onChange({ label })} />
      <label className="field">
        <span>How it was stated</span>
        <select
          aria-label={`${proof.label || "Mark"} comparison`}
          value={proof.relation}
          onChange={(event) => onChange({ relation: event.target.value as Proof["relation"] })}
        >
          <option value="above">Above</option>
          <option value="below">Below</option>
          <option value="least">At least</option>
          <option value="exact">Exactly</option>
        </select>
      </label>
      <NumberField label="Figure" value={proof.value} onChange={(value) => onChange({ value })} wide />
    </>
  );
}

function TextField({
  label,
  value,
  onChange,
  wide = false,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  wide?: boolean;
}) {
  return (
    <label className={`field${wide ? " is-wide" : ""}`}>
      <span>{label}</span>
      <input aria-label={label} value={value} onChange={(event) => onChange(event.target.value)} />
    </label>
  );
}

function NumberField({
  label,
  value,
  onChange,
  wide = false,
}: {
  label: string;
  value: number;
  onChange: (value: number) => void;
  wide?: boolean;
}) {
  const [focus, setFocus] = useState(false);
  const [text, setText] = useState(String(value));
  const shown = focus ? text : value.toLocaleString("en-US");
  return (
    <label className={`field${wide ? " is-wide" : ""}`}>
      <span>{label}</span>
      <input
        aria-label={label}
        inputMode="numeric"
        value={shown}
        onFocus={(event) => {
          setFocus(true);
          setText(String(value));
          event.currentTarget.select();
        }}
        onBlur={() => setFocus(false)}
        onChange={(event) => {
          const raw = event.target.value.replace(/[^\d]/g, "");
          setText(raw);
          onChange(raw === "" ? 0 : Number(raw));
        }}
      />
    </label>
  );
}

function Drop({
  title,
  detail,
  accept,
  fileName,
  onFile,
}: {
  title: string;
  detail: string;
  accept: string;
  fileName: string | null;
  onFile: (file: File) => void;
}) {
  const [over, setOver] = useState(false);
  const take = (file: File | undefined) => {
    if (file) onFile(file);
  };
  return (
    <label
      className={`drop${over ? " is-over" : ""}`}
      onDragEnter={(event) => {
        event.preventDefault();
        setOver(true);
      }}
      onDragOver={(event) => {
        event.preventDefault();
        setOver(true);
      }}
      onDragLeave={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setOver(false);
      }}
      onDrop={(event) => {
        event.preventDefault();
        setOver(false);
        take(event.dataTransfer.files[0]);
      }}
    >
      <input
        type="file"
        accept={accept}
        onChange={(event) => {
          take(event.target.files?.[0]);
          event.target.value = "";
        }}
      />
      <strong>{title}</strong>
      <span>{fileName ?? detail}</span>
    </label>
  );
}

function Look({
  options,
  setOptions,
  listening,
  mix = true,
}: {
  options: DeckOptions;
  setOptions: (update: (current: DeckOptions) => DeckOptions) => void;
  listening: boolean;
  mix?: boolean;
}) {
  return (
    <div className="look">
      <Choice
        label="Type size"
        value={options.type}
        options={[
          { value: "s", label: "Compact" },
          { value: "m", label: "Regular" },
          { value: "l", label: "Large" },
        ]}
        onChange={(type) => setOptions((current) => ({ ...current, type }))}
      />
      {mix && (
        <Choice
          label="Mix chart"
          value={options.mix}
          options={[
            { value: "donut", label: "Donut" },
            { value: "columns", label: "Columns" },
            { value: "meter", label: "Meter" },
          ]}
          onChange={(mixChart) => setOptions((current) => ({ ...current, mix: mixChart }))}
        />
      )}
      {listening && (
        <Choice
          label="Listening"
          value={options.listening}
          options={[
            { value: "rows", label: "Rows" },
            { value: "columns", label: "Columns" },
          ]}
          onChange={(listeningChart) => setOptions((current) => ({ ...current, listening: listeningChart }))}
        />
      )}
    </div>
  );
}

function Choice<T extends string>({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: T;
  options: { value: T; label: string }[];
  onChange: (value: T) => void;
}) {
  return (
    <div className="choice" role="group" aria-label={label}>
      <span>{label}</span>
      <div>
        {options.map((option) => (
          <button
            key={option.value}
            type="button"
            className={value === option.value ? "is-on" : ""}
            aria-pressed={value === option.value}
            onClick={() => onChange(option.value)}
          >
            {option.label}
          </button>
        ))}
      </div>
    </div>
  );
}

function Home({
  files,
  addFiles,
  clearFile,
  issue,
  busy,
  busyLabel,
}: {
  files: Held[];
  addFiles: (files: File[]) => void;
  clearFile: () => void;
  issue: string | null;
  busy: boolean;
  busyLabel: string;
}) {
  const [over, setOver] = useState(false);
  const [beat, setBeat] = useState(0);
  const [paused, setPaused] = useState(false);
  const reduced = useRef(false);
  useEffect(() => {
    reduced.current = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  }, []);
  useEffect(() => {
    if (paused || reduced.current) return;
    const timer = window.setTimeout(() => setBeat((current) => (current + 1) % GUIDE_BEATS.length), GUIDE_MS);
    return () => window.clearTimeout(timer);
  }, [beat, paused]);
  return (
    <div className="home">
      <header className="home-bar">
        <div className="brand">
          <span className="mark" aria-hidden />
          <strong>Studio</strong>
        </div>
      </header>
      <main className="home-stage" aria-labelledby="home-title">
        <div className="home-hero">
          <p className="home-kicker">Restyle</p>
          <h1 id="home-title">
            Drop the deck.
            <span>
              Leave with a better one<span className="home-dot">.</span>
            </span>
          </h1>
          <form
            className="home-form"
            onSubmit={(event) => {
              event.preventDefault();
              const file = files[0]?.file;
              if (file) void addFiles([file]);
            }}
          >
            <label
              className={`lane${over ? " is-over" : ""}`}
              onDragEnter={(event) => {
                event.preventDefault();
                setOver(true);
              }}
              onDragOver={(event) => {
                event.preventDefault();
                setOver(true);
              }}
              onDragLeave={(event) => {
                if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setOver(false);
              }}
              onDrop={(event) => {
                event.preventDefault();
                setOver(false);
                addFiles([...event.dataTransfer.files]);
              }}
            >
              <input
                type="file"
                accept=".pptx,.ppt"
                disabled={busy}
                onChange={(event) => {
                  addFiles([...(event.target.files ?? [])]);
                  event.target.value = "";
                }}
              />
              <span className="lane-arrow" aria-hidden="true">
                <svg viewBox="0 0 24 24" width="26" height="26">
                  <path d="M12 16V4" />
                  <path d="M7 9l5-5 5 5" />
                  <path d="M4 20h16" />
                </svg>
              </span>
              <strong>{busy ? busyLabel : "Drop a PowerPoint here"}</strong>
              <span>{busy ? (busyLabel === "Choosing the layout" ? "Deciding what stays and what changes." : "Every slide is kept.") : "Click to upload. Every slide is kept."}</span>
            </label>
            {files.length > 0 && (
              <ul className="held">
                {files.map((item) => (
                  <li key={item.id}>
                    <span className="pill deck">PowerPoint</span>
                    <b>{item.file.name}</b>
                    <button type="button" aria-label={`Remove ${item.file.name}`} onClick={clearFile}>
                      ×
                    </button>
                  </li>
                ))}
              </ul>
            )}
            {issue && <p className="home-issue">{issue}</p>}
            {busy ? (
              <button className="home-start" type="submit" disabled>
                {busyLabel}
              </button>
            ) : null}
          </form>
        </div>
        <HomePreview index={beat} onPause={setPaused} />
        <HomeBeats index={beat} onPick={setBeat} onPause={setPaused} />
      </main>
      {busy ? (
        <div className="home-loading" role="status" aria-live="polite">
          <div>
            <p>Loading</p>
            <span>{busyLabel === "Choosing the layout" ? "Choosing the layout." : "Reading the deck."}</span>
          </div>
        </div>
      ) : null}
    </div>
  );
}
