import type { CSSProperties } from "react";

export const GUIDE_MS = 3600;

export const GUIDE_BEATS = [
  {
    id: "drop",
    kicker: "01",
    title: "Drop the deck",
    say: "Start with the file you already have.",
    detail: "One PowerPoint. Every slide is kept.",
    look: "down",
  },
  {
    id: "title",
    kicker: "02",
    title: "Open on the title",
    say: "The first slide stays a title. Nothing else.",
    detail: "Dark, with the subject and the presenter.",
    look: "left",
  },
  {
    id: "chart",
    kicker: "03",
    title: "Draw the numbers",
    say: "Same cells. A clearer picture.",
    detail: "A table becomes a chart. The written cells stay.",
    look: "right",
  },
  {
    id: "picture",
    kicker: "04",
    title: "Keep the pictures",
    say: "Pictures stay pictures.",
    detail: "A chart saved as an image stays an image.",
    look: "left",
  },
  {
    id: "leave",
    kicker: "05",
    title: "Leave with the file",
    say: "Then take the deck with you.",
    detail: "Present it, rearrange it, then download.",
    look: "center",
  },
] as const;

export type GuideBeat = (typeof GUIDE_BEATS)[number];

export function GuideBot({ look }: { look: GuideBeat["look"] }) {
  return (
    <div className={`guide-bot is-${look}`} aria-hidden="true">
      <span className="guide-antenna" />
      <span className="guide-face">
        <span className="guide-eye"><i /></span>
        <span className="guide-eye"><i /></span>
        <span className="guide-mouth" />
      </span>
    </div>
  );
}

function GuideScene({ beat }: { beat: GuideBeat }) {
  if (beat.id === "drop") {
    return (
      <div className="scene-drop">
        <div className="scene-file">
          <b>Deck.pptx</b>
          <span>PowerPoint</span>
        </div>
      </div>
    );
  }
  if (beat.id === "chart") {
    return (
      <div className="scene-chart">
        <div className="scene-rows" aria-hidden="true">
          <i /><i /><i />
        </div>
        <div className="scene-bars" aria-hidden="true">
          <i style={{ "--h": "72%" } as CSSProperties} />
          <i style={{ "--h": "44%" } as CSSProperties} />
          <i style={{ "--h": "86%" } as CSSProperties} />
        </div>
      </div>
    );
  }
  if (beat.id === "picture") {
    return (
      <div className="scene-picture">
        <div className="scene-shot">
          <i /><i /><b />
        </div>
      </div>
    );
  }
  return (
    <div className="scene-leave">
      <span>Present</span>
      <span>Rearrange</span>
      <span>Download</span>
    </div>
  );
}

function stillInTour(target: EventTarget | null): boolean {
  return target instanceof Node && Boolean(
    document.getElementById("home-preview")?.contains(target) ||
    document.getElementById("home-beats")?.contains(target),
  );
}

/** A deck on the table. The top slide is the beat. The sheets behind stay put. */
export function HomePreview({ index, onPause }: { index: number; onPause: (paused: boolean) => void }) {
  const beat = GUIDE_BEATS[index] ?? GUIDE_BEATS[0];
  return (
    <div
      className="home-preview"
      id="home-preview"
      role="tabpanel"
      aria-labelledby={`beat-${beat.id}`}
      onMouseEnter={() => onPause(true)}
      onMouseLeave={(event) => {
        if (!stillInTour(event.relatedTarget)) onPause(false);
      }}
      onFocus={() => onPause(true)}
      onBlur={(event) => {
        if (!stillInTour(event.relatedTarget)) onPause(false);
      }}
    >
      <div className="deck-tilt">
        <div className="deck-sheet is-far" aria-hidden="true" />
        <div className="deck-sheet is-near" aria-hidden="true" />
        <article className={`deck-slide is-${beat.id}`} key={beat.id}>
          <div className="deck-slide-top">
            <span className="deck-kicker">{beat.kicker}</span>
          </div>
          {beat.id === "title" ? (
            <div className="deck-cover">
              <strong>The subject</strong>
              <span>Presenter</span>
            </div>
          ) : (
            <div className="deck-slide-body" aria-hidden="true">
              <GuideScene beat={beat} />
            </div>
          )}
          <div className="deck-slide-foot">
            <p className="home-preview-say">{beat.say}</p>
            <p className="home-preview-detail">{beat.detail}</p>
          </div>
        </article>
        <div className="deck-perch">
          <GuideBot look={beat.look} />
        </div>
      </div>
    </div>
  );
}

/** The five beats, as glass tabs along the stage floor. */
export function HomeBeats({
  index,
  onPick,
  onPause,
}: {
  index: number;
  onPick: (index: number) => void;
  onPause: (paused: boolean) => void;
}) {
  return (
    <div
      className="home-beats"
      id="home-beats"
      role="tablist"
      aria-label="What happens to a deck"
      onMouseEnter={() => onPause(true)}
      onMouseLeave={(event) => {
        if (!stillInTour(event.relatedTarget)) onPause(false);
      }}
      onFocus={() => onPause(true)}
      onBlur={(event) => {
        if (!stillInTour(event.relatedTarget)) onPause(false);
      }}
    >
      {GUIDE_BEATS.map((item, itemIndex) => (
        <button
          key={item.id}
          id={`beat-${item.id}`}
          type="button"
          role="tab"
          aria-selected={itemIndex === index}
          aria-controls="home-preview"
          onClick={() => onPick(itemIndex)}
        >
          <b aria-hidden="true">{item.kicker}</b>
          {item.title}
        </button>
      ))}
    </div>
  );
}
