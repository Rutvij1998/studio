# Studio

Drop a PowerPoint. Leave with a restyled deck.

Live site: [deck-studio-red.vercel.app](https://deck-studio-red.vercel.app)

## What it does

Upload one `.pptx`. An older `.ppt` is turned away.

The restyle keeps every source slide. Wording and numbers stay as they were uploaded. Nothing is invented.

- Slide 1 is a dark title slide: the subject and the presenter name, with a particle wave.
- Every later slide is white.
- A list stays a list. A picture stays a picture.
- A native Office chart is redrawn from the cells stored in the file, in the studio colors.
- A chart that exists only as an image is kept as that image.

After a restyle, the deck opens full screen. Arrow keys move between slides. Download writes an editable PowerPoint.

**Compare** places each uploaded slide beside the restyle. Press `C`. On the public site this view cannot open the original pages, because that step runs on your machine.

**Grade** scores the deck from the slide structure. Press `G`. A thin slide is listed. The letter is computed in the app. A note that would add a number is dropped.

## Run it

```bash
npm install
npm run dev
```

Open [http://127.0.0.1:4310](http://127.0.0.1:4310).

```bash
npm run typecheck
npm run build
```

## Local only

These routes exist in the Vite dev server. They are not deployed with the public site.

| Route | What it does |
| --- | --- |
| `POST /api/original` | Turns the uploaded file into slide images for Compare. Needs LibreOffice (`soffice`) and Python with PyMuPDF. |
| `POST /api/grade` | May rephrase a grade note. The letter stays the one computed in the app. |

For grade notes, put `XAI_API_KEY` in `.env.local`. That file is gitignored. The key stays on the server. Without the key, the letter and the written notes still show.

## Stack

Vite, React, and TypeScript. The download is built with [PptxGenJS](https://github.com/gitbrent/PptxGenJS). On screen, titles are Outfit and body text is Plus Jakarta Sans. The PowerPoint file uses Calibri.
