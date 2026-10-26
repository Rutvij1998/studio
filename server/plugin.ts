import { spawn } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { mkdir, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import type { IncomingMessage, ServerResponse } from "node:http";
import { tmpdir } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import type { Connect, Plugin, PreviewServer, ViteDevServer } from "vite";

const LOGO_DIR = "/Users/rutvij/producthub-platform/hub/frontend/public/product-hub/partners/logos";

const SEE_GUIDE = `Read only the numbers printed in this picture. Return one JSON object and nothing else.
{"title":"","categories":[],"series":[{"name":"","values":[]}],"numbers":[{"label":"","display":"","value":0}]}
Copy every digit as printed. Do not add a month, a total, or a cause. If the picture has no figures, return {"title":"","categories":[],"series":[],"numbers":[]}.`;

type Loader = {
  ssrLoadModule: (url: string) => Promise<Record<string, unknown>>;
  middlewares: Connect.Server;
};

function envValue(name: string): string {
  const live = process.env[name]?.trim();
  if (live) return live;
  const files = [
    "/Users/rutvij/deck-studio/.env",
    "/Users/rutvij/people-vantage/.env.local",
    "/Users/rutvij/market-intel/.env.local",
    "/Users/rutvij/producthub-platform/hub/frontend/.env.local",
  ];
  const pattern = new RegExp(`^${name}=(.*)$`, "m");
  for (const file of files) {
    if (!existsSync(file)) continue;
    const match = readFileSync(file, "utf8").match(pattern);
    const value = match?.[1]?.trim().replace(/^['"]|['"]$/g, "");
    if (value) return value;
  }
  return "";
}

function send(res: ServerResponse, status: number, body: unknown) {
  const payload = JSON.stringify(body);
  res.statusCode = status;
  res.setHeader("content-type", "application/json");
  res.end(payload);
}

async function readRaw(req: IncomingMessage, limit: number): Promise<Buffer> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    size += buffer.length;
    if (size > limit) throw new Error("large");
    chunks.push(buffer);
  }
  return Buffer.concat(chunks);
}

async function readBody(req: IncomingMessage, limit: number): Promise<string> {
  return (await readRaw(req, limit)).toString("utf8");
}

function runProcess(command: string, args: string[], cwd: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd, stdio: ["ignore", "ignore", "pipe"] });
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      reject(new Error("slow"));
    }, 50000);
    child.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(new Error("convert"));
    });
  });
}

let rasterChain: Promise<void> = Promise.resolve();

/** Draw each slide of the uploaded file. Local only. The pages stay in file order. */
async function rasterSlides(bytes: Buffer): Promise<string[]> {
  const run = rasterChain.then(() => drawSlides(bytes));
  rasterChain = run.then(() => undefined, () => undefined);
  return run;
}

async function drawSlides(bytes: Buffer): Promise<string[]> {
  const dir = await mkdtemp(path.join(tmpdir(), "studio-original-"));
  try {
    const pptx = path.join(dir, "deck.pptx");
    const pages = path.join(dir, "pages");
    await mkdir(pages);
    await mkdir(path.join(dir, "profile"));
    await writeFile(pptx, bytes);
    const soffice = existsSync("/opt/homebrew/bin/soffice") ? "/opt/homebrew/bin/soffice" : "soffice";
    const profile = pathToFileURL(path.join(dir, "profile")).href;
    await runProcess(soffice, [
      "--headless",
      "--norestore",
      "--nolockcheck",
      `-env:UserInstallation=${profile}`,
      "--convert-to",
      "pdf",
      "--outdir",
      dir,
      pptx,
    ], dir);
    const pdfName = (await readdir(dir)).find((name) => name.toLowerCase().endsWith(".pdf"));
    if (!pdfName) return [];
    await runProcess("python3", ["-c", [
      "import fitz, sys",
      "doc = fitz.open(sys.argv[1])",
      "for index, page in enumerate(doc):",
      "    zoom = 1600 / max(page.rect.width, 1)",
      "    page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False).save(f'{sys.argv[2]}/{index + 1:02d}.png')",
    ].join("\n"), path.join(dir, pdfName), pages], dir);
    const names = (await readdir(pages)).filter((name) => name.endsWith(".png")).sort();
    const slides: string[] = [];
    for (const name of names.slice(0, 80)) {
      const png = await readFile(path.join(pages, name));
      slides.push(`data:image/png;base64,${png.toString("base64")}`);
    }
    return slides;
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

async function grok(input: unknown, timeout = 45000): Promise<unknown> {
  const key = envValue("XAI_API_KEY");
  if (!key) {
    const error = new Error("key");
    error.name = "MissingKey";
    throw error;
  }
  const response = await fetch("https://api.x.ai/v1/responses", {
    method: "POST",
    headers: { authorization: `Bearer ${key}`, "content-type": "application/json" },
    body: JSON.stringify({ model: "grok-4.7", store: false, input }),
    signal: AbortSignal.timeout(timeout),
  });
  if (!response.ok) throw new Error("model");
  return response.json();
}

function dataUrl(bytes: Buffer, mime: string): string {
  return `data:${mime};base64,${bytes.toString("base64")}`;
}

async function logoDevPng(domain: string): Promise<string> {
  const key = envValue("VITE_LOGO_DEV_PUBLISHABLE_KEY");
  if (!key || !/^[a-z0-9.-]+$/.test(domain)) return "";
  const response = await fetch(`https://img.logo.dev/${domain}?token=${encodeURIComponent(key)}&format=png&size=128&fallback=404`, {
    signal: AbortSignal.timeout(12000),
  });
  if (!response.ok) return "";
  const type = response.headers.get("content-type") ?? "";
  if (!type.includes("png")) return "";
  const bytes = Buffer.from(await response.arrayBuffer());
  if (bytes.length < 80) return "";
  return dataUrl(bytes, "image/png");
}

function localMark(file: string): string {
  if (!/^[a-z0-9]+(?:[.-][a-z0-9]+)*\.(svg|png)$/.test(file) || file.includes("..")) return "";
  const full = path.resolve(LOGO_DIR, file);
  if (!full.startsWith(`${LOGO_DIR}${path.sep}`)) return "";
  if (!existsSync(full)) return "";
  const bytes = readFileSync(full);
  const mime = file.endsWith(".png") ? "image/png" : "image/svg+xml";
  return dataUrl(bytes, mime);
}

async function markFor(loader: Loader, name: string): Promise<string> {
  const mod = await loader.ssrLoadModule("/src/mark.ts");
  const markRequest = mod.markRequest as (value: string) => { domain: string | null; file: string | null } | null;
  const request = markRequest(name);
  if (!request) return "";
  if (request.file === "ee-flex.svg") {
    const local = localMark(request.file);
    if (local) return local;
  }
  if (request.domain) {
    const remote = await logoDevPng(request.domain);
    if (remote) return remote;
  }
  if (request.file) return localMark(request.file);
  return "";
}

async function shapeFor(loader: Loader, text: string): Promise<{ prompt?: string; error?: string }> {
  const mod = await loader.ssrLoadModule("/src/shape.ts");
  const guide = String(mod.SHAPE_GUIDE ?? "");
  const textFromResponse = mod.textFromResponse as (body: unknown) => string;
  const acceptShape = mod.acceptShape as (source: string, shaped: string) => string | null;
  try {
    const body = await grok([
      {
        role: "user",
        content: [{ type: "input_text", text: `${guide}\n\nNote:\n${text.slice(0, 12000)}` }],
      },
    ]);
    const prompt = acceptShape(text, textFromResponse(body));
    if (!prompt) return { error: "A figure in the note changed, so the original was kept." };
    return { prompt };
  } catch (error) {
    if (error instanceof Error && error.name === "MissingKey") {
      return { error: "The language step needs an API key on the server." };
    }
    return { error: "The note stayed as written." };
  }
}

async function seeFor(loader: Loader, slides: { index: number; mime: string; base64: string }[]): Promise<{ index: number; text: string }[]> {
  const mod = await loader.ssrLoadModule("/src/shape.ts");
  const textFromResponse = mod.textFromResponse as (body: unknown) => string;
  const reads: { index: number; text: string }[] = [];
  for (const slide of slides.slice(0, 6)) {
    if (slide.mime !== "image/png" && slide.mime !== "image/jpeg") continue;
    if (!slide.base64 || slide.base64.length > 8_000_000) continue;
    try {
      const body = await grok([
        {
          role: "user",
          content: [
            { type: "input_image", image_url: `data:${slide.mime};base64,${slide.base64}`, detail: "high" },
            { type: "input_text", text: SEE_GUIDE },
          ],
        },
      ]);
      const text = textFromResponse(body);
      if (text) reads.push({ index: slide.index, text });
    } catch (error) {
      if (error instanceof Error && error.name === "MissingKey") break;
    }
  }
  return reads;
}

type GradeRow = { index: number; title: string; text: string; problem: string };

async function directFor(loader: Loader, raw: unknown): Promise<{ drop: string[]; keep: number[]; slides: unknown[] }> {
  const director = await loader.ssrLoadModule("/src/director.ts");
  const shape = await loader.ssrLoadModule("/src/shape.ts");
  const packFacts = director.packFacts as (value: unknown) => unknown;
  const guide = String(director.DIRECTOR_GUIDE ?? "");
  const textFromResponse = shape.textFromResponse as (body: unknown) => string;
  const packed = packFacts(raw);
  try {
    const body = await grok([
      {
        role: "user",
        content: [{ type: "input_text", text: `${guide}\n\nDeck:\n${JSON.stringify(packed)}` }],
      },
    ], 90000);
    const text = textFromResponse(body);
    const start = text.indexOf("{");
    const end = text.lastIndexOf("}");
    if (start < 0 || end <= start) return { drop: [], keep: [], slides: [] };
    const parsed = JSON.parse(text.slice(start, end + 1)) as { drop?: unknown; keep?: unknown; slides?: unknown };
    const drop = Array.isArray(parsed.drop) ? parsed.drop.filter((item): item is string => typeof item === "string").slice(0, 8) : [];
    const keep = Array.isArray(parsed.keep) ? parsed.keep.filter((item): item is number => typeof item === "number").slice(0, 24) : [];
    const slides = Array.isArray(parsed.slides) ? parsed.slides.slice(0, 24) : [];
    return { drop, keep, slides };
  } catch {
    return { drop: [], keep: [], slides: [] };
  }
}

async function gradeFor(loader: Loader, rows: GradeRow[]): Promise<{ summary: string; notes: { index: number; note: string }[] }> {
  const critic = await loader.ssrLoadModule("/src/critic.ts");
  const shape = await loader.ssrLoadModule("/src/shape.ts");
  const guide = String(critic.GRADE_GUIDE ?? "");
  const acceptNote = critic.acceptNote as (source: string, note: string) => string | null;
  const textFromResponse = shape.textFromResponse as (body: unknown) => string;
  const packed = rows.slice(0, 40).map((row) => ({
    index: row.index,
    title: row.title.slice(0, 90),
    text: row.text.slice(0, 280),
    problem: row.problem.slice(0, 180),
  }));
  try {
    const body = await grok([
      {
        role: "user",
        content: [{ type: "input_text", text: `${guide}\n\nSlides:\n${JSON.stringify(packed)}` }],
      },
    ]);
    const raw = textFromResponse(body);
    const start = raw.indexOf("{");
    const end = raw.lastIndexOf("}");
    if (start < 0 || end <= start) return { summary: "", notes: [] };
    const parsed = JSON.parse(raw.slice(start, end + 1)) as { summary?: unknown; slides?: unknown };
    const source = packed.map((row) => `${row.title} ${row.text}`).join("\n");
    const summary = typeof parsed.summary === "string" ? acceptNote(source, parsed.summary) ?? "" : "";
    const notes: { index: number; note: string }[] = [];
    const slides = Array.isArray(parsed.slides) ? parsed.slides : [];
    for (const item of slides) {
      if (!item || typeof item !== "object" || notes.length >= 8) continue;
      const row = item as { index?: unknown; note?: unknown };
      if (typeof row.index !== "number" || typeof row.note !== "string") continue;
      const known = packed.find((slide) => slide.index === row.index);
      if (!known) continue;
      const note = acceptNote(`${known.title} ${known.text}`, row.note);
      if (note) notes.push({ index: known.index, note });
    }
    return { summary, notes };
  } catch {
    return { summary: "", notes: [] };
  }
}

function attach(server: Loader) {
  server.middlewares.use(async (req, res, next) => {
    const url = req.url?.split("?")[0] ?? "";
    if (!url.startsWith("/api/")) return next();
    try {
      if (req.method === "POST" && url === "/api/mark") {
        const body = JSON.parse(await readBody(req, 8_000)) as { name?: unknown };
        const name = typeof body.name === "string" ? body.name : "";
        const dataUrlValue = name ? await markFor(server, name) : "";
        send(res, 200, { dataUrl: dataUrlValue });
        return;
      }
      if (req.method === "POST" && url === "/api/shape") {
        const body = JSON.parse(await readBody(req, 200_000)) as { text?: unknown };
        const text = typeof body.text === "string" ? body.text : "";
        if (text.trim().length < 8) {
          send(res, 200, { error: "The note stayed as written." });
          return;
        }
        send(res, 200, await shapeFor(server, text));
        return;
      }
      if (req.method === "POST" && url === "/api/original") {
        const bytes = await readRaw(req, 40_000_000);
        if (bytes.length < 64) {
          send(res, 200, { slides: [] });
          return;
        }
        send(res, 200, { slides: await rasterSlides(bytes) });
        return;
      }
      if (req.method === "POST" && url === "/api/direct") {
        const body = JSON.parse(await readBody(req, 200_000)) as unknown;
        if (!envValue("XAI_API_KEY")) {
          send(res, 200, { drop: [], keep: [], slides: [] });
          return;
        }
        send(res, 200, await directFor(server, body));
        return;
      }
      if (req.method === "POST" && url === "/api/grade") {
        const body = JSON.parse(await readBody(req, 200_000)) as { slides?: unknown };
        const rows = Array.isArray(body.slides) ? body.slides.flatMap((item) => {
          if (!item || typeof item !== "object") return [];
          const row = item as { index?: unknown; title?: unknown; text?: unknown; problem?: unknown };
          if (typeof row.index !== "number") return [];
          return [{
            index: row.index,
            title: typeof row.title === "string" ? row.title : "",
            text: typeof row.text === "string" ? row.text : "",
            problem: typeof row.problem === "string" ? row.problem : "",
          }];
        }) : [];
        if (!envValue("XAI_API_KEY")) {
          send(res, 200, { summary: "", notes: [] });
          return;
        }
        send(res, 200, await gradeFor(server, rows));
        return;
      }
      if (req.method === "POST" && url === "/api/see") {
        const body = JSON.parse(await readBody(req, 12_000_000)) as { slides?: unknown };
        const slides = Array.isArray(body.slides) ? body.slides : [];
        const clean = slides
          .map((item) => {
            if (!item || typeof item !== "object") return null;
            const row = item as { index?: unknown; mime?: unknown; base64?: unknown };
            if (typeof row.index !== "number" || typeof row.mime !== "string" || typeof row.base64 !== "string") return null;
            return { index: row.index, mime: row.mime, base64: row.base64 };
          })
          .filter((item): item is { index: number; mime: string; base64: string } => item != null);
        if (!envValue("XAI_API_KEY")) {
          send(res, 200, { slides: [], error: "The language step needs an API key on the server." });
          return;
        }
        send(res, 200, { slides: await seeFor(server, clean) });
        return;
      }
      send(res, 404, { error: "Not found." });
    } catch {
      send(res, 400, { error: "That request could not be read." });
    }
  });
}

export function studioPlugin(): Plugin {
  return {
    name: "studio-language",
    configureServer(server) {
      attach(server as ViteDevServer);
    },
    configurePreviewServer(server) {
      attach(server as PreviewServer);
    },
  };
}
