type OcrLine = { text: string; confidence: number };
type OcrBlock = { paragraphs?: { lines?: OcrLine[] }[] };
type OcrResult = { data: { text: string; confidence?: number; blocks?: OcrBlock[] | null } };

type OcrWorker = {
  recognize: (
    image: CanvasImageSource | Blob,
    options?: Record<string, unknown>,
    output?: Record<string, boolean>,
  ) => Promise<OcrResult>;
  terminate: () => Promise<unknown>;
};

export type ImageTextResult = { status: "text" | "empty" | "skipped" | "error"; text: string; detail?: string };

/** How long one image may take before we give up on it and move on. */
const IMAGE_TIMEOUT_MS = 60_000;
/** Lines Tesseract is less sure about than this are dropped as noise. */
const MIN_LINE_CONFIDENCE = 45;
/** Images are scaled so their longest edge is at most this many pixels. */
const MAX_EDGE = 1800;
/** Very small images are scaled up so letters are big enough to read. */
const MIN_EDGE = 300;

let sink: (status: string, progress: number) => void = () => {};

export function ocrPoolSize(): number {
  const cores = typeof navigator !== "undefined" ? navigator.hardwareConcurrency || 2 : 2;
  return Math.max(1, Math.min(4, Math.floor(cores / 2)));
}

/**
 * A small pool of Tesseract workers. Each worker loads the English model once
 * and is reused for every image. Images are handed to the next free worker,
 * so several are read at the same time on multi-core machines.
 */
class OcrPool {
  private idle: OcrWorker[] = [];
  private waiting: { resolve: (worker: OcrWorker) => void; reject: (error: unknown) => void }[] = [];
  private starting = 0;
  private alive = 0;

  constructor(private size: number) {}

  private async spawn(): Promise<OcrWorker> {
    const { createWorker } = await import("tesseract.js");
    const worker = (await createWorker("eng", 1, {
      logger: (message: { status: string; progress?: number }) => {
        if (message.status !== "recognizing text") sink(message.status, message.progress ?? 0);
      },
    })) as unknown as OcrWorker;
    return worker;
  }

  private acquire(): Promise<OcrWorker> {
    const ready = this.idle.pop();
    if (ready) return Promise.resolve(ready);
    if (this.alive + this.starting < this.size) {
      this.starting += 1;
      return this.spawn().then(
        (worker) => {
          this.starting -= 1;
          this.alive += 1;
          return worker;
        },
        (error) => {
          this.starting -= 1;
          throw error;
        },
      );
    }
    return new Promise((resolve, reject) => this.waiting.push({ resolve, reject }));
  }

  private release(worker: OcrWorker) {
    const next = this.waiting.shift();
    if (next) next.resolve(worker);
    else this.idle.push(worker);
  }

  private discard(worker: OcrWorker) {
    this.alive -= 1;
    void worker.terminate().catch(() => {});
    // Someone may be waiting for a worker; let them start a fresh one.
    const next = this.waiting.shift();
    if (next) void this.acquire().then(next.resolve, next.reject);
  }

  async run(source: CanvasImageSource | Blob): Promise<OcrResult> {
    const worker = await this.acquire();
    let timer: ReturnType<typeof setTimeout> | undefined;
    try {
      const result = await Promise.race([
        worker.recognize(source, {}, { text: true, blocks: true }),
        new Promise<never>((_, reject) => {
          timer = setTimeout(() => reject(new Error("Reading this image took too long, so it was skipped.")), IMAGE_TIMEOUT_MS);
        }),
      ]);
      clearTimeout(timer);
      this.release(worker);
      return result;
    } catch (error) {
      clearTimeout(timer);
      // A worker that failed or timed out may be stuck; replace it.
      this.discard(worker);
      throw error;
    }
  }

  async close() {
    const workers = this.idle.splice(0);
    this.alive -= workers.length;
    await Promise.all(workers.map((worker) => worker.terminate().catch(() => {})));
  }
}

let pool: OcrPool | null = null;

function getPool(): OcrPool {
  if (!pool) pool = new OcrPool(ocrPoolSize());
  return pool;
}

/** Free the OCR workers (and their memory) once a file is done. */
export async function closeOcr(): Promise<void> {
  const current = pool;
  pool = null;
  await current?.close();
}

function blobFrom(bytes: Uint8Array, mime: string): Blob {
  const copy = bytes.slice();
  return new Blob([copy.buffer], { type: mime });
}

async function rasterToCanvas(bytes: Uint8Array, mime: string): Promise<HTMLCanvasElement> {
  const bitmap = await createImageBitmap(blobFrom(bytes, mime));
  try {
    const longest = Math.max(bitmap.width, bitmap.height);
    let scale = Math.min(1, MAX_EDGE / longest);
    if (longest < MIN_EDGE) scale = Math.min(3, MIN_EDGE / longest);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(bitmap.width * scale));
    canvas.height = Math.max(1, Math.round(bitmap.height * scale));
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Could not prepare the image.");
    // White behind transparent pixels, so dark text on a clear background is readable.
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.imageSmoothingQuality = "high";
    context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    return canvas;
  } finally {
    bitmap.close();
  }
}

function extractSvgText(svg: string): string {
  const doc = new DOMParser().parseFromString(svg, "image/svg+xml");
  return [...doc.querySelectorAll("text")]
    .map((node) => node.textContent?.replace(/\s+/g, " ").trim() ?? "")
    .filter(Boolean)
    .join("\n");
}

export function cleanupOcr(raw: string): string {
  return raw
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** A line is worth keeping if Tesseract trusts it and it has at least one real word. */
function usefulLine(line: OcrLine): boolean {
  if (line.confidence < MIN_LINE_CONFIDENCE) return false;
  return /[\p{L}\p{N}]{2,}/u.test(line.text);
}

export function textFromResult(result: OcrResult): string {
  const blocks = result.data.blocks;
  if (!blocks?.length) return cleanupOcr(result.data.text ?? "");
  const paragraphs: string[] = [];
  for (const block of blocks) {
    for (const paragraph of block.paragraphs ?? []) {
      const lines = (paragraph.lines ?? []).filter(usefulLine).map((line) => line.text.trim());
      if (lines.length) paragraphs.push(lines.join("\n"));
    }
  }
  return cleanupOcr(paragraphs.join("\n\n"));
}

const RASTER = new Set(["image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp", "image/tiff"]);

export async function readImageText(
  bytes: Uint8Array | null,
  mime: string,
  skipReason: string | undefined,
  onProgress: (status: string, progress: number) => void,
): Promise<ImageTextResult> {
  if (skipReason || !bytes) {
    return { status: "skipped", text: "", detail: skipReason ?? "This image has no data." };
  }
  if (mime === "image/svg+xml") {
    const text = extractSvgText(new TextDecoder().decode(bytes));
    return text
      ? { status: "text", text }
      : { status: "skipped", text: "", detail: "This SVG has no text layer." };
  }
  if (!RASTER.has(mime)) {
    return {
      status: "skipped",
      text: "",
      detail: "This image format cannot be read for text. Raster images such as PNG and JPEG can.",
    };
  }
  sink = onProgress;
  try {
    let source: HTMLCanvasElement | Blob;
    try {
      source = await rasterToCanvas(bytes, mime);
    } catch {
      source = blobFrom(bytes, mime);
    }
    const result = await getPool().run(source);
    const text = textFromResult(result);
    return text ? { status: "text", text } : { status: "empty", text: "" };
  } catch (error) {
    return {
      status: "error",
      text: "",
      detail: error instanceof Error ? error.message : "Could not read text from this image.",
    };
  }
}
