import { closeOcr, ocrPoolSize, readImageText } from "./ocr.ts";
import { looksLikePdf, parsePdf } from "./pdf.ts";
import { parsePptx } from "./parse.ts";
import { OCR_OFF_DETAIL } from "./types.ts";
import type { Deck, DeckBlock, DeckSlide, ImageResult, ParsedDeck, ParsedSlide, RawBlock } from "./types.ts";

export type ExtractOptions = {
  /** Read the words inside pictures (OCR). On by default. */
  readImages?: boolean;
};

export type ExtractHooks = {
  onStatus: (message: string) => void;
  onProgress?: (done: number, total: number) => void;
  onSlide: (slide: DeckSlide, index: number, total: number) => void;
};

function previewImage(block: Extract<RawBlock, { kind: "image" }>): ImageResult {
  if (block.image.skipReason || !block.image.bytes) {
    return {
      name: block.image.name,
      alt: block.image.alt,
      status: "skipped",
      text: "",
      detail: block.image.skipReason ?? "This image has no data.",
    };
  }
  return {
    name: block.image.name,
    alt: block.image.alt,
    status: "pending",
    text: "",
    detail: "Still reading this image.",
  };
}

function previewSlide(slide: ParsedSlide): DeckSlide {
  const blocks: DeckBlock[] = [];
  if (slide.error) blocks.push({ kind: "error", message: slide.error });
  for (const block of slide.blocks) {
    if (block.kind === "image") {
      blocks.push({ kind: "image", image: previewImage(block) });
      continue;
    }
    const next = withoutGeometry(block);
    if (next) blocks.push(next);
  }
  if (slide.notes.length) blocks.push({ kind: "notes", lines: slide.notes });
  return { index: slide.index, hidden: slide.hidden, title: slide.title, blocks };
}

function withoutGeometry(block: RawBlock): DeckBlock | null {
  if (block.kind === "text") return { kind: "text", role: block.role, lines: block.lines };
  if (block.kind === "table") return { kind: "table", rows: block.rows };
  if (block.kind === "chart") return { kind: "chart" };
  return null;
}

/** Run at most `size` tasks at once, in the order they were queued. */
function limiter(size: number) {
  let active = 0;
  const queue: (() => void)[] = [];
  return async function run<T>(task: () => Promise<T>): Promise<T> {
    if (active >= size) await new Promise<void>((resolve) => queue.push(resolve));
    active += 1;
    try {
      return await task();
    } finally {
      active -= 1;
      queue.shift()?.();
    }
  };
}

/** Cheap content fingerprint, so an image that repeats (logo, header strip) is read once. */
function fingerprint(bytes: Uint8Array, mime: string): string {
  let a = 0x811c9dc5;
  let b = 0x01000193 ^ bytes.length;
  for (let i = 0; i < bytes.length; i++) {
    const value = bytes[i] ?? 0;
    a = Math.imul(a ^ value, 0x01000193);
    b = Math.imul(b ^ value, 0x5bd1e995) ^ (b >>> 15);
  }
  return `${mime}:${bytes.length}:${(a >>> 0).toString(16)}${(b >>> 0).toString(16)}`;
}

type ImageJob = {
  firstSlide: number;
  promise: Promise<Awaited<ReturnType<typeof readImageText>>>;
};

async function materialize(
  parsed: ParsedDeck,
  fileName: string,
  unit: Deck["unit"],
  hooks: ExtractHooks,
  options: ExtractOptions,
): Promise<Deck> {
  const readImages = options.readImages !== false;
  const run = limiter(ocrPoolSize());
  const jobs = new Map<string, ImageJob>();
  let uniqueDone = 0;
  let repeats = 0;
  let preparing = false;

  const report = () => {
    const total = jobs.size;
    if (!total) return;
    const extra = repeats ? ` · ${repeats} repeated ${repeats === 1 ? "image" : "images"} reused` : "";
    hooks.onStatus(`Reading text in images · ${uniqueDone} of ${total} done${extra}`);
    hooks.onProgress?.(uniqueDone, total);
  };

  // Queue every distinct image up front. A few are read at the same time
  // (one per OCR worker), and repeats share the first reading.
  const plan = parsed.slides.map((slide) =>
    slide.blocks.map((block) => {
      if (block.kind !== "image" || !readImages) return null;
      const { bytes, mime, skipReason } = block.image;
      if (skipReason || !bytes) return null;
      const key = fingerprint(bytes, mime);
      const known = jobs.get(key);
      if (known) {
        repeats += 1;
        return { job: known, repeat: known.firstSlide !== slide.index || undefined };
      }
      const job: ImageJob = {
        firstSlide: slide.index,
        promise: run(() =>
          readImageText(bytes, mime, skipReason, (status) => {
            if (!preparing && /load|initializ/i.test(status)) {
              preparing = true;
              hooks.onStatus("Preparing text recognition (the first run downloads an English model)");
            }
          }),
        ).then((result) => {
          uniqueDone += 1;
          report();
          return result;
        }),
      };
      jobs.set(key, job);
      return { job, repeat: undefined };
    }),
  );
  report();

  const slides: DeckSlide[] = [];
  try {
    for (const [slideAt, slide] of parsed.slides.entries()) {
      const blocks: DeckBlock[] = [];
      if (slide.error) blocks.push({ kind: "error", message: slide.error });
      for (const [blockAt, block] of slide.blocks.entries()) {
        if (block.kind !== "image") {
          const next = withoutGeometry(block);
          if (next) blocks.push(next);
          continue;
        }
        const base = { name: block.image.name, alt: block.image.alt };
        const planned = plan[slideAt]?.[blockAt];
        if (!readImages && !block.image.skipReason && block.image.bytes) {
          blocks.push({
            kind: "image",
            image: { ...base, status: "skipped", text: "", detail: OCR_OFF_DETAIL },
          });
          continue;
        }
        if (!planned) {
          blocks.push({ kind: "image", image: previewImage(block) });
          continue;
        }
        const result = await planned.job.promise;
        blocks.push({
          kind: "image",
          image: {
            ...base,
            status: result.status,
            text: result.text,
            detail: result.detail,
            repeatOf: planned.repeat ? planned.job.firstSlide : undefined,
          },
        });
      }
      if (slide.notes.length) blocks.push({ kind: "notes", lines: slide.notes });
      const ready: DeckSlide = {
        index: slide.index,
        hidden: slide.hidden,
        title: slide.title,
        blocks,
      };
      slides.push(ready);
      if (jobs.size === 0) hooks.onProgress?.(slides.length, parsed.slides.length);
      hooks.onSlide(ready, slides.length, parsed.slides.length);
    }
  } finally {
    if (jobs.size) {
      await Promise.allSettled([...jobs.values()].map((job) => job.promise));
      await closeOcr();
    }
  }

  return { fileName, unit, slides };
}

export async function extractDeck(
  buffer: ArrayBuffer,
  fileName: string,
  hooks: ExtractHooks,
  options: ExtractOptions = {},
): Promise<Deck> {
  if (looksLikePdf(buffer, fileName)) {
    hooks.onStatus("Opening the PDF");
    const parsed = await parsePdf(buffer, {
      onStart(index, total) {
        hooks.onStatus(`Reading page ${index} of ${total}`);
        hooks.onProgress?.(index - 1, total);
      },
      onPage(slide, index, total) {
        hooks.onProgress?.(index, total);
        hooks.onSlide(previewSlide(slide), index, total);
      },
    });
    return materialize(parsed, fileName, "page", hooks, options);
  }
  hooks.onStatus("Reading slides");
  return materialize(await parsePptx(buffer), fileName, "slide", hooks, options);
}