import { DeckError } from "./parse.ts";
import type { ParsedDeck, ParsedSlide, RawBlock, TextLine } from "./types.ts";

type PdfJs = {
  getDocument: (src: { data: Uint8Array; verbosity?: number }) => { promise: Promise<PdfDoc> };
  GlobalWorkerOptions: { workerSrc: string };
  OPS: {
    save?: number;
    restore?: number;
    transform?: number;
    paintImageXObject?: number;
    paintInlineImageXObject?: number;
  };
};

type PdfDoc = {
  numPages: number;
  getPage: (index: number) => Promise<PdfPage>;
};

type PdfPage = {
  getTextContent: () => Promise<{
    items: PdfTextItem[];
    styles: Record<string, { fontFamily?: string }>;
  }>;
  getOperatorList: () => Promise<{ fnArray: number[]; argsArray: unknown[][] }>;
  objs: PdfObjects;
  commonObjs: PdfObjects;
  cleanup?: () => boolean;
};

type PdfObjects = {
  get: (name: string, callback?: (obj: PdfImage | null) => void) => PdfImage | null | void;
  has: (name: string) => boolean;
};

type PdfTextItem = {
  str?: string;
  transform?: number[];
  width?: number;
  height?: number;
  fontName?: string;
  hasEOL?: boolean;
};

type PdfImage = {
  width?: number;
  height?: number;
  kind?: number;
  data?: Uint8Array | Uint8ClampedArray;
  bitmap?: CanvasImageSource;
};

export type PdfParseHooks = {
  onStart?: (index: number, total: number) => void;
  onPage?: (slide: ParsedSlide, index: number, total: number) => void;
};

function yieldToUi(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

let pdfjsPromise: Promise<PdfJs> | null = null;

async function loadPdfjs(): Promise<PdfJs> {
  if (!pdfjsPromise) {
    pdfjsPromise = import("pdfjs-dist/legacy/build/pdf.mjs").then(async (mod) => {
      const pdfjs = mod as unknown as PdfJs;
      if (!pdfjs.GlobalWorkerOptions.workerSrc && typeof window !== "undefined") {
        const worker = await import("pdfjs-dist/legacy/build/pdf.worker.min.mjs?url");
        pdfjs.GlobalWorkerOptions.workerSrc = worker.default;
      }
      return pdfjs;
    });
    pdfjsPromise.catch(() => {
      pdfjsPromise = null;
    });
  }
  return pdfjsPromise;
}

export function looksLikePdf(buffer: ArrayBuffer, fileName = ""): boolean {
  const bytes = new Uint8Array(buffer.slice(0, 5));
  if (bytes.length >= 4 && bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) {
    return true;
  }
  return /\.pdf$/i.test(fileName);
}

function multiply(left: number[], right: number[]): number[] {
  const [a1, b1, c1, d1, e1, f1] = left;
  const [a2, b2, c2, d2, e2, f2] = right;
  return [
    a1 * a2 + c1 * b2,
    b1 * a2 + d1 * b2,
    a1 * c2 + c1 * d2,
    b1 * c2 + d1 * d2,
    a1 * e2 + c1 * f2 + e1,
    b1 * e2 + d1 * f2 + f1,
  ];
}

function lineFromItem(item: PdfTextItem, styles: Record<string, { fontFamily?: string }>): TextLine | null {
  const raw = item.str?.replace(/\s+/g, " ").trim() ?? "";
  if (!raw) return null;
  const family = styles[item.fontName ?? ""]?.fontFamily ?? "";
  const bulletMatch = raw.match(/^(?:[•·●▪◦]|[-–—])\s+(.*)$/);
  const text = bulletMatch?.[1]?.trim() || raw;
  if (!text) return null;
  return {
    level: 0,
    bullet: Boolean(bulletMatch),
    runs: [
      {
        text,
        bold: /bold|black/i.test(family),
        italic: /italic|oblique/i.test(family),
      },
    ],
  };
}

function textBlocks(pageText: { items: PdfTextItem[]; styles: Record<string, { fontFamily?: string }> }): RawBlock[] {
  const items = pageText.items.filter((item) => item.str && item.transform);
  const sorted = [...items].sort((a, b) => (b.transform?.[5] ?? 0) - (a.transform?.[5] ?? 0) || (a.transform?.[4] ?? 0) - (b.transform?.[4] ?? 0));
  const rows: PdfTextItem[][] = [];
  for (const item of sorted) {
    const y = item.transform?.[5] ?? 0;
    const height = item.height || 12;
    const row = rows.find((entry) => Math.abs((entry[0]?.transform?.[5] ?? 0) - y) <= Math.max(2, height * 0.45));
    if (row) row.push(item);
    else rows.push([item]);
  }
  const blocks: RawBlock[] = [];
  rows.forEach((row, index) => {
    row.sort((a, b) => (a.transform?.[4] ?? 0) - (b.transform?.[4] ?? 0));
    const parts: TextLine[] = [];
    for (const item of row) {
      const line = lineFromItem(item, pageText.styles);
      if (line) parts.push(line);
    }
    if (!parts.length) return;
    const text = parts.map((line) => line.runs.map((run) => run.text).join("")).join(" ");
    const y = -(row[0]?.transform?.[5] ?? 0);
    const x = row[0]?.transform?.[4] ?? 0;
    blocks.push({
      kind: "text",
      role: index === 0 && text.length <= 80 ? "title" : "body",
      lines: [
        {
          level: 0,
          bullet: parts.every((line) => line.bullet) && parts.length === 1 ? parts[0].bullet : false,
          runs: parts.flatMap((line, partIndex) =>
            partIndex === 0 ? line.runs : [{ text: " ", bold: false, italic: false }, ...line.runs],
          ),
        },
      ],
      x,
      y,
    });
    const only = blocks[blocks.length - 1];
    if (only?.kind === "text" && parts.length === 1 && parts[0]?.bullet) only.lines[0].bullet = true;
  });
  return blocks;
}

function sniffRaster(data: Uint8Array): string | null {
  if (data.length > 3 && data[0] === 0xff && data[1] === 0xd8) return "image/jpeg";
  if (data.length > 8 && data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) return "image/png";
  return null;
}

function toRgba(image: PdfImage): Uint8ClampedArray | null {
  const width = image.width ?? 0;
  const height = image.height ?? 0;
  const data = image.data;
  if (!width || !height || !data) return null;
  const pixels = width * height;
  const out = new Uint8ClampedArray(pixels * 4);
  if (data.length === pixels) {
    for (let i = 0; i < pixels; i++) {
      const value = data[i] ?? 0;
      out[i * 4] = value;
      out[i * 4 + 1] = value;
      out[i * 4 + 2] = value;
      out[i * 4 + 3] = 255;
    }
    return out;
  }
  if ((image.kind === 2 || data.length === pixels * 3) && data.length >= pixels * 3) {
    for (let i = 0; i < pixels; i++) {
      out[i * 4] = data[i * 3] ?? 0;
      out[i * 4 + 1] = data[i * 3 + 1] ?? 0;
      out[i * 4 + 2] = data[i * 3 + 2] ?? 0;
      out[i * 4 + 3] = 255;
    }
    return out;
  }
  if ((image.kind === 3 || data.length >= pixels * 4) && data.length >= pixels * 4) {
    out.set(data.subarray(0, pixels * 4));
    return out;
  }
  return null;
}

async function encodeImage(image: PdfImage): Promise<{ bytes: Uint8Array; mime: string } | null> {
  if (image.data) {
    const mime = sniffRaster(image.data instanceof Uint8Array ? image.data : new Uint8Array(image.data));
    if (mime) return { bytes: image.data instanceof Uint8Array ? image.data : new Uint8Array(image.data), mime };
  }
  if (typeof document === "undefined") return null;
  const width = image.width ?? 0;
  const height = image.height ?? 0;
  if (!width || !height) return null;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) return null;
  if (image.bitmap) {
    context.drawImage(image.bitmap, 0, 0, width, height);
  } else {
    const rgba = toRgba(image);
    if (!rgba) return null;
    const pixels = new Uint8ClampedArray(rgba.byteLength);
    pixels.set(rgba);
    context.putImageData(new ImageData(pixels, width, height), 0, 0);
  }
  const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!blob) return null;
  return { bytes: new Uint8Array(await blob.arrayBuffer()), mime: "image/png" };
}

const IMAGE_WAIT_MS = 15000;

/**
 * pdf.js keeps images that repeat across pages (logos, header strips) in
 * `page.commonObjs` under a `g_` id, not in `page.objs`. Asking `page.objs`
 * for one of those never calls back, which froze the whole run on the first
 * page that reused an image. Look in the right store, and never wait forever.
 */
function imageObject(page: PdfPage, id: string): Promise<PdfImage | null> {
  const store = id.startsWith("g_") ? page.commonObjs : page.objs;
  return new Promise((resolve) => {
    let settled = false;
    const finish = (value: PdfImage | null) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(value);
    };
    const timer = setTimeout(() => finish(null), IMAGE_WAIT_MS);
    try {
      if (store.has(id)) finish((store.get(id) as PdfImage | null) ?? null);
      else store.get(id, (obj) => finish(obj ?? null));
    } catch {
      finish(null);
    }
  });
}

async function imagesOnPage(page: PdfPage, ops: PdfJs["OPS"]): Promise<RawBlock[]> {
  const list = await page.getOperatorList();
  let matrix = [1, 0, 0, 1, 0, 0];
  const stack: number[][] = [];
  const blocks: RawBlock[] = [];
  const seen = new Set<string>();
  let count = 0;
  for (let index = 0; index < list.fnArray.length; index++) {
    const fn = list.fnArray[index];
    const args = list.argsArray[index] ?? [];
    if (fn === ops.save) stack.push(matrix.slice());
    else if (fn === ops.restore) matrix = stack.pop() ?? [1, 0, 0, 1, 0, 0];
    else if (fn === ops.transform && args.length >= 6) matrix = multiply(matrix, args.map(Number));
    else if (fn === ops.paintImageXObject || fn === ops.paintInlineImageXObject) {
      count += 1;
      const key = fn === ops.paintImageXObject ? String(args[0]) : `inline-${count}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const source =
        fn === ops.paintInlineImageXObject
          ? (args[0] as PdfImage)
          : await imageObject(page, String(args[0]));
      if (!source?.width || !source.height || source.width * source.height < MIN_IMAGE_PIXELS) continue;
      const encoded = await encodeImage(source);
      blocks.push({
        kind: "image",
        x: matrix[4] ?? 0,
        y: -(matrix[5] ?? 0),
        image: encoded
          ? { name: `Image ${blocks.length + 1}`, alt: "", mime: encoded.mime, bytes: encoded.bytes }
          : {
              name: `Image ${blocks.length + 1}`,
              alt: "",
              mime: "application/octet-stream",
              bytes: null,
              skipReason: "This image was found, but it could not be prepared for reading.",
            },
      });
    }
  }
  return blocks;
}

const MIN_IMAGE_PIXELS = 2000;

export async function parsePdf(buffer: ArrayBuffer, hooks?: PdfParseHooks): Promise<ParsedDeck> {
  const bytes = new Uint8Array(buffer.slice(0));
  if (!(bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46)) {
    throw new DeckError("That file is not a PDF.");
  }
  const pdfjs = await loadPdfjs();
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    throw new DeckError("The PDF reader could not start in this environment.");
  }
  let doc: PdfDoc;
  try {
    doc = await pdfjs.getDocument({ data: bytes, verbosity: 0 }).promise;
  } catch (error) {
    const name = error instanceof Error ? error.name : "";
    if (name === "PasswordException") throw new DeckError("This PDF is password-protected.");
    throw new DeckError("Could not open that PDF. It may be damaged.");
  }
  if (!doc.numPages) throw new DeckError("This PDF has no pages.");
  const total = doc.numPages;
  const slides: ParsedSlide[] = [];
  for (let index = 1; index <= total; index++) {
    hooks?.onStart?.(index, total);
    await yieldToUi();
    let slide: ParsedSlide;
    try {
      const page = await doc.getPage(index);
      const text = await page.getTextContent();
      const blocks = [...textBlocks(text), ...(await imagesOnPage(page, pdfjs.OPS))];
      blocks.sort((a, b) => a.y - b.y || a.x - b.x);
      const titleBlock = blocks.find((block) => block.kind === "text" && block.role === "title");
      const title =
        titleBlock && titleBlock.kind === "text"
          ? titleBlock.lines.map((line) => line.runs.map((run) => run.text).join("")).join(" ").trim()
          : "";
      slide = { index, hidden: false, title, blocks, notes: [] };
      // Images are already copied out; let pdf.js drop its decoded copies.
      page.cleanup?.();
    } catch (error) {
      slide = {
        index,
        hidden: false,
        title: "",
        blocks: [],
        notes: [],
        error: error instanceof DeckError ? error.message : "Could not read this page.",
      };
    }
    slides.push(slide);
    hooks?.onPage?.(slide, index, total);
    await yieldToUi();
  }
  return { slides };
}
