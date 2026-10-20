import JSZip from "jszip";
import type { ImagePayload, ParsedDeck, ParsedSlide, RawBlock, TextLine, TextRun } from "./types.ts";
import {
  attrLocal,
  children,
  findChild,
  findDesc,
  findRoot,
  localName,
  parseXml,
  plainLine,
  textContent,
  type XmlNode,
} from "./xml.ts";

export class DeckError extends Error {}

const SKIP_PLACEHOLDERS = new Set(["dt", "ftr", "sldNum", "sldImg", "hdr"]);
const RASTER_EXT = new Set(["png", "jpg", "jpeg", "gif", "bmp", "webp", "tif", "tiff"]);

type Rel = { type: string; target: string; external: boolean };

type WalkCtx = {
  zip: JSZip;
  slidePart: string;
  rels: Record<string, Rel>;
  allowImages: boolean;
};

function isOle(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0;
}

function isZip(bytes: Uint8Array): boolean {
  return bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

function zipFile(zip: JSZip, path: string): JSZip.JSZipObject | null {
  const normalized = path.replace(/^\/+/, "");
  const direct = zip.file(normalized);
  if (direct && !direct.dir) return direct;
  const lower = normalized.toLowerCase();
  return zip.file(/.*/).find((file) => !file.dir && file.name.replace(/^\/+/, "").toLowerCase() === lower) ?? null;
}

async function readText(zip: JSZip, path: string): Promise<string | null> {
  const file = zipFile(zip, path);
  if (!file) return null;
  return file.async("string");
}

function relsPath(part: string): string {
  const slash = part.lastIndexOf("/");
  const dir = slash === -1 ? "" : part.slice(0, slash);
  const file = slash === -1 ? part : part.slice(slash + 1);
  return `${dir}/_rels/${file}.rels`;
}

function resolvePath(fromPart: string, target: string): string {
  if (target.startsWith("/")) return target.replace(/^\/+/, "");
  const parts = fromPart.split("/").slice(0, -1);
  for (const segment of target.split("/")) {
    if (!segment || segment === ".") continue;
    if (segment === "..") parts.pop();
    else parts.push(segment);
  }
  return parts.join("/");
}

function parseRels(xml: string): Record<string, Rel> {
  const root = findRoot(parseXml(xml), "Relationships");
  const map: Record<string, Rel> = {};
  if (!root) return map;
  for (const node of children(root)) {
    if (localName(node) !== "Relationship") continue;
    const id = attrLocal(node, "Id", false);
    if (!id) continue;
    map[id] = {
      type: attrLocal(node, "Type", false) ?? "",
      target: attrLocal(node, "Target", false) ?? "",
      external: attrLocal(node, "TargetMode", false) === "External",
    };
  }
  return map;
}

function extOf(path: string): string {
  const clean = path.split("?")[0] ?? path;
  const dot = clean.lastIndexOf(".");
  return dot === -1 ? "" : clean.slice(dot + 1).toLowerCase();
}

function mimeFromExt(ext: string): string {
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "webp":
      return "image/webp";
    case "tif":
    case "tiff":
      return "image/tiff";
    case "svg":
      return "image/svg+xml";
    case "emf":
      return "image/emf";
    case "wmf":
      return "image/wmf";
    default:
      return "application/octet-stream";
  }
}

function shapeOffset(node: XmlNode): { x: number; y: number; chx: number; chy: number } {
  const holder = findChild(node, "spPr") ?? findChild(node, "grpSpPr");
  const xfrm = (holder && findChild(holder, "xfrm")) ?? findChild(node, "xfrm");
  if (!xfrm) return { x: 0, y: 0, chx: 0, chy: 0 };
  const off = findChild(xfrm, "off");
  const chOff = findChild(xfrm, "chOff");
  return {
    x: Number(off ? attrLocal(off, "x", false) : 0) || 0,
    y: Number(off ? attrLocal(off, "y", false) : 0) || 0,
    chx: Number(chOff ? attrLocal(chOff, "x", false) : 0) || 0,
    chy: Number(chOff ? attrLocal(chOff, "y", false) : 0) || 0,
  };
}

function placeholderType(node: XmlNode): string | null {
  const nv = findChild(node, "nvSpPr") ?? findChild(node, "nvPicPr") ?? findChild(node, "nvGraphicFramePr");
  const ph = nv ? findDesc(nv, "ph") : undefined;
  if (!ph) return null;
  return attrLocal(ph, "type", false) ?? "body";
}

function roleOf(ph: string | null): "title" | "subtitle" | "body" {
  if (ph === "title" || ph === "ctrTitle") return "title";
  if (ph === "subTitle") return "subtitle";
  return "body";
}

function readRun(node: XmlNode, lines: TextRun[][]) {
  let bold = false;
  let italic = false;
  for (const child of children(node)) {
    const name = localName(child);
    if (name === "rPr") {
      const b = attrLocal(child, "b", false);
      const i = attrLocal(child, "i", false);
      bold = b === "1" || b === "true";
      italic = i === "1" || i === "true";
    } else if (name === "br") {
      lines.push([]);
    } else if (name === "t") {
      const text = textContent(child);
      if (text) lines[lines.length - 1]?.push({ text, bold, italic });
    }
  }
}

function readParagraph(node: XmlNode): TextLine[] {
  const pPr = findChild(node, "pPr");
  const level = Math.max(0, Number(pPr ? attrLocal(pPr, "lvl", false) : 0) || 0);
  let bullet = false;
  if (pPr) {
    const marks = children(pPr).map((child) => localName(child));
    const hasNone = marks.includes("buNone");
    const hasBullet = marks.some((mark) => mark === "buChar" || mark === "buAutoNum" || mark === "buBlip");
    bullet = hasBullet && !hasNone;
  }
  const chunks: TextRun[][] = [[]];
  for (const child of children(node)) {
    const name = localName(child);
    if (name === "br") chunks.push([]);
    else if (name === "r" || name === "fld") readRun(child, chunks);
  }
  return chunks
    .map((runs) => ({ level, bullet, runs }))
    .filter((line) => line.runs.some((run) => run.text.trim()));
}

function readTxBody(node: XmlNode | undefined): TextLine[] {
  if (!node) return [];
  const lines: TextLine[] = [];
  for (const child of children(node)) {
    if (localName(child) !== "p") continue;
    lines.push(...readParagraph(child));
  }
  return lines;
}

function readTable(tbl: XmlNode): string[][] {
  const rows: string[][] = [];
  for (const tr of children(tbl)) {
    if (localName(tr) !== "tr") continue;
    const row: string[] = [];
    for (const tc of children(tr)) {
      if (localName(tc) !== "tc") continue;
      const lines = readTxBody(findChild(tc, "txBody"));
      row.push(
        lines
          .map((line) => plainLine(line))
          .filter(Boolean)
          .join(" ")
          .replace(/\s+/g, " ")
          .trim(),
      );
    }
    if (row.length) rows.push(row);
  }
  return rows;
}

function collectEmbeds(node: XmlNode, out: string[]) {
  const name = localName(node);
  if (name === "blip" || name === "svgBlip") {
    const id = attrLocal(node, "embed", true);
    if (id) out.push(id);
  }
  for (const child of children(node)) {
    if (localName(child) === "#text") continue;
    collectEmbeds(child, out);
  }
}

function pictureMeta(pic: XmlNode): { name: string; alt: string } {
  const cNvPr = findDesc(pic, "cNvPr");
  return {
    name: (cNvPr && attrLocal(cNvPr, "name", false)) || "Image",
    alt: (cNvPr && attrLocal(cNvPr, "descr", false)) || "",
  };
}

function rankEmbed(target: string): number {
  const ext = extOf(target);
  if (RASTER_EXT.has(ext)) return 3;
  if (ext === "svg") return 2;
  return 1;
}

async function materializeImage(ids: string[], meta: { name: string; alt: string }, ctx: WalkCtx): Promise<ImagePayload> {
  const linked = ids
    .map((id) => ({ id, rel: ctx.rels[id] }))
    .filter((entry): entry is { id: string; rel: Rel } => !!entry.rel);
  const embedded = linked.filter((entry) => !entry.rel.external);
  embedded.sort((a, b) => rankEmbed(b.rel.target) - rankEmbed(a.rel.target));
  const chosen = embedded[0] ?? linked[0];
  if (!chosen) {
    return {
      ...meta,
      mime: "application/octet-stream",
      bytes: null,
      skipReason: "The picture link inside the slide is broken.",
    };
  }
  if (chosen.rel.external) {
    return {
      ...meta,
      mime: mimeFromExt(extOf(chosen.rel.target)),
      bytes: null,
      skipReason: "This image is linked from outside the file, so its text was not read.",
    };
  }
  const path = resolvePath(ctx.slidePart, chosen.rel.target);
  const file = zipFile(ctx.zip, path);
  if (!file) {
    return {
      ...meta,
      mime: mimeFromExt(extOf(chosen.rel.target)),
      bytes: null,
      skipReason: "The image is missing from the PowerPoint file.",
    };
  }
  const bytes = await file.async("uint8array");
  return { ...meta, mime: mimeFromExt(extOf(chosen.rel.target)), bytes };
}

async function walkShapes(nodes: XmlNode[], ox: number, oy: number, blocks: RawBlock[], ctx: WalkCtx) {
  for (const node of nodes) {
    const name = localName(node);
    if (!name || name === "#text") continue;
    if (name === "AlternateContent") {
      const fallback = children(node).find((child) => localName(child) === "Fallback");
      const choice = children(node).find((child) => localName(child) === "Choice");
      await walkShapes(children(fallback ?? choice ?? node), ox, oy, blocks, ctx);
      continue;
    }
    if (name === "grpSp") {
      const group = shapeOffset(node);
      await walkShapes(children(node), ox + group.x - group.chx, oy + group.y - group.chy, blocks, ctx);
      continue;
    }
    if (name === "sp" || name === "cxnSp") {
      const ph = placeholderType(node);
      if (ph && SKIP_PLACEHOLDERS.has(ph)) continue;
      const lines = readTxBody(findChild(node, "txBody"));
      if (!lines.length) continue;
      const off = shapeOffset(node);
      blocks.push({ kind: "text", role: roleOf(ph), lines, x: ox + off.x, y: oy + off.y });
      continue;
    }
    if (name === "pic") {
      if (!ctx.allowImages) continue;
      const off = shapeOffset(node);
      const ids: string[] = [];
      collectEmbeds(node, ids);
      const image = await materializeImage(ids, pictureMeta(node), ctx);
      blocks.push({ kind: "image", image, x: ox + off.x, y: oy + off.y });
      continue;
    }
    if (name === "graphicFrame") {
      const off = shapeOffset(node);
      const x = ox + off.x;
      const y = oy + off.y;
      const table = findDesc(node, "tbl");
      if (table) {
        const rows = readTable(table);
        if (rows.length) blocks.push({ kind: "table", rows, x, y });
        continue;
      }
      if (findDesc(node, "chart")) {
        blocks.push({ kind: "chart", x, y });
      }
    }
  }
}

function slideTitle(blocks: RawBlock[]): string {
  for (const role of ["title", "subtitle"] as const) {
    const block = blocks.find((entry) => entry.kind === "text" && entry.role === role);
    if (!block || block.kind !== "text") continue;
    const text = block.lines.map((line) => plainLine(line)).filter(Boolean).join(" ");
    if (text) return text;
  }
  return "";
}

function sortBlocks(blocks: RawBlock[]): RawBlock[] {
  return [...blocks].sort((a, b) => {
    const rowA = Math.round(a.y / 120000);
    const rowB = Math.round(b.y / 120000);
    if (rowA !== rowB) return rowA - rowB;
    return a.x - b.x;
  });
}

async function parseSlide(zip: JSZip, part: string, index: number): Promise<ParsedSlide> {
  const xml = await readText(zip, part);
  if (!xml) {
    return { index, hidden: false, title: "", blocks: [], notes: [], error: "This slide is missing from the file." };
  }
  const relsXml = await readText(zip, relsPath(part));
  const rels = relsXml ? parseRels(relsXml) : {};
  const root = findRoot(parseXml(xml), "sld");
  if (!root) {
    return { index, hidden: false, title: "", blocks: [], notes: [], error: "This slide could not be read." };
  }
  const hidden = attrLocal(root, "show", false) === "0";
  const cSld = findChild(root, "cSld");
  const blocks: RawBlock[] = [];
  const ctx: WalkCtx = { zip, slidePart: part, rels, allowImages: true };
  if (cSld) {
    const spTree = findChild(cSld, "spTree");
    if (spTree) await walkShapes(children(spTree), 0, 0, blocks, ctx);
    const bg = findChild(cSld, "bg");
    if (bg) {
      const ids: string[] = [];
      collectEmbeds(bg, ids);
      if (ids.length) {
        const image = await materializeImage(ids, { name: "Background", alt: "" }, ctx);
        blocks.push({ kind: "image", image, x: 0, y: Number.MAX_SAFE_INTEGER / 4 });
      }
    }
  }
  const notesRel = Object.values(rels).find((rel) => rel.type.endsWith("/notesSlide") && !rel.external);
  let notes: TextLine[] = [];
  if (notesRel) {
    const notesPart = resolvePath(part, notesRel.target);
    const notesXml = await readText(zip, notesPart);
    if (notesXml) {
      const notesRoot = findRoot(parseXml(notesXml), "notes");
      const notesTree = notesRoot ? findDesc(notesRoot, "spTree") : undefined;
      if (notesTree) {
        const noteBlocks: RawBlock[] = [];
        await walkShapes(children(notesTree), 0, 0, noteBlocks, { ...ctx, allowImages: false });
        notes = noteBlocks.flatMap((block) => (block.kind === "text" ? block.lines : []));
      }
    }
  }
  const ordered = sortBlocks(blocks);
  return { index, hidden, title: slideTitle(ordered), blocks: ordered, notes };
}

export async function parsePptx(buffer: ArrayBuffer): Promise<ParsedDeck> {
  const bytes = new Uint8Array(buffer);
  if (isOle(bytes)) {
    throw new DeckError("This is an older .ppt file. Save it as .pptx in PowerPoint, then try again.");
  }
  if (!isZip(bytes)) {
    throw new DeckError("That file is not a PowerPoint deck. Use a .pptx file.");
  }
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    throw new DeckError("Could not open that file. It may be damaged or password-protected.");
  }
  const presentation = await readText(zip, "ppt/presentation.xml");
  if (!presentation) {
    throw new DeckError("That file does not contain a PowerPoint presentation.");
  }
  const rels = parseRels((await readText(zip, relsPath("ppt/presentation.xml"))) ?? "");
  const root = findRoot(parseXml(presentation), "presentation");
  const list = root ? findDesc(root, "sldIdLst") : undefined;
  const ids = list ? children(list).filter((node) => localName(node) === "sldId") : [];
  if (!ids.length) throw new DeckError("This deck has no slides.");

  const slides: ParsedSlide[] = [];
  for (let i = 0; i < ids.length; i++) {
    const relId = attrLocal(ids[i], "id", true);
    const rel = relId ? rels[relId] : undefined;
    if (!rel || rel.external) {
      slides.push({
        index: i + 1,
        hidden: false,
        title: "",
        blocks: [],
        notes: [],
        error: "This slide could not be found in the file.",
      });
      continue;
    }
    try {
      slides.push(await parseSlide(zip, resolvePath("ppt/presentation.xml", rel.target), i + 1));
    } catch (error) {
      slides.push({
        index: i + 1,
        hidden: false,
        title: "",
        blocks: [],
        notes: [],
        error: error instanceof Error ? error.message : "Could not read this slide.",
      });
    }
  }
  return { slides };
}
