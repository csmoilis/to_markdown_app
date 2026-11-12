import { OCR_OFF_DETAIL } from "./types.ts";
import type { Deck, DeckSlide, ImageResult, TextLine, TextRun } from "./types.ts";
import { plainLine } from "./xml.ts";

function styleRun(run: TextRun): string {
  const lead = run.text.match(/^\s*/)?.[0] ?? "";
  const trail = run.text.match(/\s*$/)?.[0] ?? "";
  const core = run.text.trim();
  if (!core) return run.text;
  let wrapped = core;
  if (run.bold && run.italic) wrapped = `***${core}***`;
  else if (run.bold) wrapped = `**${core}**`;
  else if (run.italic) wrapped = `*${core}*`;
  return `${lead}${wrapped}${trail}`;
}

export function lineToMarkdown(line: TextLine): string {
  const text = line.runs.map(styleRun).join("").trim();
  if (!text) return "";
  const indent = "  ".repeat(line.level);
  if (line.bullet || line.level > 0) return `${indent}- ${text}`;
  return text;
}

function linesToMarkdown(lines: TextLine[]): string {
  return lines.map(lineToMarkdown).filter(Boolean).join("\n");
}

function escapeCell(value: string): string {
  return value.replace(/\|/g, "\\|").replace(/\n/g, " ").trim();
}

function tableToMarkdown(rows: string[][]): string {
  if (!rows.length) return "";
  const width = Math.max(...rows.map((row) => row.length));
  const padded = rows.map((row) => Array.from({ length: width }, (_, index) => escapeCell(row[index] ?? "")));
  const header = `| ${padded[0].join(" | ")} |`;
  const rule = `| ${padded[0].map(() => "---").join(" | ")} |`;
  const body = padded.slice(1).map((row) => `| ${row.join(" | ")} |`);
  return [header, rule, ...body].join("\n");
}

function imageToMarkdown(image: ImageResult, unit: "slide" | "page"): string {
  const label = image.alt ? `${image.name} — ${image.alt}` : image.name;
  if (image.status === "skipped" && image.detail === OCR_OFF_DETAIL) return "";
  if (image.repeatOf !== undefined) {
    // Same picture as an earlier slide/page: point there instead of repeating it.
    if (image.status !== "text") return "";
    return `### Image: ${label}\n\n_Same image as ${unit} ${image.repeatOf}; its text is shown there._`;
  }
  if (image.status === "text") {
    const quoted = image.text
      .split("\n")
      .map((line) => `> ${line}`.trimEnd())
      .join("\n");
    return `### Text from image: ${label}\n\n${quoted}`;
  }
  if (image.status === "pending") {
    return `### Image: ${label}\n\n_Still reading this image._`;
  }
  if (image.status === "empty") {
    return `### Image: ${label}\n\n_No readable text found in this image._`;
  }
  const detail = image.detail ?? "Text could not be read from this image.";
  return `### Image: ${label}\n\n_${detail}_`;
}

export function slideToMarkdown(slide: DeckSlide, unit: "slide" | "page" = "slide"): string {
  const label = unit === "page" ? "Page" : "Slide";
  const title = slide.title ? ` — ${slide.title.replace(/\s+/g, " ").trim()}` : "";
  const hidden = slide.hidden ? " (hidden)" : "";
  const parts = [`## ${label} ${slide.index}${title}${hidden}`, ""];
  let skippedTitle = false;
  for (const block of slide.blocks) {
    if (block.kind === "text") {
      const plain = block.lines.map((line) => plainLine(line)).filter(Boolean).join(" ");
      if (block.role === "title" && !skippedTitle && plain === slide.title) {
        skippedTitle = true;
        continue;
      }
      const markdown = linesToMarkdown(block.lines);
      if (markdown) parts.push(markdown, "");
      continue;
    }
    if (block.kind === "table") {
      const markdown = tableToMarkdown(block.rows);
      if (markdown) parts.push(markdown, "");
      continue;
    }
    if (block.kind === "image") {
      const markdown = imageToMarkdown(block.image, unit);
      if (markdown) parts.push(markdown, "");
      continue;
    }
    if (block.kind === "chart") {
      parts.push("### Chart", "", "_Chart values live in the workbook, not as slide text._", "");
      continue;
    }
    if (block.kind === "notes") {
      const markdown = linesToMarkdown(block.lines);
      if (markdown) parts.push("### Speaker notes", "", markdown, "");
      continue;
    }
    if (block.kind === "error") {
      parts.push(`_${block.message}_`, "");
    }
  }
  return `${parts.join("\n").replace(/\n{3,}/g, "\n\n").trim()}\n`;
}

/** Slides/pages still kept for export, i.e. not in `removed` (a set of `slide.index` values). */
export function keptSlides(deck: Deck, removed?: ReadonlySet<number>): DeckSlide[] {
  if (!removed || removed.size === 0) return deck.slides;
  return deck.slides.filter((slide) => !removed.has(slide.index));
}

/**
 * Whole-file markdown. Sections whose `slide.index` is in `removed` are left out, so the preview,
 * Copy markdown and Download .md only contain what the user kept.
 */
export function deckToMarkdown(deck: Deck, removed?: ReadonlySet<number>): string {
  const title = deck.fileName.replace(/\.(pptx|pptm|ppsx|potx|ppt|pdf)$/i, "") || "Document";
  const body = keptSlides(deck, removed)
    .map((slide) => slideToMarkdown(slide, deck.unit))
    .join("\n");
  return `# ${title}\n\n${body}`;
}

export function downloadName(fileName: string): string {
  const base = fileName.replace(/\.(pptx|pptm|ppsx|potx|ppt|pdf)$/i, "") || "document";
  return `${base}.md`;
}
