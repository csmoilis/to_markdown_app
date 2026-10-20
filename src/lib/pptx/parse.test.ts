import assert from "node:assert/strict";
import test from "node:test";
import { buildPptx } from "./build.ts";
import { deckToMarkdown } from "./markdown.ts";
import { parsePptx } from "./parse.ts";
import type { Deck } from "./types.ts";

const PNG = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

test("extracts text, tables, notes, groups, fallback text, and images", async () => {
  const buffer = await buildPptx([
    {
      title: "Hello & Co",
      titleBold: true,
      subtitle: "A quiet subtitle",
      bullets: ["First point", "Second point"],
      groupText: "Inside group",
      fallbackText: "FALLBACK ONLY",
      choiceText: "CHOICE ONLY",
      notes: "Say this out loud",
      chart: true,
    },
    {
      title: "Numbers",
      table: [
        ["Region", "Total"],
        ["North", "18420"],
      ],
      image: { bytes: PNG, ext: "png" },
      hidden: true,
    },
  ]);
  const deck = await parsePptx(buffer);
  assert.equal(deck.slides.length, 2);

  const first = deck.slides[0];
  assert.equal(first?.title, "Hello & Co");
  assert.equal(first?.hidden, false);
  const title = first?.blocks.find((block) => block.kind === "text" && block.role === "title");
  assert.ok(title && title.kind === "text");
  assert.equal(title.lines[0]?.runs[0]?.bold, true);
  const body = first.blocks.filter((block) => block.kind === "text");
  const texts = body.flatMap((block) => (block.kind === "text" ? block.lines.map((line) => line.runs.map((run) => run.text).join("")) : []));
  assert.ok(texts.some((text) => text.includes("First point") && text.includes("")));
  assert.ok(texts.includes("Inside group"));
  assert.ok(texts.includes("FALLBACK ONLY"));
  assert.equal(texts.includes("CHOICE ONLY"), false);
  assert.equal(
    first.notes.map((line) => line.runs.map((run) => run.text).join("")).join(" "),
    "Say this out loud",
  );
  assert.ok(first.blocks.some((block) => block.kind === "chart"));

  const second = deck.slides[1];
  assert.equal(second?.hidden, true);
  const table = second?.blocks.find((block) => block.kind === "table");
  assert.ok(table && table.kind === "table");
  assert.deepEqual(table.rows, [
    ["Region", "Total"],
    ["North", "18420"],
  ]);
  const image = second.blocks.find((block) => block.kind === "image");
  assert.ok(image && image.kind === "image");
  assert.equal(image.image.mime, "image/png");
  assert.equal(image.image.bytes?.length, PNG.length);
  assert.equal(image.image.alt, "Sample banner");
});

test("rejects legacy ppt and non-decks", async () => {
  const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0]).buffer;
  await assert.rejects(parsePptx(ole), /older \.ppt/);
  const text = new TextEncoder().encode("hello").buffer;
  await assert.rejects(parsePptx(text), /not a PowerPoint/);
});

test("renders a markdown document", () => {
  const deck: Deck = {
    fileName: "Q1.pptx",
    unit: "slide",
    slides: [
      {
        index: 1,
        hidden: false,
        title: "Plan",
        blocks: [
          { kind: "text", role: "title", lines: [{ level: 0, bullet: false, runs: [{ text: "Plan", bold: true, italic: false }] }] },
          { kind: "text", role: "body", lines: [{ level: 0, bullet: true, runs: [{ text: "Ship it", bold: false, italic: false }] }] },
          {
            kind: "image",
            image: { name: "Banner", alt: "Sample banner", status: "text", text: "NORTHWIND 2026\nInvoice total 18420" },
          },
          { kind: "notes", lines: [{ level: 0, bullet: false, runs: [{ text: "Speak slowly", bold: false, italic: false }] }] },
        ],
      },
    ],
  };
  const markdown = deckToMarkdown(deck);
  assert.match(markdown, /^# Q1\n/);
  assert.match(markdown, /## Slide 1 — Plan/);
  assert.match(markdown, /- Ship it/);
  assert.match(markdown, /### Text from image: Banner — Sample banner/);
  assert.match(markdown, /> NORTHWIND 2026/);
  assert.match(markdown, /### Speaker notes/);
  assert.doesNotMatch(markdown, /\*\*Plan\*\*/);
});
