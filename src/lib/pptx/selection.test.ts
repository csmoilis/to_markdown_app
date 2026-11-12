import assert from "node:assert/strict";
import test from "node:test";
import { deckToMarkdown, keptSlides } from "./markdown.ts";
import { emptySelection, removeSection, restoreAll, restoreSection, toggleSection, undo } from "./selection.ts";
import type { Deck } from "./types.ts";

function line(text: string) {
  return { level: 0, bullet: false, runs: [{ text, bold: false, italic: false }] };
}

const deck: Deck = {
  fileName: "Q1.pptx",
  unit: "slide",
  slides: [1, 2, 3].map((index) => ({
    index,
    title: `Title ${index}`,
    hidden: false,
    blocks: [{ kind: "text" as const, role: "body" as const, lines: [line(`Body ${index}`)] }],
  })),
};

test("removed sections are left out of the markdown", () => {
  let selection = removeSection(emptySelection, 2);
  const markdown = deckToMarkdown(deck, selection.removed);
  assert.match(markdown, /## Slide 1 — Title 1/);
  assert.doesNotMatch(markdown, /Slide 2|Body 2/);
  assert.match(markdown, /## Slide 3 — Title 3/);
  assert.deepEqual(
    keptSlides(deck, selection.removed).map((slide) => slide.index),
    [1, 3],
  );
  selection = restoreSection(selection, 2);
  assert.equal(deckToMarkdown(deck, selection.removed), deckToMarkdown(deck));
});

test("without a removed set every section is kept", () => {
  assert.equal(keptSlides(deck).length, 3);
  assert.equal(deckToMarkdown(deck, new Set()), deckToMarkdown(deck));
});

test("toggle, restore all, and undo", () => {
  let selection = toggleSection(emptySelection, 1);
  selection = toggleSection(selection, 3);
  assert.deepEqual([...selection.removed].sort(), [1, 3]);
  selection = toggleSection(selection, 1);
  assert.deepEqual([...selection.removed], [3]);
  selection = restoreAll(selection);
  assert.equal(selection.removed.size, 0);
  selection = undo(selection);
  assert.deepEqual([...selection.removed], [3]);
  selection = undo(selection);
  assert.deepEqual([...selection.removed].sort(), [1, 3]);
  selection = undo(selection);
  selection = undo(selection);
  assert.equal(selection.removed.size, 0);
  assert.equal(undo(selection), selection);
});

test("no-op changes do not add undo steps", () => {
  const once = removeSection(emptySelection, 2);
  assert.equal(removeSection(once, 2), once);
  assert.equal(restoreAll(emptySelection), emptySelection);
});
