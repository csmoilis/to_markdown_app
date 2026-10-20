import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";
import { GlobalWorkerOptions } from "pdfjs-dist/legacy/build/pdf.mjs";
import { parsePdf } from "./pdf.ts";

const require = createRequire(import.meta.url);
GlobalWorkerOptions.workerSrc = pathToFileURL(require.resolve("pdfjs-dist/legacy/build/pdf.worker.mjs")).href;

function textPdf(): Uint8Array {
  const objects = [
    "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
    "2 0 obj\n<< /Type /Pages /Count 2 /Kids [3 0 R 6 0 R] >>\nendobj\n",
    "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 240] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n",
    "",
    "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n",
    "6 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 240] /Contents 7 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n",
    "",
  ];
  const page1 = "BT /F1 24 Tf 40 180 Td (Hello PDF) Tj ET\nBT /F1 14 Tf 40 140 Td (Second line of page one) Tj ET\n";
  const page2 = "BT /F1 18 Tf 40 160 Td (Page two title) Tj ET\n";
  objects[3] = `4 0 obj\n<< /Length ${page1.length} >>\nstream\n${page1}endstream\nendobj\n`;
  objects[6] = `7 0 obj\n<< /Length ${page2.length} >>\nstream\n${page2}endstream\nendobj\n`;
  return wrap(objects);
}

function imagePdf(): Uint8Array {
  const width = 80;
  const height = 40;
  const pixels = Buffer.alloc(width * height * 3, 250);
  const objects = [
    "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
    "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n",
    "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 220] /Contents 4 0 R /Resources << /XObject << /Im1 5 0 R >> /Font << /F1 6 0 R >> >> >>\nendobj\n",
    "",
    "",
    "6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n",
  ];
  const content = "BT /F1 18 Tf 20 180 Td (Caption above) Tj ET\nq 200 0 0 80 20 40 cm /Im1 Do Q\n";
  objects[3] = `4 0 obj\n<< /Length ${content.length} >>\nstream\n${content}endstream\nendobj\n`;
  const header = `5 0 obj\n<< /Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length ${pixels.length} >>\nstream\n`;
  return wrap(objects, {
    4: Buffer.concat([Buffer.from(header), pixels, Buffer.from("endstream\nendobj\n")]),
  });
}

function asBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function wrap(objects: string[], binary: Record<number, Buffer> = {}): Uint8Array {
  const chunks: Buffer[] = [Buffer.from("%PDF-1.4\n")];
  const offsets = [0];
  let length = chunks[0].length;
  objects.forEach((source, index) => {
    offsets.push(length);
    const body = binary[index] ?? Buffer.from(source);
    chunks.push(body);
    length += body.length;
  });
  const startxref = length;
  let xref = `xref\n0 ${offsets.length}\n0000000000 65535 f \n`;
  for (let index = 1; index < offsets.length; index++) xref += `${String(offsets[index]).padStart(10, "0")} 00000 n \n`;
  chunks.push(Buffer.from(`${xref}trailer\n<< /Size ${offsets.length} /Root 1 0 R >>\nstartxref\n${startxref}\n%%EOF\n`));
  return new Uint8Array(Buffer.concat(chunks));
}

test("reports each page as it is read", async () => {
  const started: number[] = [];
  const finished: string[] = [];
  const deck = await parsePdf(asBuffer(textPdf()), {
    onStart(index, total) {
      assert.equal(total, 2);
      started.push(index);
    },
    onPage(slide) {
      finished.push(slide.title);
    },
  });
  assert.deepEqual(started, [1, 2]);
  assert.deepEqual(finished, ["Hello PDF", "Page two title"]);
  assert.equal(deck.slides.length, 2);
});

test("extracts text from each PDF page", async () => {
  const deck = await parsePdf(asBuffer(textPdf()));
  assert.equal(deck.slides.length, 2);
  assert.equal(deck.slides[0]?.title, "Hello PDF");
  const body = deck.slides[0]?.blocks.find((block) => block.kind === "text" && block.role === "body");
  assert.ok(body && body.kind === "text");
  assert.match(body.lines.map((line) => line.runs.map((run) => run.text).join("")).join(" "), /Second line of page one/);
  assert.equal(deck.slides[1]?.title, "Page two title");
});

test("keeps pictures on a PDF page for text reading", async () => {
  const deck = await parsePdf(asBuffer(imagePdf()));
  assert.equal(deck.slides[0]?.title, "Caption above");
  const image = deck.slides[0]?.blocks.find((block) => block.kind === "image");
  assert.ok(image && image.kind === "image");
  assert.equal(image.image.name, "Image 1");
});

test("rejects a file that is not a PDF", async () => {
  await assert.rejects(parsePdf(new TextEncoder().encode("hello").buffer), /not a PDF/);
});
