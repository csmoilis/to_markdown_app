# to_markdown_app (Folio)

Turn a PowerPoint (`.pptx`) or PDF into Markdown, one section per slide or page. Slide titles, body text, bullets, tables and speaker notes come from the file's own text. Text inside pictures (screenshots, diagrams, scanned exhibits) is read with OCR and added to the markdown under each slide or page.

Everything runs on your computer, in your browser. Your file is never uploaded, and there are no paid APIs or keys.

## What you need

- [Node.js](https://nodejs.org/) 22.12 or newer (it includes npm). Works on macOS, Windows and Linux.
- An internet connection the **first time** you read text from images: the OCR engine ([tesseract.js](https://github.com/naptha/tesseract.js)) downloads its English model (about 10 MB) once, and the browser caches it after that.

You don't need anything else: no Python, and no system Tesseract install. OCR runs inside the browser through WebAssembly.

## Run it

```bash
git clone https://github.com/csmoilis/to_markdown_app.git
cd to_markdown_app
npm install
npm run dev
```

Then open <http://localhost:8080>, drop a `.pptx` or `.pdf` file on the page (30 MB max), and use **Copy markdown** or **Download .md** once it finishes.

- **Read text inside images**: on by default. Turn it off to get only the slide or page text, which takes a few seconds even for long decks.
- If `npm run dev` prints `spawn vite ENOENT`, run `npm install` again in this folder.

Production build (optional): `npm run build && npm run preview`, then open <http://127.0.0.1:8081>.

## Supported files

| Format | Supported |
| --- | --- |
| `.pptx`, `.pptm`, `.ppsx`, `.potx` | Yes |
| `.pdf` | Yes. Each page becomes a section. |
| `.ppt` (old binary format) | No. Open it in PowerPoint and save it as `.pptx` first. |

The app reads text from PNG, JPEG, GIF, BMP, WebP and TIFF pictures. For SVG it uses the text stored in the file. EMF and WMF drawings are skipped.

## Why it used to be slow and then stop, and what changed

On a 137-page lecture PDF (442 image placements, 148 distinct images), the old version stopped on **page 3** and never moved on. With the fixes below the same file finishes in about **12 seconds** on an 8-core machine and about **29 seconds** limited to 2 cores.

| Version | Time on the 137-page PDF (8 cores) |
| --- | --- |
| Original | Never finishes. Stuck on "Reading page 3 of 137" (stopped after 4 minutes) |
| Hang fixed, nothing else changed | 45.6 s |
| Current | 11.7 s (29 s on 2 cores) |

1. **The freeze (main bug).** pdf.js stores images that repeat across pages, such as logos and header strips, in a shared store (`page.commonObjs`, ids that start with `g_`) and not in the page's own store. The app only checked the page store, so its callback never ran and the run waited forever on the first repeated image. The app now checks the right store, and every image lookup gives up after 15 s, so nothing can block the run again (`src/lib/pptx/pdf.ts`).
2. **Repeated images are read once.** Each picture gets a content fingerprint. A logo or banner that shows up on 100 pages is OCR'd once. Later copies get "Same image as page N" in the markdown, or nothing if the picture had no text. On the test PDF this cut 442 OCR jobs to 148.
3. **Several OCR workers in parallel.** A small pool of Tesseract workers (half your CPU cores, between 1 and 4) loads the model once and reads images side by side, in page order. The workers are closed when the file is done to free memory.
4. **One bad image can't stop the run.** Each image gets 60 s. If one takes longer, its worker is replaced, that image is marked as skipped, and the rest carry on.
5. **Cleaner OCR output.** Lines Tesseract isn't confident about (decorative photos and strips that came out as junk like `YY a!`) are dropped. Very small images are scaled up before OCR, and very large ones are scaled down to 1800 px.
6. **Less memory.** pdf.js's decoded copy of each page is released once its images are copied out.
7. **Optional OCR toggle** in the UI, as described above.

The main code is in `src/lib/pptx/`: `pdf.ts` (PDF pages), `parse.ts` (PPTX slides), `ocr.ts` (OCR worker pool), `extract.ts` (pipeline, dedupe and parallelism), `markdown.ts` (output). The UI is `src/components/folio-app.tsx`.

## Limits

- OCR quality depends on the image. Clean screenshots and text images read well. Diagrams, stylised fonts and photos can come out partial or out of order.
- PDF body text comes from the PDF's own text layer, line by line. A multi-column slide can come out with columns interleaved. A scanned PDF (pages that are just pictures) is read only through OCR of those pictures.
- The first OCR run needs internet to download the English model.
