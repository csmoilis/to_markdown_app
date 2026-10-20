import { Button } from "@/components/ui/button";
import { extractDeck } from "@/lib/pptx/extract";
import { looksLikePdf } from "@/lib/pptx/pdf";
import { deckToMarkdown, downloadName, slideToMarkdown } from "@/lib/pptx/markdown";
import { buildSampleFile } from "@/lib/pptx/sample";
import type { Deck, DeckBlock, DeckSlide, TextLine } from "@/lib/pptx/types";
import { plainLine } from "@/lib/pptx/xml";
import { cn } from "@/lib/utils";
import { Copy, Download, FileUp } from "lucide-react";
import { useEffect, useRef, useState } from "react";

const MAX_BYTES = 30 * 1024 * 1024;
const INDENT = ["pl-0", "pl-4", "pl-8", "pl-12", "pl-16"] as const;

type View = "reading" | "markdown";
type Scope = "slide" | "deck";

function LineView({ line }: { line: TextLine }) {
  const marker = line.bullet || line.level > 0;
  return (
    <p className={cn("text-base leading-normal", INDENT[Math.min(line.level, INDENT.length - 1)])}>
      {marker ? <span className="mr-2 text-muted-foreground">–</span> : null}
      {line.runs.map((run, index) => {
        if (run.bold && run.italic) {
          return (
            <strong key={index} className="italic">
              {run.text}
            </strong>
          );
        }
        if (run.bold) return <strong key={index}>{run.text}</strong>;
        if (run.italic) return <em key={index}>{run.text}</em>;
        return <span key={index}>{run.text}</span>;
      })}
    </p>
  );
}

function Lines({ lines }: { lines: TextLine[] }) {
  return (
    <div className="space-y-2">
      {lines.map((line, index) => (
        <LineView key={index} line={line} />
      ))}
    </div>
  );
}

function BlockView({ block, unit }: { block: DeckBlock; unit: "slide" | "page" }) {
  if (block.kind === "text") {
    const plain = block.lines.map((line) => plainLine(line)).join(" ");
    if (!plain) return null;
    return <Lines lines={block.lines} />;
  }
  if (block.kind === "table") {
    const [header, ...body] = block.rows;
    if (!header) return null;
    return (
      <div className="overflow-x-auto">
        <table className="w-full border-collapse text-left text-sm">
          <thead>
            <tr>
              {header.map((cell, index) => (
                <th key={index} className="border-b border-border px-3 py-2 font-medium">
                  {cell}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {(body.length ? body : []).map((row, rowIndex) => (
              <tr key={rowIndex}>
                {header.map((_, cellIndex) => (
                  <td key={cellIndex} className="border-b border-border px-3 py-2 align-top">
                    {row[cellIndex] ?? ""}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
  if (block.kind === "image") {
    const label = block.image.alt ? `${block.image.name} — ${block.image.alt}` : block.image.name;
    if (block.image.repeatOf !== undefined) {
      if (block.image.status !== "text") return null;
      return (
        <section className="border-t border-border pt-4">
          <h3 className="text-sm font-medium text-muted-foreground">Image · {label}</h3>
          <p className="mt-2 text-sm text-muted-foreground">
            Same image as {unit} {block.image.repeatOf}. Its text is shown there.
          </p>
        </section>
      );
    }
    return (
      <section className="border-t border-border pt-4">
        <h3 className="text-sm font-medium text-muted-foreground">
          {block.image.status === "text" ? "Text from image" : "Image"} · {label}
        </h3>
        {block.image.status === "pending" ? (
          <p className="mt-2 text-sm text-muted-foreground">Still reading this image.</p>
        ) : block.image.status === "text" ? (
          <p className="mt-2 whitespace-pre-wrap text-base leading-normal">{block.image.text}</p>
        ) : (
          <p className="mt-2 text-sm text-muted-foreground">
            {block.image.status === "empty"
              ? "No readable text found in this image."
              : (block.image.detail ?? "Text could not be read from this image.")}
          </p>
        )}
      </section>
    );
  }
  if (block.kind === "chart") {
    return (
      <section className="border-t border-border pt-4">
        <h3 className="text-sm font-medium text-muted-foreground">Chart</h3>
        <p className="mt-2 text-sm text-muted-foreground">Chart values live in the workbook, not as slide text.</p>
      </section>
    );
  }
  if (block.kind === "notes") {
    return (
      <section className="border-t border-border pt-4">
        <h3 className="text-sm font-medium text-muted-foreground">Speaker notes</h3>
        <div className="mt-2">
          <Lines lines={block.lines} />
        </div>
      </section>
    );
  }
  return <p className="text-sm text-destructive">{block.message}</p>;
}

function Reading({ slide, unit }: { slide: DeckSlide; unit: "slide" | "page" }) {
  const noun = unit === "page" ? "Page" : "Slide";
  const visible = slide.blocks.filter((block) => {
    if (block.kind !== "text" || block.role !== "title") return true;
    const plain = block.lines.map((line) => plainLine(line)).filter(Boolean).join(" ");
    return plain !== slide.title;
  });
  const empty = visible.every((block) => block.kind === "text" && !block.lines.length);
  return (
    <article>
      <p className="text-sm font-medium text-muted-foreground">
        {noun} {slide.index}
        {slide.hidden ? " · hidden" : ""}
      </p>
      <h2 className="mt-1 font-display text-3xl leading-tight tracking-tight">
        {slide.title || (unit === "page" ? "Untitled page" : "Untitled slide")}
      </h2>
      <div className="mt-6 space-y-4">
        {empty || visible.length === 0 ? (
          <p className="text-muted-foreground">This {unit} has no text.</p>
        ) : (
          visible.map((block, index) => <BlockView key={index} block={block} unit={unit} />)
        )}
      </div>
    </article>
  );
}

export function FolioApp() {
  const inputRef = useRef<HTMLInputElement>(null);
  const resultRef = useRef<HTMLElement>(null);
  const runRef = useRef(0);
  const [over, setOver] = useState(false);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("Ready");
  const [error, setError] = useState<string | null>(null);
  const [deck, setDeck] = useState<Deck | null>(null);
  const [done, setDone] = useState(0);
  const [total, setTotal] = useState(0);
  const [active, setActive] = useState(0);
  const [view, setView] = useState<View>("reading");
  const [scope, setScope] = useState<Scope>("deck");
  const [copied, setCopied] = useState(false);
  const [readImages, setReadImages] = useState(true);

  async function runBuffer(buffer: ArrayBuffer, fileName: string, unit: "slide" | "page") {
    const run = ++runRef.current;
    setBusy(true);
    setError(null);
    setDeck(null);
    setDone(0);
    setTotal(0);
    setActive(0);
    setView("reading");
    setStatus("Opening deck");
    try {
      const next = await extractDeck(buffer, fileName, {
        onStatus: (message) => {
          if (runRef.current === run) setStatus(message);
        },
        onProgress: (completed, pageTotal) => {
          if (runRef.current !== run) return;
          setDone(completed);
          setTotal(pageTotal);
        },
        onSlide: (slide) => {
          if (runRef.current !== run) return;
          setDeck((current) => {
            const slides = current?.slides.slice() ?? [];
            const at = slides.findIndex((item) => item.index === slide.index);
            if (at === -1) slides.push(slide);
            else slides[at] = slide;
            slides.sort((a, b) => a.index - b.index);
            return { fileName, unit, slides };
          });
        },
      }, { readImages });
      if (runRef.current !== run) return;
      setDeck(next);
      setStatus("Done");
    } catch (caught) {
      if (runRef.current !== run) return;
      setError(caught instanceof Error ? caught.message : "Could not read that file.");
      setStatus("Could not read the file");
    } finally {
      if (runRef.current === run) setBusy(false);
    }
  }

  async function runFile(file: File) {
    if (file.size > MAX_BYTES) {
      setError("That file is over 30 MB. Try a smaller file.");
      return;
    }
    const buffer = await file.arrayBuffer();
    await runBuffer(buffer, file.name, looksLikePdf(buffer, file.name) ? "page" : "slide");
  }

  async function runSample() {
    setStatus("Preparing sample deck");
    setError(null);
    try {
      const sample = await buildSampleFile();
      await runBuffer(sample.buffer, sample.name, "slide");
    } catch (caught) {
      setBusy(false);
      setError(caught instanceof Error ? caught.message : "Could not build the sample deck.");
    }
  }

  const slide = deck?.slides[Math.min(active, Math.max(0, (deck?.slides.length ?? 1) - 1))];
  const markdown = deck
    ? scope === "slide" && slide
      ? slideToMarkdown(slide, deck.unit)
      : deckToMarkdown(deck)
    : "";
  const noun = deck?.unit === "page" ? "page" : "slide";
  const nounPlural = deck?.unit === "page" ? "pages" : "slides";
  const Noun = deck?.unit === "page" ? "Page" : "Slide";
  const progress = total > 0 ? Math.round((done / total) * 100) : null;
  const imageCount = deck
    ? deck.slides.reduce((count, item) => count + item.blocks.filter((block) => block.kind === "image").length, 0)
    : 0;
  const pendingImages = deck
    ? deck.slides.reduce(
        (count, item) => count + item.blocks.filter((block) => block.kind === "image" && block.image.status === "pending").length,
        0,
      )
    : 0;
  const finished = Boolean(deck) && !busy && !error && pendingImages === 0;
  const statusText = finished
    ? `Done. ${deck?.slides.length ?? 0} ${(deck?.slides.length ?? 0) === 1 ? noun : nounPlural} are ready. Use Copy markdown or Download .md.`
    : !busy && deck && pendingImages > 0
      ? `Stopped early. ${pendingImages} ${pendingImages === 1 ? "image was" : "images were"} not read. The text already found is below.`
      : status;

  useEffect(() => {
    if (!finished) return;
    resultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  }, [finished]);

  function copyMarkdown() {
    if (!markdown) return;
    void navigator.clipboard.writeText(markdown).then(() => {
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1600);
    });
  }

  function downloadMarkdown() {
    if (!deck) return;
    const blob = new Blob([deckToMarkdown(deck)], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = downloadName(deck.fileName);
    link.click();
    URL.revokeObjectURL(url);
  }

  return (
    <div className="min-h-dvh bg-background text-foreground">
      <header className="border-b border-border">
        <div className="mx-auto flex max-w-6xl flex-col gap-3 px-5 py-8 md:flex-row md:items-end md:justify-between">
          <div>
            <h1 className="font-display text-4xl leading-tight tracking-tight">Folio</h1>
            <p className="mt-2 max-w-xl text-base text-muted-foreground">
              Turn a PowerPoint or PDF into markdown. Each slide or page keeps its text, tables, and notes. Words
              inside images are read too.
            </p>
          </div>
          <p className="max-w-xs text-sm text-muted-foreground md:text-right">The file stays in this browser.</p>
        </div>
      </header>
      <main className="mx-auto max-w-6xl px-5 py-8 pb-16">
        <div
          className={cn("rounded-4xl border bg-card p-4", over ? "border-foreground" : "border-border")}
          onDragOver={(event) => {
            event.preventDefault();
            setOver(true);
          }}
          onDragLeave={() => setOver(false)}
          onDrop={(event) => {
            event.preventDefault();
            setOver(false);
            const file = event.dataTransfer.files[0];
            if (file) void runFile(file);
          }}
        >
          <label className="flex min-h-44 cursor-pointer flex-col items-center justify-center rounded-2xl border border-dashed border-border px-6 py-8 text-center">
            <FileUp className="size-6" aria-hidden="true" />
            <span className="mt-4 text-lg font-medium">Drop a .pptx or .pdf here</span>
            <span className="mt-1 text-sm text-muted-foreground">or choose a file, up to 30 MB</span>
            <input
              ref={inputRef}
              className="sr-only"
              type="file"
              accept=".pptx,.pptm,.ppsx,.potx,.pdf,application/pdf,application/vnd.openxmlformats-officedocument.presentationml.presentation"
              onChange={(event) => {
                const file = event.target.files?.[0];
                event.target.value = "";
                if (file) void runFile(file);
              }}
            />
          </label>
          <div className="mt-4 flex flex-col gap-3 sm:flex-row">
            <Button className="sm:min-w-36" disabled={busy} onClick={() => inputRef.current?.click()}>
              Choose file
            </Button>
            <Button className="sm:min-w-36" variant="secondary" disabled={busy} onClick={() => void runSample()}>
              Try sample deck
            </Button>
            <label className="flex items-center gap-2 text-sm text-muted-foreground sm:ml-auto">
              <input
                type="checkbox"
                className="size-4 accent-current"
                checked={readImages}
                disabled={busy}
                onChange={(event) => setReadImages(event.target.checked)}
              />
              Read text inside images (slower)
            </label>
          </div>
        </div>

        {busy || error || deck ? (
          <div className="mt-6">
            <p className="text-sm text-muted-foreground" aria-live="polite">
              {statusText}
            </p>
            {busy || finished ? (
              <div
                className="mt-3 h-1 overflow-hidden rounded-full bg-secondary"
                role="progressbar"
                aria-valuemin={0}
                aria-valuemax={total || 1}
                aria-valuenow={finished ? total || 1 : done}
                aria-label={statusText}
              >
                {finished ? (
                  <div className="h-full w-full bg-primary" />
                ) : progress === null || done === 0 ? (
                  <div className="folio-indeterminate h-full bg-primary" />
                ) : (
                  <div className="h-full bg-primary" style={{ width: `${progress}%` }} />
                )}
              </div>
            ) : null}
            {error ? (
              <p className="mt-3 text-sm text-destructive" role="alert">
                {error}
              </p>
            ) : null}
          </div>
        ) : null}

        {deck && slide ? (
          <section ref={resultRef} className="mt-8">
            <div className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
              <div>
                <h2 className="font-display text-2xl tracking-tight">{deck.fileName}</h2>
                <p className="mt-1 text-sm text-muted-foreground">
                  {deck.slides.length} {deck.slides.length === 1 ? noun : nounPlural}
                  {imageCount ? ` · ${imageCount} ${imageCount === 1 ? "image" : "images"}` : ""}
                </p>
              </div>
              <div className="flex flex-col gap-3 sm:flex-row">
                <Button variant="secondary" onClick={copyMarkdown}>
                  <Copy className="size-4" aria-hidden="true" />
                  {copied ? "Copied" : "Copy markdown"}
                </Button>
                <Button onClick={downloadMarkdown} disabled={busy}>
                  <Download className="size-4" aria-hidden="true" />
                  Download .md
                </Button>
              </div>
            </div>

            <div className="mt-6 grid gap-6 md:grid-cols-[13rem_minmax(0,1fr)]">
              <nav aria-label={Noun === "Page" ? "Pages" : "Slides"} className="flex gap-2 overflow-x-auto md:flex-col md:overflow-visible">
                {deck.slides.map((item, index) => {
                  const selected = index === active;
                  return (
                    <button
                      key={item.index}
                      type="button"
                      onClick={() => setActive(index)}
                      className={cn(
                        "min-h-11 w-44 shrink-0 rounded-lg border px-3 py-2 text-left md:w-full",
                        selected
                          ? "border-primary bg-primary text-primary-foreground"
                          : "border-border bg-card text-foreground",
                      )}
                    >
                      <span className="block text-xs font-medium tabular-nums">{Noun} {item.index}</span>
                      <span className="block truncate text-sm">{item.title || "Untitled"}</span>
                    </button>
                  );
                })}
              </nav>

              <div className="min-w-0 rounded-4xl border border-border bg-card p-6 md:p-8">
                <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                  <div className="flex gap-2">
                    <Button variant={view === "reading" ? "primary" : "ghost"} onClick={() => setView("reading")}>
                      Reading
                    </Button>
                    <Button variant={view === "markdown" ? "primary" : "ghost"} onClick={() => setView("markdown")}>
                      Markdown
                    </Button>
                  </div>
                  {view === "markdown" ? (
                    <div className="flex gap-2">
                      <Button variant={scope === "slide" ? "secondary" : "ghost"} onClick={() => setScope("slide")}>
                        This {noun}
                      </Button>
                      <Button variant={scope === "deck" ? "secondary" : "ghost"} onClick={() => setScope("deck")}>
                        {deck.unit === "page" ? "Entire file" : "Entire deck"}
                      </Button>
                    </div>
                  ) : null}
                </div>
                <div className="mt-6">
                  {view === "reading" ? (
                    <Reading slide={slide} unit={deck.unit} />
                  ) : (
                    <pre className="overflow-x-auto whitespace-pre-wrap break-words font-sans text-sm leading-normal">
                      {markdown}
                    </pre>
                  )}
                </div>
              </div>
            </div>
          </section>
        ) : null}
      </main>
    </div>
  );
}
