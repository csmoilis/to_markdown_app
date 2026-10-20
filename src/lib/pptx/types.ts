export type TextRun = {
  text: string;
  bold: boolean;
  italic: boolean;
};

export type TextLine = {
  level: number;
  bullet: boolean;
  runs: TextRun[];
};

export type ImagePayload = {
  name: string;
  alt: string;
  mime: string;
  bytes: Uint8Array | null;
  skipReason?: string;
};

export type RawBlock =
  | { kind: "text"; role: "title" | "subtitle" | "body"; lines: TextLine[]; x: number; y: number }
  | { kind: "table"; rows: string[][]; x: number; y: number }
  | { kind: "image"; image: ImagePayload; x: number; y: number }
  | { kind: "chart"; x: number; y: number };

export type ParsedSlide = {
  index: number;
  hidden: boolean;
  title: string;
  blocks: RawBlock[];
  notes: TextLine[];
  error?: string;
};

export type ParsedDeck = {
  slides: ParsedSlide[];
};

export type ImageResult = {
  name: string;
  alt: string;
  status: "text" | "empty" | "skipped" | "error" | "pending";
  text: string;
  detail?: string;
  /** Set when the same picture already appeared on an earlier slide or page (its number). */
  repeatOf?: number;
};

export type DeckBlock =
  | { kind: "text"; role: "title" | "subtitle" | "body"; lines: TextLine[] }
  | { kind: "table"; rows: string[][] }
  | { kind: "image"; image: ImageResult }
  | { kind: "chart" }
  | { kind: "notes"; lines: TextLine[] }
  | { kind: "error"; message: string };

export type DeckSlide = {
  index: number;
  hidden: boolean;
  title: string;
  blocks: DeckBlock[];
};

export type Deck = {
  fileName: string;
  unit: "slide" | "page";
  slides: DeckSlide[];
};

/** Detail used for pictures when "Read text inside images" is off; these are left out of the markdown. */
export const OCR_OFF_DETAIL = "Reading text in images is turned off.";
