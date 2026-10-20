import { XMLParser } from "fast-xml-parser";

export type XmlNode = Record<string, unknown>;

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: false,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
});

export function parseXml(xml: string): XmlNode[] {
  const tree = parser.parse(xml);
  return Array.isArray(tree) ? (tree as XmlNode[]) : [];
}

export function tagName(node: XmlNode): string | null {
  for (const key of Object.keys(node)) {
    if (key !== ":@" ) return key;
  }
  return null;
}

export function localName(node: XmlNode): string | null {
  const name = tagName(node);
  if (!name) return null;
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

export function children(node: XmlNode): XmlNode[] {
  const name = tagName(node);
  if (!name || name === "#text") return [];
  const value = node[name];
  if (!Array.isArray(value)) return [];
  return value.filter((entry): entry is XmlNode => !!entry && typeof entry === "object");
}

export function attrs(node: XmlNode): Record<string, string> {
  const raw = node[":@"];
  if (!raw || typeof raw !== "object") return {};
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    out[key] = value == null ? "" : String(value);
  }
  return out;
}

export function attrLocal(node: XmlNode | null | undefined, local: string, namespaced: boolean): string | undefined {
  if (!node) return undefined;
  for (const [key, value] of Object.entries(attrs(node))) {
    const raw = key.startsWith("@_") ? key.slice(2) : key;
    const colon = raw.indexOf(":");
    const hasNs = colon !== -1;
    if (hasNs !== namespaced) continue;
    const name = hasNs ? raw.slice(colon + 1) : raw;
    if (name === local) return value;
  }
  return undefined;
}

export function findChild(node: XmlNode, name: string): XmlNode | undefined {
  return children(node).find((child) => localName(child) === name);
}

export function findDesc(node: XmlNode, name: string): XmlNode | undefined {
  for (const child of children(node)) {
    if (localName(child) === "#text") continue;
    if (localName(child) === name) return child;
    const nested = findDesc(child, name);
    if (nested) return nested;
  }
  return undefined;
}

export function findRoot(tree: XmlNode[], name: string): XmlNode | null {
  return tree.find((node) => localName(node) === name) ?? null;
}

export function textContent(node: XmlNode): string {
  if (tagName(node) === "#text") return String(node["#text"] ?? "");
  return children(node).map(textContent).join("");
}

export function plainLine(line: { runs: { text: string }[] }): string {
  return line.runs
    .map((run) => run.text)
    .join("")
    .replace(/\s+/g, " ")
    .trim();
}
