import type { JSONContent } from "../core";
import { decodePassthroughData } from "../extensions/passthrough";

/**
 * Detection of document content that the editor carries but cannot edit —
 * `altChunk` imports, `subDoc` references, SmartArt diagrams, embedded OLE
 * objects, raw/custom XML and content parts. The passthrough atoms round-trip
 * these byte-faithfully (see docs/passthrough.md), but a user must not be
 * silently misled into thinking the whole document is editable: the editor
 * reads this report after a load and shows the unsupported-content warning bar.
 *
 * Glossary parts are NOT reported: office-open parses them into
 * `documentExtras.glossary` and the editor's Quick Parts surfaces edit/insert
 * them, so they are supported content.
 */

export type UnsupportedContentKind =
  | "altChunk"
  | "subDoc"
  | "smartArt"
  | "oleObject"
  | "rawXml"
  | "customXml"
  | "contentPart";

export interface UnsupportedContentItem {
  readonly kind: UnsupportedContentKind;
  readonly count: number;
}

export interface UnsupportedContentReport {
  readonly items: readonly UnsupportedContentItem[];
  /** Total number of unsupported atoms found (sum of item counts). */
  readonly total: number;
}

/** Every kind the report can contain, in stable display order. */
export const UNSUPPORTED_CONTENT_KINDS: readonly UnsupportedContentKind[] = [
  "altChunk",
  "subDoc",
  "smartArt",
  "oleObject",
  "rawXml",
  "customXml",
  "contentPart",
];

/** Block `passthrough` atom tags the report recognizes. */
const BLOCK_TAGS: Readonly<Record<string, UnsupportedContentKind>> = {
  altChunk: "altChunk",
  rawXml: "rawXml",
  customXml: "customXml",
};

/** Inline `inlinePassthrough` atom tags the report recognizes. Comment range
 *  markers, bookmarks, proofErr, symbol runs and fields are deliberately NOT
 *  listed: they are metadata with their own visible semantics, not content a
 *  user would expect to edit. */
const INLINE_TAGS: Readonly<Record<string, UnsupportedContentKind>> = {
  smartArt: "smartArt",
  object: "oleObject",
  subDoc: "subDoc",
  rawXml: "rawXml",
  customXml: "customXml",
  contentPart: "contentPart",
};

/** The office-open union tag inside a passthrough atom's encoded data. */
function passthroughTag(attrs: Record<string, unknown> | undefined): string | undefined {
  const data = attrs?.data;
  if (typeof data !== "string") return undefined;
  try {
    const decoded = decodePassthroughData<Record<string, unknown>>(data);
    return Object.keys(decoded)[0];
  } catch {
    return undefined;
  }
}

/** Count the uneditable container content in a document JSON tree. */
export function detectUnsupportedContent(json: JSONContent): UnsupportedContentReport {
  const counts = new Map<UnsupportedContentKind, number>();
  const walk = (node: JSONContent): void => {
    const tag = passthroughTag(node.attrs as Record<string, unknown> | undefined);
    if (tag) {
      const kind =
        node.type === "passthrough"
          ? BLOCK_TAGS[tag]
          : node.type === "inlinePassthrough"
            ? INLINE_TAGS[tag]
            : undefined;
      if (kind) counts.set(kind, (counts.get(kind) ?? 0) + 1);
    }
    for (const child of node.content ?? []) walk(child);
  };
  walk(json);
  const items = UNSUPPORTED_CONTENT_KINDS.filter((kind) => counts.has(kind)).map((kind) => ({
    kind,
    count: counts.get(kind)!,
  }));
  return { items, total: items.reduce((sum, item) => sum + item.count, 0) };
}
