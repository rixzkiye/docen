/**
 * Building blocks (Word's Quick Parts / AutoText / Building Blocks Organizer) —
 * the docen model, its persistence in the document, and its projection onto
 * Word's real glossary part (word/glossary/document.xml).
 *
 * Persistence decision (D2 investigation): @office-open/docx parses the
 * glossary part into DocumentOptions.glossary and writes it back on generate,
 * and docen's documentExtras passthrough carries that field through
 * resolve→compile untouched (pinned by packages/docx/src/converters/
 * glossary.spec.ts). Blocks therefore persist in the DOCUMENT, like Word's
 * per-document/template storage:
 *
 *   documentExtras.docenBlocks = { version, blocks: BuildingBlock[] }
 *   documentExtras.glossary    = derived GlossaryDocumentOptions (the real
 *                                Word part, regenerated on every block change)
 *
 * A DOCX that carries only a glossary (created in Word) is imported lazily on
 * read — `blocksOfDocAttrs` falls back to `blocksFromGlossary`. Corruption is
 * tolerated at every level: a foreign version or malformed entry drops to the
 * glossary/defaults instead of throwing.
 */

import {
  compileDocument,
  DocPartBehavior,
  DocPartGallery,
  resolveDocument,
  type DocPartSectionOptions,
  type GlossaryDocumentOptions,
  type JSONContent,
  type SectionChild,
} from "@docen/docx";

/** The docen block-list schema version. */
export const BUILDING_BLOCKS_VERSION = 1;

/** Word's three insert modes, mapped to the docPart behaviors. */
export type BuildingBlockInsertMode = "content" | "paragraph" | "page";

/** A selection-shaped Tiptap slice (the repo's DOCEN_CLIP payload shape). */
export interface BuildingBlockSlice {
  openStart: number;
  openEnd: number;
  content: JSONContent[];
}

/** One saved building block. */
export interface BuildingBlock {
  id: string;
  name: string;
  /** A DocPartGallery token (see {@link BLOCK_GALLERIES}). */
  gallery: string;
  category: string;
  description: string;
  /** Epoch ms (docen metadata — OOXML has no timestamp slot). */
  savedAt: number;
  insertMode: BuildingBlockInsertMode;
  content: BuildingBlockSlice;
}

/** The persisted docen block list. */
export interface BuildingBlocksData {
  version: number;
  blocks: BuildingBlock[];
}

/** The galleries the UI offers — the CUSTOM_* docPart galleries Word uses for
 *  user-created building blocks, in the organizer's order. */
export const BLOCK_GALLERIES: readonly DocPartGallery[] = [
  DocPartGallery.CUSTOM_QUICK_PARTS,
  DocPartGallery.CUSTOM_AUTO_TEXT,
  DocPartGallery.CUSTOM_COVER_PAGE,
  DocPartGallery.CUSTOM_TABLES,
  DocPartGallery.CUSTOM_EQUATIONS,
  DocPartGallery.CUSTOM_HEADERS,
  DocPartGallery.CUSTOM_FOOTERS,
];

/** Word's default gallery for a saved selection. */
export const DEFAULT_BLOCK_GALLERY: DocPartGallery = DocPartGallery.CUSTOM_QUICK_PARTS;

/** Word's default category. */
export const DEFAULT_BLOCK_CATEGORY = "General";

/** True for a gallery token the UI can round-trip. */
export function isBlockGallery(token: unknown): token is DocPartGallery {
  return typeof token === "string" && (BLOCK_GALLERIES as readonly string[]).includes(token);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

/** A fresh block id — crypto.randomUUID where present (browser + Node 19+). */
function newId(): string {
  const cryptoApi = globalThis.crypto as { randomUUID?: () => string } | undefined;
  return (
    cryptoApi?.randomUUID?.() ?? `block-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`
  );
}

/** Clamp a persisted slice offset. */
function sliceOffset(value: unknown): number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0
    ? Math.min(Math.floor(value), 64)
    : 0;
}

/** Defensive parse of one block's slice — null when the content list is
 *  missing/empty or not JSON nodes. */
function parseSlice(value: unknown): BuildingBlockSlice | null {
  if (!isRecord(value) || !Array.isArray(value.content)) return null;
  const content = value.content.filter((node) => isRecord(node));
  if (content.length === 0) return null;
  return {
    openStart: sliceOffset(value.openStart),
    openEnd: sliceOffset(value.openEnd),
    content: content as JSONContent[],
  };
}

function pickString(value: unknown, fallback = ""): string {
  return typeof value === "string" ? value : fallback;
}

/** Defensive parse of one persisted block — null when malformed. */
function parseBlock(value: unknown): BuildingBlock | null {
  if (!isRecord(value)) return null;
  const name = pickString(value.name).trim();
  const content = parseSlice(value.content);
  if (!name || !content) return null;
  const gallery = isBlockGallery(value.gallery) ? value.gallery : DEFAULT_BLOCK_GALLERY;
  const insertMode = value.insertMode;
  return {
    id: pickString(value.id).trim() || newId(),
    name,
    gallery,
    category: pickString(value.category).trim() || DEFAULT_BLOCK_CATEGORY,
    description: pickString(value.description),
    savedAt:
      typeof value.savedAt === "number" && Number.isFinite(value.savedAt) ? value.savedAt : 0,
    insertMode:
      insertMode === "paragraph" || insertMode === "page" || insertMode === "content"
        ? insertMode
        : "content",
    content,
  };
}

/** The persisted block list, sanitized. A foreign/missing version yields an
 *  empty list (the caller falls back to the glossary or to no blocks). */
export function parseBuildingBlocks(value: unknown): BuildingBlocksData {
  if (!isRecord(value) || value.version !== BUILDING_BLOCKS_VERSION) {
    return { version: BUILDING_BLOCKS_VERSION, blocks: [] };
  }
  const blocks: BuildingBlock[] = [];
  const seen = new Set<string>();
  if (Array.isArray(value.blocks)) {
    for (const raw of value.blocks) {
      const block = parseBlock(raw);
      if (!block || seen.has(block.id)) continue;
      seen.add(block.id);
      blocks.push(block);
    }
  }
  return { version: BUILDING_BLOCKS_VERSION, blocks };
}

/** The documentExtras record of a doc-attrs bag, or undefined. */
function extrasOf(attrs: unknown): Record<string, unknown> | undefined {
  if (!isRecord(attrs)) return undefined;
  const extras = attrs.documentExtras;
  return isRecord(extras) ? extras : undefined;
}

// ── Glossary projection (the real DOCX part) ────────────────────────────────

/** The docPart behavior list for an insert mode. */
function behaviorsOf(mode: BuildingBlockInsertMode): DocPartBehavior[] {
  switch (mode) {
    case "paragraph":
      return [DocPartBehavior.PARAGRAPH];
    case "page":
      return [DocPartBehavior.PAGE];
    default:
      return [DocPartBehavior.CONTENT];
  }
}

/** The insert mode a docPart behavior list implies (Word's default: content). */
function insertModeOf(behaviors: unknown): BuildingBlockInsertMode {
  if (!Array.isArray(behaviors)) return "content";
  if (behaviors.includes(DocPartBehavior.PAGE)) return "page";
  if (behaviors.includes(DocPartBehavior.PARAGRAPH)) return "paragraph";
  return "content";
}

/** Compile a block slice into docPart sections (office-open's persistence
 *  shape) so Word reads the block from its Building Blocks gallery. Returns
 *  an empty list when the content cannot compile — the docen list still
 *  persists; only the DOCX projection skips that block. */
function sectionsOf(slice: BuildingBlockSlice): DocPartSectionOptions[] {
  try {
    const compiled = compileDocument({ type: "doc", content: slice.content });
    return compiled.sections.map((section) => ({
      children: section.children,
      ...(section.properties ? { properties: section.properties } : {}),
    }));
  } catch {
    return [];
  }
}

/** The real word/glossary/document.xml content for a block list. Blocks whose
 *  content cannot compile are left out of the part (they stay in the docen
 *  list, which is the runtime source of truth). */
export function glossaryOfBlocks(blocks: readonly BuildingBlock[]): GlossaryDocumentOptions {
  const parts = blocks.flatMap((block) => {
    const sections = sectionsOf(block.content);
    if (sections.length === 0) return [];
    return [
      {
        name: block.name,
        gallery: isBlockGallery(block.gallery) ? block.gallery : DEFAULT_BLOCK_GALLERY,
        ...(block.category ? { category: block.category } : {}),
        ...(block.description ? { description: block.description } : {}),
        guid: block.id,
        behaviors: behaviorsOf(block.insertMode),
        sections,
      },
    ];
  });
  return { parts };
}

/** Strip the resolver's section furniture (sectPr/header/footer attrs) off a
 *  node — an inserted block must not carry the part's section setup. */
function stripSectionAttrs(node: JSONContent): JSONContent {
  const attrs = node.attrs;
  if (
    !attrs ||
    (attrs.sectionProperties == null &&
      attrs.sectionHeaders == null &&
      attrs.sectionFooters == null)
  ) {
    return node;
  }
  const { sectionProperties: _sp, sectionHeaders: _sh, sectionFooters: _sf, ...rest } = attrs;
  return { ...node, attrs: rest };
}

/** Resolve one glossary docPart back into a Tiptap slice. Null when the part
 *  has no resolvable sections. */
function contentOfPart(part: Record<string, unknown>): BuildingBlockSlice | null {
  const rawSections = Array.isArray(part.sections) ? part.sections : [];
  const sections = rawSections.flatMap((raw) => {
    if (!isRecord(raw) || !Array.isArray(raw.children)) return [];
    return [
      {
        children: raw.children as SectionChild[],
        ...(isRecord(raw.properties)
          ? { properties: raw.properties as DocPartSectionOptions["properties"] }
          : {}),
      },
    ];
  });
  if (sections.length === 0) return null;
  try {
    const json = resolveDocument({ sections });
    const nodes = json.content ?? [];
    const content = nodes.map(stripSectionAttrs);
    // The resolver appends an empty paragraph to carry a section's sectPr when
    // the part ends on a non-textblock — drop that synthetic node.
    const last = nodes[nodes.length - 1];
    if (
      last?.type === "paragraph" &&
      last.attrs?.sectionProperties != null &&
      !last.content?.length
    ) {
      content.pop();
    }
    return { openStart: 0, openEnd: 0, content };
  } catch {
    // Content this schema cannot represent — the part is skipped, never fatal.
    return null;
  }
}

/** Import Word glossary parts into the docen block model. Malformed parts are
 *  skipped; `savedAt` is unknown for imported parts (0). */
export function blocksFromGlossary(glossary: unknown): BuildingBlock[] {
  if (!isRecord(glossary) || !Array.isArray(glossary.parts)) return [];
  return glossary.parts.flatMap((raw, index) => {
    if (!isRecord(raw)) return [];
    const name = pickString(raw.name).trim();
    if (!name) return [];
    const content = contentOfPart(raw);
    if (!content) return [];
    return [
      {
        id: pickString(raw.guid).trim() || `glossary-${index}`,
        name,
        gallery: isBlockGallery(raw.gallery) ? raw.gallery : DEFAULT_BLOCK_GALLERY,
        category: pickString(raw.category).trim() || DEFAULT_BLOCK_CATEGORY,
        description: pickString(raw.description),
        savedAt: 0,
        insertMode: insertModeOf(raw.behaviors),
        content,
      },
    ];
  });
}

// ── Document-attrs persistence ───────────────────────────────────────────────

/** The document's blocks: the docen list when present (even empty — deleting
 *  every block is a real state), else a lazy import of the Word glossary. */
export function blocksOfDocAttrs(attrs: unknown): BuildingBlock[] {
  const extras = extrasOf(attrs);
  const raw = extras?.docenBlocks;
  if (isRecord(raw) && raw.version === BUILDING_BLOCKS_VERSION && Array.isArray(raw.blocks)) {
    return parseBuildingBlocks(raw).blocks;
  }
  return blocksFromGlossary(extras?.glossary);
}

/** A new documentExtras record carrying `blocks` and the derived glossary
 *  part; existing unrelated extras keys are preserved. An empty list removes
 *  the glossary (the user deleted every block). */
export function withBlocks(
  extras: unknown,
  blocks: readonly BuildingBlock[],
): Record<string, unknown> {
  const next = isRecord(extras) ? { ...extras } : {};
  next.docenBlocks = {
    version: BUILDING_BLOCKS_VERSION,
    blocks: blocks.map((block) => ({
      ...block,
      content: { ...block.content, content: [...block.content.content] },
    })),
  };
  // The glossary part is the DOCX carrier: projected best-effort (a block
  // whose content cannot compile still persists in the docen list). An empty
  // document with no blocks carries no glossary at all.
  const glossary = blocks.length > 0 ? glossaryOfBlocks(blocks) : { parts: [] };
  if (glossary.parts.length > 0) next.glossary = glossary;
  else delete next.glossary;
  return next;
}

/** Parse the clipboard-lane slice payload (`selectionSlicePayload`). */
export function parseSlicePayload(raw: string | null | undefined): BuildingBlockSlice | null {
  if (!raw) return null;
  try {
    return parseSlice(JSON.parse(raw));
  } catch {
    return null;
  }
}

/** Build a block from a save-dialog commit. */
export function createBuildingBlock(
  input: {
    name: string;
    gallery?: string;
    category?: string;
    description?: string;
    insertMode?: BuildingBlockInsertMode;
    content: BuildingBlockSlice;
  },
  options: { id?: string; savedAt?: number } = {},
): BuildingBlock {
  return {
    id: options.id ?? newId(),
    name: input.name.trim(),
    gallery: isBlockGallery(input.gallery) ? input.gallery : DEFAULT_BLOCK_GALLERY,
    category: (input.category ?? DEFAULT_BLOCK_CATEGORY).trim() || DEFAULT_BLOCK_CATEGORY,
    description: (input.description ?? "").trim(),
    savedAt: options.savedAt ?? Date.now(),
    insertMode: input.insertMode ?? "content",
    content: input.content,
  };
}

/** True when `name` (case-insensitive) is already in a name list. */
export function isDuplicateName(names: readonly string[], name: string): boolean {
  const target = name.trim().toLowerCase();
  if (!target) return false;
  return names.some((existing) => existing.trim().toLowerCase() === target);
}

/** True when `name` (case-insensitive) already belongs to another block. */
export function isDuplicateBlockName(
  blocks: readonly BuildingBlock[],
  name: string,
  exceptId?: string,
): boolean {
  return isDuplicateName(
    blocks.filter((block) => block.id !== exceptId).map((block) => block.name),
    name,
  );
}

/** Blocks ordered for the gallery: by gallery order, then category, then name
 *  (case-insensitive, deterministic). */
export function sortBlocks(blocks: readonly BuildingBlock[]): BuildingBlock[] {
  const galleryRank = new Map(BLOCK_GALLERIES.map((gallery, index) => [gallery as string, index]));
  const key = (block: BuildingBlock): string =>
    `${String(galleryRank.get(block.gallery) ?? BLOCK_GALLERIES.length)}` +
    `|${block.category.toLowerCase()}|${block.name.toLowerCase()}`;
  return [...blocks].sort((a, b) => (key(a) < key(b) ? -1 : key(a) > key(b) ? 1 : 0));
}

/** The gallery → blocks grouping the Quick Parts menu renders. */
export function groupBlocksByGallery(
  blocks: readonly BuildingBlock[],
): Array<{ gallery: string; blocks: BuildingBlock[] }> {
  const groups: Array<{ gallery: string; blocks: BuildingBlock[] }> = [];
  for (const block of sortBlocks(blocks)) {
    let group = groups[groups.length - 1];
    if (!group || group.gallery !== block.gallery) {
      group = { gallery: block.gallery, blocks: [] };
      groups.push(group);
    }
    group.blocks.push(block);
  }
  return groups;
}

/** The F3 (AutoText) match for the text before the caret: the longest block
 *  name ending exactly at the caret, case-insensitive, on a word boundary.
 *  `back` is how many characters the name occupies (what the insert
 *  replaces). Null when no name matches (partial names never do). */
export function autotextMatch(
  textBefore: string,
  blocks: readonly BuildingBlock[],
): { block: BuildingBlock; back: number } | null {
  if (!textBefore) return null;
  const lower = textBefore.toLowerCase();
  let best: { block: BuildingBlock; back: number } | null = null;
  for (const block of blocks) {
    const name = block.name.trim();
    if (!name) continue;
    const nameLower = name.toLowerCase();
    if (!lower.endsWith(nameLower)) continue;
    const start = textBefore.length - name.length;
    // The name must start at the paragraph start or after a boundary — typing
    // it inside a word ("xSignature") never matches.
    if (start > 0 && !/[\s([{<"'“‘—–]/u.test(textBefore[start - 1]!)) continue;
    if (!best || name.length > best.block.name.length) best = { block, back: name.length };
  }
  return best;
}
