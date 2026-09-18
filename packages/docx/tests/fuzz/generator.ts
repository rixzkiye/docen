// Deterministic DOCX fuzz generator + shrinker (item 15).
//
// `generateFuzzCase(seed)` builds a bounded DocumentOptions document across
// the audited feature matrix: marks, styles, tables with merges, nested
// lists, fields, images, hyperlinks/bookmarks, notes, multi-section layout,
// headers/footers and page numbering. The same seed always yields the same
// document, so a failing seed is a permanent regression test.
//
// `shrinkCase` is a bounded delta-debugger over the generated JSON: it removes
// sections/blocks/runs/attrs while the supplied failure predicate still
// reports a failure, recording the predicate-call count.

export interface FuzzCase {
  seed: number;
  features: string[];
  doc: Record<string, unknown>;
}

class Rng {
  private state: number;

  constructor(seed: number) {
    this.state = seed >>> 0 || 0x9e3779b9;
  }

  next(): number {
    let t = (this.state += 0x6d2b79f5);
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  }

  int(max: number): number {
    return Math.floor(this.next() * max);
  }

  bool(probability = 0.5): boolean {
    return this.next() < probability;
  }

  pick<T>(items: readonly T[]): T {
    return items[this.int(items.length)]!;
  }
}

const TEXTS = [
  "alpha",
  "bravo",
  "charlie",
  "delta",
  "café",
  "中文文本",
  "→ arrow",
  "100% done",
  "a/b\\c",
  "line one",
];

const IMG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const PAGE = {
  pageSize: { width: 11906, height: 16838 },
  pageMargin: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
};

const RUN_PROPS = [
  { bold: true },
  { italic: true },
  { underline: { type: "single" } },
  { strike: true },
  { doubleStrike: true },
  { color: "FF0000" },
  { size: 28 },
  { highlight: "yellow" },
  { verticalAlign: "superscript" },
  { verticalAlign: "subscript" },
  { smallCaps: true },
  { allCaps: true },
  { characterSpacing: 20 },
  {},
  {},
  {},
];

const ALIGNMENTS = ["left", "center", "right", "both", "distribute"];
const NUMBERING = [
  { reference: "docen-bullet", level: 0 },
  { reference: "docen-bullet", level: 1 },
  { reference: "docen-bullet", level: 2 },
  { reference: "docen-ordered-1", level: 0 },
  { reference: "docen-ordered-1", level: 1 },
  { reference: "docen-multilevel-cascade-1", level: 0 },
  { reference: "docen-multilevel-hybrid-1", level: 1 },
];

const NOTE_IDS = [1, 2, 3];

interface DocState {
  footnotesUsed: Set<number>;
  endnotesUsed: Set<number>;
}

function picture(rng: Rng): Record<string, unknown> {
  const pictureOpts: Record<string, unknown> = {
    type: "png",
    data: IMG,
    transformation: { width: "24px", height: "24px" },
  };
  if (rng.bool(0.3)) {
    pictureOpts.floating = {
      horizontalPosition: { relative: "column", align: "center" },
      verticalPosition: { relative: "paragraph", offset: 45720 },
    };
  }
  if (rng.bool(0.3)) {
    pictureOpts.transformation = { width: "24px", height: "24px", rotation: 90 };
  }
  return { picture: pictureOpts };
}

function inlineContent(rng: Rng, state: DocState, allowRefs: boolean): unknown[] {
  const children: unknown[] = [];
  const count = 1 + rng.int(4);
  for (let index = 0; index < count; index++) {
    const roll = rng.next();
    if (allowRefs && roll < 0.1) {
      const id = rng.pick(NOTE_IDS);
      state.footnotesUsed.add(id);
      children.push({ footnoteReference: id });
    } else if (allowRefs && roll < 0.16) {
      const id = 1;
      state.endnotesUsed.add(id);
      children.push({ endnoteReference: id });
    } else if (roll < 0.24) {
      children.push({
        complexField: {
          instruction: rng.pick([" PAGE ", " NUMPAGES ", " CREATEDATE ", " DATE "]),
          result: rng.bool(0.5) ? "1" : undefined,
        },
      });
    } else if (roll < 0.3) {
      children.push({ simpleField: { instruction: " PAGE ", cachedValue: "1" } });
    } else if (roll < 0.38) {
      children.push({
        hyperlink: {
          url: `https://example.com/${index}`,
          tooltip: "fuzz link",
          children: [{ text: rng.pick(TEXTS) }],
        },
      });
    } else if (roll < 0.44 && state.footnotesUsed.size < 6) {
      children.push(picture(rng));
    } else if (roll < 0.48) {
      children.push({ pageBreak: true });
    } else {
      const text = rng.pick(TEXTS);
      const run: Record<string, unknown> = { text };
      Object.assign(run, rng.pick(RUN_PROPS));
      children.push(run);
    }
  }
  if (children.length === 0) children.push({ text: rng.pick(TEXTS) });
  return children;
}

function paragraph(rng: Rng, state: DocState, allowRefs: boolean, lists: boolean): unknown {
  const paragraphOpts: Record<string, unknown> = {
    children: inlineContent(rng, state, allowRefs),
  };
  if (rng.bool(0.3)) paragraphOpts.alignment = rng.pick(ALIGNMENTS);
  if (rng.bool(0.2)) paragraphOpts.indent = { left: 720, hanging: 360 };
  if (rng.bool(0.2)) paragraphOpts.spacing = { after: 0, line: 276, lineRule: "auto" };
  if (rng.bool(0.15)) paragraphOpts.style = "Heading2";
  if (rng.bool(0.1)) paragraphOpts.keepNext = true;
  if (rng.bool(0.08)) paragraphOpts.pageBreakBefore = true;
  if (rng.bool(0.08)) {
    paragraphOpts.shading = { fill: "D9E2F3", type: "clear" };
  }
  if (lists && rng.bool(0.25)) paragraphOpts.numbering = rng.pick(NUMBERING);
  return { paragraph: paragraphOpts };
}

function table(rng: Rng, state: DocState): unknown {
  const grid = 2 + rng.int(2); // 2..3 columns
  const rowCount = 2 + rng.int(2); // 2..3 rows
  const rows: unknown[] = [];
  const mergeColumn = rng.int(grid);
  const mergeRows = rng.bool(0.5) && rowCount >= 2;
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    const cells: unknown[] = [];
    for (let column = 0; column < grid; column++) {
      if (mergeRows && column === mergeColumn && rowIndex > 0) {
        cells.push({
          verticalMerge: "continue",
          children: [paragraph(rng, state, false, false)],
        });
        continue;
      }
      const span = column === 0 && rng.bool(0.3) ? 2 : 1;
      const cell: Record<string, unknown> = {
        children: [
          paragraph(rng, state, false, false),
          ...(rng.bool(0.3) ? [paragraph(rng, state, false, false)] : []),
        ],
      };
      if (span > 1 && column + span <= grid) cell.columnSpan = span;
      if (mergeRows && column === mergeColumn && rowIndex === 0) cell.verticalMerge = "restart";
      if (rng.bool(0.2)) cell.shading = { fill: "F2F2F2", type: "clear" };
      if (rng.bool(0.1)) cell.verticalAlign = "center";
      cells.push(cell);
      if (span > 1) column += span - 1;
    }
    const row: Record<string, unknown> = { cells };
    if (rowIndex === 0 && rng.bool(0.5)) row.tableHeader = true;
    if (rng.bool(0.3)) row.cantSplit = true;
    rows.push(row);
  }
  const tableOpts: Record<string, unknown> = { rows, width: { size: 5000, type: "pct" } };
  if (rng.bool(0.2)) tableOpts.alignment = "center";
  if (rng.bool(0.2)) {
    tableOpts.borders = {
      top: { style: "single", size: 4, color: "000000" },
      bottom: { style: "single", size: 4, color: "000000" },
    };
  }
  return { table: tableOpts };
}

function noteChildren(rng: Rng, state: DocState, refMark: string): unknown[] {
  const children: unknown[] = [
    {
      paragraph: {
        style: refMark === "footnoteRef" ? "FootnoteText" : "EndnoteText",
        children: [{ [refMark]: true }, { text: ` ${rng.pick(TEXTS)}` }],
      },
    },
  ];
  if (rng.bool(0.4)) children.push(paragraph(rng, state, false, false));
  if (rng.bool(0.2)) children.push(table(rng, state));
  if (rng.bool(0.2)) children.push({ paragraph: { children: [picture(rng)] } });
  return children;
}

function headerFooterSlot(rng: Rng, state: DocState): unknown[] {
  const slot: unknown[] = [paragraph(rng, state, false, false)];
  if (rng.bool(0.4)) {
    slot.push({
      paragraph: {
        alignment: "center",
        children: [{ text: "PAGE " }, { complexField: { instruction: " PAGE ", result: "1" } }],
      },
    });
  }
  if (rng.bool(0.2)) slot.push({ paragraph: { children: [picture(rng)] } });
  return slot;
}

/** Build one deterministic fuzz case. */
export function generateFuzzCase(seed: number): FuzzCase {
  const rng = new Rng(seed);
  const state: DocState = { footnotesUsed: new Set(), endnotesUsed: new Set() };
  const features: string[] = [];
  const has = (name: string, probability: number): boolean => {
    if (rng.bool(probability)) {
      features.push(name);
      return true;
    }
    return false;
  };
  const notes = has("notes", 0.7);
  const lists = has("lists", 0.6);
  const tables = has("tables", 0.6);
  const headers = has("headers", 0.6);
  const sections = has("sections", 0.5);
  const images = has("images", 0.4) || notes;
  if (images && !features.includes("images")) features.push("images");

  const sectionCount = sections ? 1 + rng.int(2) : 1;
  const doc: Record<string, unknown> = { sections: [] };
  for (let sectionIndex = 0; sectionIndex < sectionCount; sectionIndex++) {
    const properties: Record<string, unknown> = { ...PAGE };
    if (rng.bool(0.4)) properties.titlePage = true;
    if (rng.bool(0.3)) {
      properties.pageNumberType = {
        format: rng.pick(["decimal", "lowerRoman", "upperRoman", "lowerLetter"]),
        start: 1 + rng.int(5),
      };
    }
    if (notes && rng.bool(0.4)) {
      properties.footnoteProperties = {
        formatType: rng.pick(["decimal", "lowerRoman", "upperLetter"]),
        numStart: 1 + rng.int(5),
        numRestart: rng.pick(["continuous", "eachSect", "eachPage"]),
      };
    }
    if (notes && rng.bool(0.3)) {
      properties.endnoteProperties = {
        formatType: rng.pick(["decimal", "lowerRoman", "upperRoman"]),
        pos: rng.pick(["sectEnd", "docEnd"]),
      };
    }
    if (rng.bool(0.2)) properties.columns = { count: 2, space: 720 };
    const children: unknown[] = [];
    const section: Record<string, unknown> = { properties, children };
    const blocks = 1 + rng.int(4);
    for (let block = 0; block < blocks; block++) {
      if (tables && rng.bool(0.3)) children.push(table(rng, state));
      else children.push(paragraph(rng, state, true, lists));
    }
    if (headers) {
      const slots = ["default", "first", "even"];
      const group: Record<string, unknown> = {};
      for (const slot of slots) {
        if (slot === "first" && properties.titlePage !== true) continue;
        if (rng.bool(0.5)) group[slot] = headerFooterSlot(rng, state);
      }
      if (Object.keys(group).length > 0) section.headers = group;
      if (rng.bool(0.5)) {
        section.footers = { default: headerFooterSlot(rng, state) };
      }
    }
    (doc.sections as unknown[]).push(section);
  }

  if (notes && state.footnotesUsed.size > 0) {
    doc.footnotes = [...state.footnotesUsed]
      .sort((a, b) => a - b)
      .map((id) => ({
        id,
        children: noteChildren(rng, state, "footnoteRef"),
      }));
  }
  if (notes && state.endnotesUsed.size > 0) {
    doc.endnotes = [...state.endnotesUsed]
      .sort((a, b) => a - b)
      .map((id) => ({
        id,
        children: noteChildren(rng, state, "endnoteRef"),
      }));
  }
  if (rng.bool(0.2)) doc.settings = { evenAndOddHeaders: true };
  if (notes && rng.bool(0.2)) {
    doc.footnoteSeparators = {
      separator: { id: -1, children: [{ paragraph: { children: [{ text: "SEP" }] } }] },
    };
  }
  return { seed, features: [...new Set(features)].sort(), doc };
}

/**
 * Bounded delta-debugging shrinker. `fails(value)` returns a failure message
 * (or null when the value passes); the shrinker returns a structurally
 * reduced document that still fails, plus the predicate-call count.
 */
export function shrinkCase(
  doc: Record<string, unknown>,
  fails: (candidate: Record<string, unknown>) => string | null,
  budget = 80,
): { doc: Record<string, unknown>; calls: number } {
  const state = { calls: 0, budget };
  const predicate = (candidate: unknown): boolean => {
    state.calls++;
    return fails(candidate as Record<string, unknown>) !== null;
  };
  const shrunk = shrinkValue(structuredClone(doc), predicate, state) as Record<string, unknown>;
  return { doc: shrunk, calls: state.calls };
}

function shrinkValue(
  value: unknown,
  fails: (candidate: unknown) => boolean,
  state: { calls: number; budget: number },
): unknown {
  if (state.calls >= state.budget) return value;
  if (Array.isArray(value)) {
    let list = value;
    for (let index = list.length - 1; index >= 0 && list.length > 1; index--) {
      if (state.calls >= state.budget) break;
      const candidate = list.slice();
      candidate.splice(index, 1);
      if (fails(candidate)) {
        list = candidate;
        index = list.length;
      }
    }
    return list.map((entry) => shrinkValue(entry, fails, state));
  }
  if (value && typeof value === "object") {
    const record = value as Record<string, unknown>;
    const keep = new Set(["sections", "children", "rows", "cells", "content"]);
    for (const key of Object.keys(record)) {
      if (keep.has(key) || state.calls >= state.budget) continue;
      const candidate = { ...record };
      delete candidate[key];
      if (fails(candidate)) delete record[key];
    }
    for (const key of Object.keys(record)) {
      record[key] = shrinkValue(record[key], fails, state);
    }
    return record;
  }
  return value;
}
