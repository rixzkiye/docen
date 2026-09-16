import { Extension } from "@docen/docx/core";
import type { Node as PMNode, Schema } from "@tiptap/pm/model";
import { Fragment, Slice } from "@tiptap/pm/model";
import type { Transaction } from "@tiptap/pm/state";
import { DocAttrStep } from "@tiptap/pm/transform";

import type { PageOf } from "./toc";

/**
 * Index & Table of Authorities commands — ribbon References tab:
 * - Index: Mark Entry, Mark All, Insert Index, Update Index.
 * - Table of Authorities: Mark Citation, Insert TOA, Update TOA.
 *
 * Marking an entry seeds an `XE "…"` field (invisible index marker).
 * Marking a citation seeds a `TA \l "…" \s "…" \c N` field.
 *
 * Index entries ride Index1/Index2 styled paragraphs with dotted leader tabs.
 * TOA entries ride TOAHeading and TableOfAuthorities styled paragraphs.
 */

/** Word's built-in index styles indent 220 twips per level. */
const INDEX_INDENT_TW = 220;

export interface XeFieldData {
  entry: string;
  bold?: boolean;
  italic?: boolean;
  crossReference?: string;
}

export function parseXeInstruction(instruction: string): XeFieldData | null {
  const trimmed = instruction.trim();
  if (!/^XE\b/i.test(trimmed)) return null;
  const m = /^XE\s+(?:"([^"]*)"|([^\s\\]+))/i.exec(trimmed);
  if (!m) return null;
  const entry = m[1] ?? m[2];
  const rest = trimmed.slice(m[0].length);

  const bold = /(?:^|\s)\\b(?:\s|$)/i.test(rest);
  const italic = /(?:^|\s)\\i(?:\s|$)/i.test(rest);
  const tMatch = /(?:^|\s)\\t\s+"([^"]*)"/i.exec(rest) || /(?:^|\s)\\t\s+([^\s\\]+)/i.exec(rest);
  const crossReference = tMatch ? (tMatch[1] ?? tMatch[2]) : undefined;

  return { entry, bold, italic, crossReference };
}

export interface TaFieldData {
  longCitation: string;
  shortCitation?: string;
  category: number;
  bold?: boolean;
  italic?: boolean;
}

export function parseTaInstruction(instruction: string): TaFieldData | null {
  const trimmed = instruction.trim();
  if (!/^TA\b/i.test(trimmed)) return null;

  const lMatch = /\\l\s+"([^"]*)"/i.exec(trimmed);
  if (!lMatch) return null;
  const longCitation = lMatch[1];

  const sMatch = /\\s\s+"([^"]*)"/i.exec(trimmed);
  const shortCitation = sMatch ? sMatch[1] : undefined;

  const cMatch = /\\c\s+(\d+)/i.exec(trimmed);
  const category = cMatch ? parseInt(cMatch[1], 10) : 1;

  const bold = /(?:^|\s)\\b(?:\s|$)/i.test(trimmed);
  const italic = /(?:^|\s)\\i(?:\s|$)/i.test(trimmed);

  return { longCitation, shortCitation, category, bold, italic };
}

export const TOA_DEFAULT_CATEGORIES: Record<number, string> = {
  1: "Cases",
  2: "Statutes",
  3: "Other Authorities",
  4: "Rules",
  5: "Treatises",
  6: "Regulations",
  7: "Constitutional Provisions",
};

export interface IndexPageRef {
  page: number;
  bold?: boolean;
  italic?: boolean;
}

export interface SubEntryRecord {
  pages: IndexPageRef[];
  crossReferences: string[];
}

/** Collect the document's XE fields, grouped by main and sub-entry. */
export function collectEntries(
  doc: PMNode,
  pageOf?: PageOf,
): Map<string, Map<string, SubEntryRecord>> {
  const groups = new Map<string, Map<string, SubEntryRecord>>();

  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return true;
    const page = pageOf?.(pos + 1);
    node.forEach((child) => {
      if (child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs.data ?? "{}")) as {
          simpleField?: { instruction?: string };
        };
        const instr = (data.simpleField?.instruction ?? "").trim();
        const xe = parseXeInstruction(instr);
        if (!xe) return;

        const sep = xe.entry.indexOf(":");
        const main = sep < 0 ? xe.entry : xe.entry.slice(0, sep);
        const sub = sep < 0 ? "" : xe.entry.slice(sep + 1);

        let subs = groups.get(main);
        if (!subs) {
          subs = new Map<string, SubEntryRecord>();
          groups.set(main, subs);
        }

        let record = subs.get(sub);
        if (!record) {
          record = { pages: [], crossReferences: [] };
          subs.set(sub, record);
        }

        if (xe.crossReference) {
          if (!record.crossReferences.includes(xe.crossReference)) {
            record.crossReferences.push(xe.crossReference);
          }
        } else if (typeof page === "number") {
          const existing = record.pages.find((p) => p.page === page);
          if (existing) {
            if (xe.bold) existing.bold = true;
            if (xe.italic) existing.italic = true;
          } else {
            record.pages.push({ page, bold: xe.bold, italic: xe.italic });
          }
        }
      } catch {
        // opaque verbatim blobs without field data — skip
      }
    });
    return true;
  });

  return groups;
}

/** Index entry paragraphs from the collected XE fields: mains sorted with
 *  document language collation, each `主:子` sub-entry nested under main. */
export function buildIndexParagraphs(
  groups: Map<string, Map<string, SubEntryRecord>>,
  tabPositionTw = 9350,
  schema?: Schema,
): { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] {
  const out: { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] = [];

  const entry = (text: string, level: number, record: SubEntryRecord) => {
    const afterTab: unknown[] = [];

    if (record.crossReferences.length > 0) {
      for (let i = 0; i < record.crossReferences.length; i++) {
        if (i > 0) afterTab.push({ type: "text", text: "; " });
        afterTab.push({
          type: "text",
          text: record.crossReferences[i],
          ...(schema?.marks?.italic ? { marks: [{ type: "italic" }] } : {}),
        });
      }
    }

    if (record.pages.length > 0) {
      if (afterTab.length > 0) afterTab.push({ type: "text", text: ", " });
      const sortedPages = [...record.pages].sort((a, b) => a.page - b.page);
      for (let i = 0; i < sortedPages.length; i++) {
        if (i > 0) afterTab.push({ type: "text", text: ", " });
        const p = sortedPages[i];
        const marks: { type: string }[] = [];
        if (p.bold && schema?.marks?.bold) marks.push({ type: "bold" });
        if (p.italic && schema?.marks?.italic) marks.push({ type: "italic" });
        afterTab.push({
          type: "text",
          text: String(p.page),
          ...(marks.length > 0 ? { marks } : {}),
        });
      }
    }

    return {
      type: "paragraph",
      attrs: {
        style: `Index${level + 1}`,
        tabStops: [{ type: "right", position: tabPositionTw, leader: "dot" }],
        ...(level > 0 ? { indent: { left: level * INDEX_INDENT_TW } } : {}),
      },
      content: [{ type: "text", text }, { type: "tab" }, ...afterTab],
    };
  };

  for (const main of [...groups.keys()].sort((a, b) => a.localeCompare(b, "zh"))) {
    const subs = groups.get(main)!;
    const own = subs.get("") ?? { pages: [], crossReferences: [] };
    out.push(entry(main, 0, own));
    for (const sub of [...subs.keys()]
      .filter((s) => s !== "")
      .sort((a, b) => a.localeCompare(b, "zh"))) {
      out.push(entry(sub, 1, subs.get(sub)!));
    }
  }

  return out;
}

export interface CitationRecord {
  longCitation: string;
  shortCitation?: string;
  category: number;
  pages: IndexPageRef[];
}

/** Collect the document's TA citation fields, grouped by category. */
export function collectCitations(
  doc: PMNode,
  pageOf?: PageOf,
): Map<number, Map<string, CitationRecord>> {
  const categories = new Map<number, Map<string, CitationRecord>>();

  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return true;
    const page = pageOf?.(pos + 1);
    node.forEach((child) => {
      if (child.type.name !== "inlinePassthrough") return;
      try {
        const data = JSON.parse(String(child.attrs.data ?? "{}")) as {
          simpleField?: { instruction?: string };
        };
        const instr = (data.simpleField?.instruction ?? "").trim();
        const ta = parseTaInstruction(instr);
        if (!ta) return;

        let catMap = categories.get(ta.category);
        if (!catMap) {
          catMap = new Map();
          categories.set(ta.category, catMap);
        }

        let record = catMap.get(ta.longCitation);
        if (!record && ta.shortCitation) {
          for (const existing of catMap.values()) {
            if (existing.shortCitation && existing.shortCitation === ta.shortCitation) {
              record = existing;
              break;
            }
          }
        }
        if (!record) {
          record = {
            longCitation: ta.longCitation,
            shortCitation: ta.shortCitation,
            category: ta.category,
            pages: [],
          };
          catMap.set(ta.longCitation, record);
        }

        if (typeof page === "number") {
          const existing = record.pages.find((p) => p.page === page);
          if (existing) {
            if (ta.bold) existing.bold = true;
            if (ta.italic) existing.italic = true;
          } else {
            record.pages.push({ page, bold: ta.bold, italic: ta.italic });
          }
        }
      } catch {
        // skip
      }
    });
    return true;
  });

  return categories;
}

/** Build Table of Authorities paragraphs grouped by category. */
export function buildToaParagraphs(
  categories: Map<number, Map<string, CitationRecord>>,
  tabPositionTw = 9350,
  schema?: Schema,
  options?: { category?: number; passim?: boolean },
): { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] {
  const out: { type: string; attrs?: Record<string, unknown>; content: unknown[] }[] = [];
  const usePassim = options?.passim !== false;

  const sortedCatIds = [...categories.keys()].sort((a, b) => a - b);
  for (const catId of sortedCatIds) {
    if (options?.category !== undefined && options.category !== catId) continue;
    const catMap = categories.get(catId);
    if (!catMap || catMap.size === 0) continue;

    const catName = TOA_DEFAULT_CATEGORIES[catId] ?? `Category ${catId}`;

    out.push({
      type: "paragraph",
      attrs: {
        style: "TOAHeading",
      },
      content: [
        {
          type: "text",
          text: catName,
          ...(schema?.marks?.bold ? { marks: [{ type: "bold" }] } : {}),
        },
      ],
    });

    const sortedEntries = [...catMap.values()].sort((a, b) =>
      a.longCitation.localeCompare(b.longCitation, "en"),
    );

    for (const record of sortedEntries) {
      const sortedPages = [...record.pages].sort((a, b) => a.page - b.page);
      const afterTab: unknown[] = [];

      if (usePassim && sortedPages.length >= 5) {
        afterTab.push({ type: "text", text: "passim" });
      } else {
        for (let i = 0; i < sortedPages.length; i++) {
          if (i > 0) afterTab.push({ type: "text", text: ", " });
          const p = sortedPages[i];
          const marks: { type: string }[] = [];
          if (p.bold && schema?.marks?.bold) marks.push({ type: "bold" });
          if (p.italic && schema?.marks?.italic) marks.push({ type: "italic" });
          afterTab.push({
            type: "text",
            text: String(p.page),
            ...(marks.length > 0 ? { marks } : {}),
          });
        }
      }

      out.push({
        type: "paragraph",
        attrs: {
          style: "TableOfAuthorities",
          tabStops: [{ type: "right", position: tabPositionTw, leader: "dot" }],
        },
        content: [{ type: "text", text: record.longCitation }, { type: "tab" }, ...afterTab],
      });
    }
  }

  return out;
}

/** The first Index-styled paragraph. */
function findIndexHead(doc: PMNode): number | null {
  let pos: number | null = null;
  doc.descendants((node, at) => {
    if (pos != null) return false;
    if (node.type.name === "paragraph" && /^Index\d$/.test(String(node.attrs.style ?? ""))) {
      pos = at;
      return false;
    }
    return true;
  });
  return pos;
}

/** The first TOA-styled paragraph. */
function findToaHead(doc: PMNode): number | null {
  let pos: number | null = null;
  doc.descendants((node, at) => {
    if (pos != null) return false;
    if (
      node.type.name === "paragraph" &&
      (node.attrs.style === "TOAHeading" || node.attrs.style === "TableOfAuthorities")
    ) {
      pos = at;
      return false;
    }
    return true;
  });
  return pos;
}

/** Stamp the Index1/Index2 style definitions when missing. */
function stampIndexStyles(tr: Transaction): void {
  const styles = { ...((tr.doc.attrs.styles ?? {}) as Record<string, unknown>) };
  const paragraphStyles = (styles.paragraphStyles ?? []) as {
    id?: string;
    name?: string;
    basedOn?: string;
    next?: string;
    indent?: { left?: number };
  }[];
  const missing = (["Index1", "Index2"] as const).filter(
    (id) => !paragraphStyles.some((style) => style.id === id),
  );
  if (missing.length === 0) return;
  for (const id of missing) {
    const level = Number(id.slice(-1)) - 1;
    paragraphStyles.push({
      id,
      name: `index ${id.slice(-1)}`,
      basedOn: "Normal",
      next: id,
      ...(level > 0 ? { indent: { left: level * INDEX_INDENT_TW } } : {}),
    });
  }
  tr.step(new DocAttrStep("styles", { ...styles, paragraphStyles }));
}

/** Stamp the TOAHeading/TableOfAuthorities style definitions when missing. */
function stampToaStyles(tr: Transaction): void {
  const styles = { ...((tr.doc.attrs.styles ?? {}) as Record<string, unknown>) };
  const paragraphStyles = (styles.paragraphStyles ?? []) as {
    id?: string;
    name?: string;
    basedOn?: string;
    next?: string;
  }[];
  const missing = (["TOAHeading", "TableOfAuthorities"] as const).filter(
    (id) => !paragraphStyles.some((style) => style.id === id),
  );
  if (missing.length === 0) return;
  for (const id of missing) {
    paragraphStyles.push({
      id,
      name: id === "TOAHeading" ? "toa heading" : "table of authorities",
      basedOn: id === "TOAHeading" ? "Heading 2" : "Normal",
      next: id === "TOAHeading" ? "TableOfAuthorities" : id,
    });
  }
  tr.step(new DocAttrStep("styles", { ...styles, paragraphStyles }));
}

export const IndexCommands = Extension.create({
  name: "docenIndexCommands",

  addCommands() {
    return {
      "mark-entry":
        (options?: {
          entry?: string;
          subentry?: string;
          bold?: boolean;
          italic?: boolean;
          crossReference?: string;
        }) =>
        ({ state, tr, dispatch }) => {
          let entryText = options?.entry;
          if (!entryText) {
            const { empty, from, to } = state.selection;
            entryText = empty ? "" : state.doc.textBetween(from, to, " ").trim();
          }
          if (!entryText) return false;
          if (options?.subentry) {
            entryText = `${entryText}:${options.subentry}`;
          }
          const parts = [`XE "${entryText.replaceAll('"', "''")}"`];
          if (options?.bold) parts.push("\\b");
          if (options?.italic) parts.push("\\i");
          if (options?.crossReference)
            parts.push(`\\t "${options.crossReference.replaceAll('"', "''")}"`);
          const instruction = parts.join(" ");

          if (dispatch) {
            const seed = {
              type: "inlinePassthrough",
              attrs: {
                data: JSON.stringify({ simpleField: { instruction } }),
              },
            };
            const node = state.schema.nodeFromJSON(seed);
            if (node) {
              tr.insert(state.selection.to, node);
            }
          }
          return true;
        },

      "mark-entry-all":
        (options?: {
          text?: string;
          entry?: string;
          subentry?: string;
          bold?: boolean;
          italic?: boolean;
          crossReference?: string;
          matchCase?: boolean;
        }) =>
        ({ state, tr, dispatch }) => {
          let targetText = options?.text;
          if (!targetText) {
            const { empty, from, to } = state.selection;
            targetText = empty ? "" : state.doc.textBetween(from, to, " ").trim();
          }
          if (!targetText) return false;

          let entryText = options?.entry ?? targetText;
          if (options?.subentry) {
            entryText = `${entryText}:${options.subentry}`;
          }
          const parts = [`XE "${entryText.replaceAll('"', "''")}"`];
          if (options?.bold) parts.push("\\b");
          if (options?.italic) parts.push("\\i");
          if (options?.crossReference)
            parts.push(`\\t "${options.crossReference.replaceAll('"', "''")}"`);
          const instruction = parts.join(" ");

          const matches: number[] = [];
          const matchCase = options?.matchCase ?? false;
          const query = matchCase ? targetText : targetText.toLowerCase();

          state.doc.descendants((node, pos) => {
            if (node.isText && node.text) {
              const textContent = matchCase ? node.text : node.text.toLowerCase();
              let startIndex = 0;
              while (startIndex < textContent.length) {
                const index = textContent.indexOf(query, startIndex);
                if (index === -1) break;
                matches.push(pos + index + targetText.length);
                startIndex = index + query.length;
              }
            }
            return true;
          });

          if (matches.length === 0) return false;

          if (dispatch) {
            const seed = {
              type: "inlinePassthrough",
              attrs: {
                data: JSON.stringify({ simpleField: { instruction } }),
              },
            };
            const node = state.schema.nodeFromJSON(seed);
            if (node) {
              for (const insertPos of matches.reverse()) {
                tr.insert(insertPos, node);
              }
            }
          }
          return true;
        },

      "insert-index":
        (pageOf?: PageOf, tabPositionTw?: number) =>
        ({ state, tr, dispatch }) => {
          const paragraphs = buildIndexParagraphs(
            collectEntries(state.doc, pageOf),
            tabPositionTw,
            state.schema,
          );
          if (paragraphs.length === 0) return false;
          if (dispatch) {
            stampIndexStyles(tr);
            const nodes = paragraphs
              .map((para) => state.schema.nodeFromJSON(para))
              .filter((node): node is PMNode => !!node);
            tr.replaceSelection(new Slice(Fragment.fromArray(nodes), 0, 0)).scrollIntoView();
          }
          return true;
        },

      "update-index":
        (pageOf?: PageOf, tabPositionTw?: number) =>
        ({ state, tr, dispatch }) => {
          const head = findIndexHead(state.doc);
          if (head == null) return false;
          const paragraphs = buildIndexParagraphs(
            collectEntries(state.doc, pageOf),
            tabPositionTw,
            state.schema,
          );
          if (paragraphs.length === 0) return false;
          if (dispatch) {
            stampIndexStyles(tr);
            const nodes = paragraphs
              .map((para) => state.schema.nodeFromJSON(para))
              .filter((node): node is PMNode => !!node);
            const doomed: { from: number; to: number }[] = [];
            state.doc.descendants((node, at) => {
              if (
                node.type.name === "paragraph" &&
                /^Index\d$/.test(String(node.attrs.style ?? ""))
              )
                doomed.push({ from: at, to: at + node.nodeSize });
              return true;
            });
            for (const range of doomed.reverse()) tr.delete(range.from, range.to);
            tr.insert(head, nodes);
          }
          return true;
        },

      "mark-citation":
        (options?: {
          longCitation?: string;
          shortCitation?: string;
          category?: number;
          bold?: boolean;
          italic?: boolean;
        }) =>
        ({ state, tr, dispatch }) => {
          let longCitation = options?.longCitation;
          if (!longCitation) {
            const { empty, from, to } = state.selection;
            longCitation = empty ? "" : state.doc.textBetween(from, to, " ").trim();
          }
          if (!longCitation) return false;

          const shortCitation = options?.shortCitation ?? longCitation;
          const category = options?.category ?? 1;

          const parts = [
            `TA \\l "${longCitation.replaceAll('"', "''")}"`,
            `\\s "${shortCitation.replaceAll('"', "''")}"`,
            `\\c ${category}`,
          ];
          if (options?.bold) parts.push("\\b");
          if (options?.italic) parts.push("\\i");
          const instruction = parts.join(" ");

          if (dispatch) {
            const seed = {
              type: "inlinePassthrough",
              attrs: {
                data: JSON.stringify({ simpleField: { instruction } }),
              },
            };
            const node = state.schema.nodeFromJSON(seed);
            if (node) {
              tr.insert(state.selection.to, node);
            }
          }
          return true;
        },

      "insert-toa":
        (
          pageOf?: PageOf,
          tabPositionTw?: number,
          options?: { category?: number; passim?: boolean },
        ) =>
        ({ state, tr, dispatch }) => {
          const paragraphs = buildToaParagraphs(
            collectCitations(state.doc, pageOf),
            tabPositionTw,
            state.schema,
            options,
          );
          if (paragraphs.length === 0) return false;
          if (dispatch) {
            stampToaStyles(tr);
            const nodes = paragraphs
              .map((para) => state.schema.nodeFromJSON(para))
              .filter((node): node is PMNode => !!node);
            tr.replaceSelection(new Slice(Fragment.fromArray(nodes), 0, 0)).scrollIntoView();
          }
          return true;
        },

      "update-toa":
        (pageOf?: PageOf, tabPositionTw?: number) =>
        ({ state, tr, dispatch }) => {
          const head = findToaHead(state.doc);
          if (head == null) return false;
          const paragraphs = buildToaParagraphs(
            collectCitations(state.doc, pageOf),
            tabPositionTw,
            state.schema,
          );
          if (paragraphs.length === 0) return false;
          if (dispatch) {
            stampToaStyles(tr);
            const nodes = paragraphs
              .map((para) => state.schema.nodeFromJSON(para))
              .filter((node): node is PMNode => !!node);
            const doomed: { from: number; to: number }[] = [];
            state.doc.descendants((node, at) => {
              if (
                node.type.name === "paragraph" &&
                (node.attrs.style === "TOAHeading" || node.attrs.style === "TableOfAuthorities")
              ) {
                doomed.push({ from: at, to: at + node.nodeSize });
              }
              return true;
            });
            for (const range of doomed.reverse()) tr.delete(range.from, range.to);
            tr.insert(head, nodes);
          }
          return true;
        },
    };
  },
});

declare module "@tiptap/core" {
  interface Commands<ReturnType> {
    docenIndexCommands: {
      "mark-entry": (options?: {
        entry?: string;
        subentry?: string;
        bold?: boolean;
        italic?: boolean;
        crossReference?: string;
      }) => ReturnType;
      "mark-entry-all": (options?: {
        text?: string;
        entry?: string;
        subentry?: string;
        bold?: boolean;
        italic?: boolean;
        crossReference?: string;
        matchCase?: boolean;
      }) => ReturnType;
      "insert-index": (pageOf?: PageOf, tabPositionTw?: number) => ReturnType;
      "update-index": (pageOf?: PageOf, tabPositionTw?: number) => ReturnType;
      "mark-citation": (options?: {
        longCitation?: string;
        shortCitation?: string;
        category?: number;
        bold?: boolean;
        italic?: boolean;
      }) => ReturnType;
      "insert-toa": (
        pageOf?: PageOf,
        tabPositionTw?: number,
        options?: { category?: number; passim?: boolean },
      ) => ReturnType;
      "update-toa": (pageOf?: PageOf, tabPositionTw?: number) => ReturnType;
    };
  }
}
