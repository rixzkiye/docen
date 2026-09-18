import type { MarkupRangeOptions, MoveRangeStartOptions, ParagraphChild } from "@office-open/docx";

import { Node } from "../core";
import type { ParseInlineRule } from "./types";
import { attrNative } from "./utils";

type MoveFromRangeStartBranch = Extract<ParagraphChild, { moveFromRangeStart: unknown }>;
type MoveFromRangeEndBranch = Extract<ParagraphChild, { moveFromRangeEnd: unknown }>;
type MoveToRangeStartBranch = Extract<ParagraphChild, { moveToRangeStart: unknown }>;
type MoveToRangeEndBranch = Extract<ParagraphChild, { moveToRangeEnd: unknown }>;

export const parseMoveFromRangeStartInline: ParseInlineRule<MoveFromRangeStartBranch> = {
  match: (child): child is MoveFromRangeStartBranch => "moveFromRangeStart" in child,
  convert: (child) => {
    const m = child.moveFromRangeStart as MoveRangeStartOptions;
    return {
      type: "moveFromRangeStart",
      attrs: {
        id: m.id != null ? Number(m.id) : 0,
        name: m.name ?? null,
        author: m.author ?? null,
        date: m.date ?? null,
        displacedByCustomXml: m.displacedByCustomXml ?? null,
        colFirst: m.colFirst ?? null,
        colLast: m.colLast ?? null,
      },
    };
  },
};

export const parseMoveFromRangeEndInline: ParseInlineRule<MoveFromRangeEndBranch> = {
  match: (child): child is MoveFromRangeEndBranch => "moveFromRangeEnd" in child,
  convert: (child) => {
    const rawId =
      typeof child.moveFromRangeEnd === "object" && child.moveFromRangeEnd !== null
        ? (child.moveFromRangeEnd as MarkupRangeOptions).id
        : child.moveFromRangeEnd;
    return {
      type: "moveFromRangeEnd",
      attrs: {
        id: rawId != null ? Number(rawId) : 0,
        displacedByCustomXml:
          typeof child.moveFromRangeEnd === "object" && child.moveFromRangeEnd !== null
            ? ((child.moveFromRangeEnd as MarkupRangeOptions).displacedByCustomXml ?? null)
            : null,
      },
    };
  },
};

export const parseMoveToRangeStartInline: ParseInlineRule<MoveToRangeStartBranch> = {
  match: (child): child is MoveToRangeStartBranch => "moveToRangeStart" in child,
  convert: (child) => {
    const m = child.moveToRangeStart as MoveRangeStartOptions;
    return {
      type: "moveToRangeStart",
      attrs: {
        id: m.id != null ? Number(m.id) : 0,
        name: m.name ?? null,
        author: m.author ?? null,
        date: m.date ?? null,
        displacedByCustomXml: m.displacedByCustomXml ?? null,
        colFirst: m.colFirst ?? null,
        colLast: m.colLast ?? null,
      },
    };
  },
};

export const parseMoveToRangeEndInline: ParseInlineRule<MoveToRangeEndBranch> = {
  match: (child): child is MoveToRangeEndBranch => "moveToRangeEnd" in child,
  convert: (child) => {
    const rawId =
      typeof child.moveToRangeEnd === "object" && child.moveToRangeEnd !== null
        ? (child.moveToRangeEnd as MarkupRangeOptions).id
        : child.moveToRangeEnd;
    return {
      type: "moveToRangeEnd",
      attrs: {
        id: rawId != null ? Number(rawId) : 0,
        displacedByCustomXml:
          typeof child.moveToRangeEnd === "object" && child.moveToRangeEnd !== null
            ? ((child.moveToRangeEnd as MarkupRangeOptions).displacedByCustomXml ?? null)
            : null,
      },
    };
  },
};

export const MoveFromRangeStart = Node.create({
  name: "moveFromRangeStart",
  group: "inline",
  inline: true,
  atom: true,
  selectable: false,

  addAttributes() {
    return {
      id: attrNative(0),
      name: attrNative(null),
      author: attrNative(null),
      date: attrNative(null),
      displacedByCustomXml: attrNative(null),
      colFirst: attrNative(null),
      colLast: attrNative(null),
    };
  },

  parseHTML() {
    return [{ tag: "span.docx-move-from-range-start" }];
  },

  renderHTML({ HTMLAttributes }) {
    return ["span", { class: "docx-move-from-range-start", "data-move-id": HTMLAttributes.id }];
  },

  parseDocxInline: parseMoveFromRangeStartInline,
});

export const MoveFromRangeEnd = Node.create({
  name: "moveFromRangeEnd",
  group: "inline",
  inline: true,
  atom: true,
  selectable: false,

  addAttributes() {
    return {
      id: attrNative(0),
      displacedByCustomXml: attrNative(null),
    };
  },

  parseHTML() {
    return [{ tag: "span.docx-move-from-range-end" }];
  },

  renderHTML({ HTMLAttributes }) {
    return ["span", { class: "docx-move-from-range-end", "data-move-id": HTMLAttributes.id }];
  },

  parseDocxInline: parseMoveFromRangeEndInline,
});

export const MoveToRangeStart = Node.create({
  name: "moveToRangeStart",
  group: "inline",
  inline: true,
  atom: true,
  selectable: false,

  addAttributes() {
    return {
      id: attrNative(0),
      name: attrNative(null),
      author: attrNative(null),
      date: attrNative(null),
      displacedByCustomXml: attrNative(null),
      colFirst: attrNative(null),
      colLast: attrNative(null),
    };
  },

  parseHTML() {
    return [{ tag: "span.docx-move-to-range-start" }];
  },

  renderHTML({ HTMLAttributes }) {
    return ["span", { class: "docx-move-to-range-start", "data-move-id": HTMLAttributes.id }];
  },

  parseDocxInline: parseMoveToRangeStartInline,
});

export const MoveToRangeEnd = Node.create({
  name: "moveToRangeEnd",
  group: "inline",
  inline: true,
  atom: true,
  selectable: false,

  addAttributes() {
    return {
      id: attrNative(0),
      displacedByCustomXml: attrNative(null),
    };
  },

  parseHTML() {
    return [{ tag: "span.docx-move-to-range-end" }];
  },

  renderHTML({ HTMLAttributes }) {
    return ["span", { class: "docx-move-to-range-end", "data-move-id": HTMLAttributes.id }];
  },

  parseDocxInline: parseMoveToRangeEndInline,
});
