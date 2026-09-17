import type { ParagraphChild, PermStartOptions } from "@office-open/docx";

import { Node } from "../core";
import type { ParseInlineRule } from "./types";
import { attrNative } from "./utils";

type PermStartBranch = Extract<ParagraphChild, { permStart: unknown }>;
type PermEndBranch = Extract<ParagraphChild, { permEnd: unknown }>;

export const parsePermStartInline: ParseInlineRule<PermStartBranch> = {
  match: (child): child is PermStartBranch => "permStart" in child,
  convert: (child) => {
    const ps = child.permStart as PermStartOptions;
    return {
      type: "permStart",
      attrs: {
        id: ps.id != null ? Number(ps.id) : 0,
        editGroup: ps.editGroup ?? "everyone",
        editor: ps.editor ?? null,
        colFirst: ps.colFirst ?? null,
        colLast: ps.colLast ?? null,
      },
    };
  },
};

export const parsePermEndInline: ParseInlineRule<PermEndBranch> = {
  match: (child): child is PermEndBranch => "permEnd" in child,
  convert: (child) => {
    const rawId =
      typeof child.permEnd === "object"
        ? (child.permEnd as Record<string, unknown>).id
        : child.permEnd;
    return {
      type: "permEnd",
      attrs: {
        id: rawId != null ? Number(rawId) : 0,
      },
    };
  },
};

export const PermStart = Node.create({
  name: "permStart",
  group: "inline",
  inline: true,
  atom: true,
  selectable: false,

  addAttributes() {
    return {
      id: attrNative(0),
      editGroup: attrNative("everyone"),
      editor: attrNative(null),
      colFirst: attrNative(null),
      colLast: attrNative(null),
    };
  },

  parseHTML() {
    return [{ tag: "span.docx-perm-start" }];
  },

  renderHTML({ HTMLAttributes }) {
    return ["span", { class: "docx-perm-start", "data-perm-id": HTMLAttributes.id }];
  },

  parseDocxInline: parsePermStartInline,
});

export const PermEnd = Node.create({
  name: "permEnd",
  group: "inline",
  inline: true,
  atom: true,
  selectable: false,

  addAttributes() {
    return {
      id: attrNative(0),
    };
  },

  parseHTML() {
    return [{ tag: "span.docx-perm-end" }];
  },

  renderHTML({ HTMLAttributes }) {
    return ["span", { class: "docx-perm-end", "data-perm-id": HTMLAttributes.id }];
  },

  parseDocxInline: parsePermEndInline,
});
