import { Node as TiptapNode } from "@tiptap/core";

import type { AnyExtension } from "../core";
import { Chart } from "./chart";
import { ColumnBreak } from "./column-break";
import { Document } from "./document";
import { Model3D, Ink } from "./drawing-3d-ink";
import { FormField } from "./form-field";
import { Image } from "./image";
import { Link } from "./link";
import {
  Bold,
  Code,
  Highlight,
  Italic,
  Strike,
  Subscript,
  Superscript,
  Underline,
  SoftHyphen,
  Dir,
  Bdo,
} from "./marks";
import { MathInline } from "./math";
import {
  MoveFromRangeStart,
  MoveFromRangeEnd,
  MoveToRangeStart,
  MoveToRangeEnd,
} from "./move-range";
import { PageBreak } from "./page-break";
import { Paragraph } from "./paragraph";
import { Passthrough, InlinePassthrough } from "./passthrough";
import { PermStart, PermEnd } from "./perm-range";
import { Ruby } from "./ruby";
import { SdtBlock, SdtInline } from "./sdt";
import { SectionBreak } from "./section-break";
import { Tab } from "./tab";
import { Table } from "./table";
import { TableCell } from "./table-cell";
import { TableRow } from "./table-row";
import { TextStyle } from "./text-style";
import { Textbox } from "./textbox";
import { TocField } from "./toc-field";
import { FormatChange, Insertion, Deletion, MoveFrom, MoveTo } from "./track-change";
import { WpgGroup } from "./wpg-group";
import { WpsShape } from "./wps-shape";

// Nodes
/** The inline text atom — plain `Node.create` (the upstream extension added
 *  nothing we use). */
const Text = TiptapNode.create({ name: "text", group: "inline" });

/** Hard line break (<w:br/>) — `Node.create` inline atom. The upstream
 *  extension's keymap/input rules never fire in the viewless canvas route
 *  (typing goes through the textarea bridge). */
const HardBreak = TiptapNode.create({
  name: "hardBreak",
  inline: true,
  group: "inline",
  selectable: false,
  addAttributes() {
    return {
      variant: {
        default: "textWrapping",
        parseHTML: (el) => el.getAttribute("data-variant") || "textWrapping",
        renderHTML: (attrs) =>
          attrs.variant && attrs.variant !== "textWrapping"
            ? { "data-variant": attrs.variant }
            : {},
      },
    };
  },
  parseHTML() {
    return [{ tag: "br" }];
  },
});

export const tiptapNodeExtensions: AnyExtension[] = [
  Document,
  Paragraph,
  Text,
  HardBreak,
  PageBreak,
  ColumnBreak,
  Tab,
  SectionBreak,
  Passthrough,
  InlinePassthrough,
  TocField,
  MathInline,
  SdtBlock,
  SdtInline,
  Textbox,
  Chart,
  Image,
  Model3D,
  Ink,
  WpgGroup,
  WpsShape,
  Table,
  TableRow,
  TableCell,
  FormField,
  PermStart,
  PermEnd,
  MoveFromRangeStart,
  MoveFromRangeEnd,
  MoveToRangeStart,
  MoveToRangeEnd,
];

// Marks
export const tiptapMarkExtensions: AnyExtension[] = [
  Bold,
  Code,
  Deletion,
  FormatChange,
  Highlight,
  Insertion,
  Italic,
  Link,
  MoveFrom,
  MoveTo,
  Ruby,
  Strike,
  Subscript,
  Superscript,
  TextStyle,
  Underline,
  SoftHyphen,
  Dir,
  Bdo,
];

// DOCX schema + DOCX-specific extensions. Editing-behavior extensions
// (UndoRedo/Dropcursor/Gapcursor/TrailingNode/ListKeymap/CharacterCount/Focus)
// live in @docen/editor — the engine stays free of editing-UX concerns.
// The markdown converter and the HTML paste parser use this array as schema;
// those extensions add no schema, so omitting them does not affect conversion.
export const docxExtensions: AnyExtension[] = [...tiptapNodeExtensions, ...tiptapMarkExtensions];

// Export all individual extensions for direct imports from @docen/docx.
// Re-export explicitly (no `export *`) so the public surface is visible.
// Customized extensions export their local version; upstream-only ones re-export
// from @tiptap/* directly, base marks (with DOCX hooks) from ./marks.
export {
  Bold,
  Code,
  Highlight,
  Italic,
  Strike,
  Subscript,
  Superscript,
  Underline,
  SoftHyphen,
  Dir,
  Bdo,
} from "./marks";
export { Document, createDocument } from "./document";
export { Paragraph } from "./paragraph";
export { detectHeadingLevel, HEADING_COMPILE_MAP } from "./paragraph";
// Flat list model: generated numbering references + level builders shared by
// compile (definition registration) and the editor list commands.
export {
  BULLET_GLYPHS,
  BULLET_REFERENCE,
  MULTILEVEL_PRESETS,
  ORDERED_FORMATS,
  ORDERED_REFERENCE_PREFIX,
  buildCustomMultilevelLevels,
  buildListLevels,
  nextMultilevelReference,
  nextOrderedReference,
} from "./list-numbering";
export { ColumnBreak } from "./column-break";
export { SectionBreak } from "./section-break";
export { Table } from "./table";
export { TableRow } from "./table-row";
export { TableCell } from "./table-cell";
export { Chart } from "./chart";
export { Image } from "./image";
export { Link } from "./link";
export { Ruby } from "./ruby";
export { TextStyle } from "./text-style";
export {
  FormatChange,
  Insertion,
  Deletion,
  MoveFrom,
  MoveTo,
  parseFormatRecords,
  parseRunMarks,
} from "./track-change";
export type { RunFormatEdit, RunFormatRecord } from "./track-change";
export {
  MoveFromRangeStart,
  MoveFromRangeEnd,
  MoveToRangeStart,
  MoveToRangeEnd,
} from "./move-range";
export { PageBreak } from "./page-break";
export { WpgGroup } from "./wpg-group";
export { WpsShape } from "./wps-shape";
export { Passthrough, InlinePassthrough } from "./passthrough";
export { TocField } from "./toc-field";
export { SdtBlock, SdtInline } from "./sdt";
export { MathInline, convertLinearToOMML, convertOMMLToLinear } from "./math";
export { Textbox } from "./textbox";
export {
  type DrawingShapeLayout,
  type ShapeWrapping,
  ptToTwip,
  twipToPt,
  emuToTwip,
  twipToEmu,
  ptToEmu,
  emuToPt,
  parseLengthToTwips,
  normalizeVmlShapeStyle,
  parseVmlShapeLayout,
  stringifyVmlShapeLayout,
} from "./drawing-shape-layout";
export { Tab } from "./tab";
export { FormField, extractFormFieldText } from "./form-field";
export { PermStart, PermEnd } from "./perm-range";
export { Model3D, Ink, Drawing3D, DrawingInk, DELEGATION_NOTICE } from "./drawing-3d-ink";
