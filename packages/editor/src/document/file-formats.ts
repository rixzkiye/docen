// Open/save format tables: filename+MIME detection for open(), per-format
// save-picker metadata, the read-only live command set, and the locally
// handled command set (the "wired" basis for ribbon greying).

import type { DocxVariant } from "@docen/docx";

/** The formats the open picker can actually load: the docx family, Markdown,
 *  RTF, or Plain Text. Flat OPC XML (.xml) is recognized too, but only to be
 *  refused with a clear error — see {@link detectOpenFormat}. */
export type OpenFormat = DocxVariant | "markdown" | "rtf" | "text";

/** Save targets: the docx family plus Markdown, PDF, RTF, HTML, Plain Text, and
 *  OpenDocument Text (ODT). */
export type SaveFormat = DocxVariant | "markdown" | "pdf" | "rtf" | "html" | "txt" | "odt";

/** Clear refusal for Flat OPC input. A Flat OPC package is a single
 *  WordprocessingML XML document (no ZIP), which office-open's archive parser
 *  does not read — the type is recognized only to surface this error instead
 *  of "Unsupported file type". */
export const FLAT_OPC_UNSUPPORTED =
  "Flat OPC XML (.xml) documents are not supported yet — open the .docx or .docm version.";

/** A refusal raised by {@link detectOpenFormat}. `code` lets the host localize
 *  the surface (the English message stays for logs and programmatic callers);
 *  `file` carries the picked file's name for the unsupported-type message. */
export class OpenFormatError extends Error {
  constructor(
    readonly code: "flat-opc" | "unsupported",
    message: string,
    /** The picked file's name (unsupported-type refusals show it). */
    readonly file?: string,
  ) {
    super(message);
    this.name = "OpenFormatError";
  }
}

/** Detect a document's format from its filename + MIME for open(). Extension
 *  first (the picker filters on it), MIME as a fallback for platforms that fill
 *  it in — the MIME's base type only, since browsers append params
 *  (`application/xml;charset=utf-8`). Throws on Flat OPC XML (recognized, but
 *  no parser exists) and on an unrecognized type so the caller surfaces the
 *  error rather than silently parsing garbage. */
export function detectOpenFormat(file: File): OpenFormat {
  const name = file.name.toLowerCase();
  if (name.endsWith(".docx")) return "docx";
  if (name.endsWith(".docm")) return "docm";
  if (name.endsWith(".dotx")) return "dotx";
  if (name.endsWith(".dotm")) return "dotm";
  if (name.endsWith(".md") || name.endsWith(".markdown")) return "markdown";
  if (name.endsWith(".rtf")) return "rtf";
  if (name.endsWith(".txt")) return "text";
  if (name.endsWith(".xml")) throw new OpenFormatError("flat-opc", FLAT_OPC_UNSUPPORTED, name);
  const type = (file.type.split(";")[0] ?? "").trim().toLowerCase();
  if (type.includes("ms-word.document.macroenabled")) return "docm";
  if (type.includes("wordprocessingml.template")) return "dotx";
  if (type.includes("ms-word.template.macroenabled")) return "dotm";
  if (type.includes("wordprocessingml.document")) return "docx";
  if (type === "text/markdown") return "markdown";
  if (type === "application/rtf" || type === "text/rtf") return "rtf";
  if (type === "text/plain") return "text";
  if (type === "application/xml" || type === "text/xml")
    throw new OpenFormatError("flat-opc", FLAT_OPC_UNSUPPORTED, file.name || type);
  throw new OpenFormatError(
    "unsupported",
    `Unsupported file type: ${file.name || type || "(unknown)"}`,
    file.name || type || "(unknown)",
  );
}

/** Per-format metadata for #saveAs: the picker description, the MIME anchoring
 *  its accept filter, and the extension stamped on the suggested name. The MIME
 *  must be a BARE type — showSaveFilePicker rejects accept keys carrying params
 *  (e.g. ";charset=utf-8") with NotSupportedError, so the picker never opens. */
export const SAVE_FORMATS: Record<SaveFormat, { description: string; mime: string; ext: string }> =
  {
    docx: {
      description: "Word Document",
      mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      ext: ".docx",
    },
    docm: {
      description: "Word Macro-Enabled Document",
      mime: "application/vnd.ms-word.document.macroEnabled.12",
      ext: ".docm",
    },
    dotx: {
      description: "Word Template",
      mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
      ext: ".dotx",
    },
    dotm: {
      description: "Word Macro-Enabled Template",
      mime: "application/vnd.ms-word.template.macroEnabled.12",
      ext: ".dotm",
    },
    markdown: { description: "Markdown", mime: "text/markdown", ext: ".md" },
    pdf: { description: "PDF Document", mime: "application/pdf", ext: ".pdf" },
    rtf: { description: "Rich Text Format", mime: "application/rtf", ext: ".rtf" },
    html: { description: "Web Page", mime: "text/html", ext: ".html" },
    txt: { description: "Plain Text", mime: "text/plain", ext: ".txt" },
    odt: {
      description: "OpenDocument Text",
      mime: "application/vnd.oasis.opendocument.text",
      ext: ".odt",
    },
  };

/** Suggested download name for a save: the document's display name with any
 *  known document extension swapped for the target format's (a .docx opened
 *  then saved as a template must not keep its .docx name). */
export function suggestedFileName(name: string, cfg: { ext: string }): string {
  return name.replace(/\.(docx|docm|dotx|dotm|md|markdown|txt|rtf|html|htm|odt)$/i, "") + cfg.ext;
}

/** Commands that stay live when the document is read-only (Viewing mode):
 *  chrome toggles, view panes, the mode switch, save, clipboard reads and
 *  selection — everything else mutates the document and is refused. */
export const READONLY_LIVE: ReadonlySet<string> = new Set([
  "toggle-navigation",
  "zoom",
  "zoom-100",
  "edit-mode",
  "save",
  "copy",
  "select",
  "search",
  "word-count",
  "show-marks",
  "show-comments",
  // Spelling (Review → Spelling & Grammar) — checking is a read operation;
  // the pane works read-only too.
  "spell-check",
  // The Styles pane — a read view over the styles model (Word's read-only
  // mode can still open it); MODIFYING a style stays a live-document action.
  "styles-pane",
  // View-surface toggles — paint-time view state, not document content.
  "toggle-ruler",
  "toggle-gridlines",
  // Field codes (Alt+F9) — a projection toggle, like the view toggles above.
  "toggle-field-codes",
  // The document views — Word's read-only mode can still switch views.
  "print-layout",
  "web-layout",
  "read-mode",
  "draft",
]);

/** Commands handled locally in #onCommand/#onChange (not routed to
 *  editor.commands — they read/write host state the editor can't reach, e.g.
 *  navigation/find/zoom). Together with {@link WIRED_DISPATCH} this is the
 *  "wired" set used to grey out unwired skeleton commands. lang-zh/lang-en are
 *  header menu items, not ribbon commands, so excluded. */
export const LOCAL_HANDLED: ReadonlySet<string> = new Set([
  // #onCommand
  "toggle-navigation",
  "search",
  "replace",
  "page-size",
  "orientation",
  "margins",
  // Columns presets write the current section's w:cols count; Line Numbers
  // toggles the section's w:lnNumType (both via #mutateCurrentSection).
  "columns",
  "line-numbers",
  "hyphenation",
  "insert-soft-hyphen",
  "zoom",
  "zoom-100",
  "save",
  "insert-picture",
  // Home → Clipboard group launcher — the Office Clipboard pane.
  "clipboard-dialog",
  "show-marks",
  // Markdown input mode — a host typing-mode flag (the bridge reads it per
  // keystroke); no editor command behind it.
  "markdown-input",
  "copy",
  "cut",
  "paste",
  "select",
  "format-painter",
  "edit-mode",
  "word-count",
  // Repeat (QAT; F4 lives in the bridge) — the host dispatches the recorded
  // insertion from the editor's storage.
  "repeat",
  // Fill Effects opens the page-picture-fill dialog (Design → Page Color
  // group); the OK path rides the fill-effects:ok event.
  "fill-effects",
  // Picture pixel tools: Compress Pictures opens its dialog (the OK path
  // rides `picture-pixels`); Set Transparent Color arms the canvas
  // eyedropper (the bridge samples the press and the host re-encodes).
  "compress-pictures",
  "picture-transparent-pick",
  // Spelling opens the proofing pane (Review → Spelling & Grammar); the
  // check itself and the pane actions are host state (#runSpellCheck).
  "spell-check",
  // Language opens the proofing-language dialog (Review → Language, the
  // status-bar language item); the commit stamps w:lang on the selection.
  "language",
  // Phonetic guide (拼音指南) opens the per-character reading dialog; the
  // commit splits the selection into per-character ruby runs.
  "phonetic-guide",
  // Chinese Layout (中文版式) opens the two-lines-in-one dialog; the commit
  // stamps the eastAsianLayout combine mark on the selection.
  "two-lines-in-one",
  // The Multilevel List gallery's last entry opens the Define New Multilevel
  // List dialog; the commit registers a document numbering definition.
  "define-new-list",
  // Insert Caption (References → Captions) opens the caption dialog; the
  // commit seeds a Caption-styled paragraph with a SEQ field.
  "insert-caption",
  // Cross-reference (References → Captions) opens its dialog; the commit
  // seeds a cached REF/PAGEREF field pointing at a document bookmark.
  "cross-reference",
  // Index — Mark Entry prompts and seeds an XE field; insert/update collect
  // the XE fields into the Index-styled entry block (commands take pageOf).
  "mark-entry",
  "mark-entry-all",
  "insert-index",
  "update-index",
  // Table of Authorities — Mark Citation seeds a TA field; insert/update TOA
  // collect the TA citations into TOAHeading/TableOfAuthorities paragraphs.
  "mark-citation",
  "insert-toa",
  "update-toa",
  // Citations & Bibliography — the Source Manager dialog in two modes (its
  // sources:ok writes attrs.bibliography; citation:ok seeds a cached CITATION
  // field), and the Bibliography block rebuild.
  "manage-sources",
  "insert-citation",
  "bibliography",
  // TOC insert/update — dispatch with the bridge's pageOf (page numbers live
  // in the canvas caret map, which editor.commands can't reach). Table of
  // Figures is the same field with the \c caption switch.
  "toc",
  "update-toc",
  "toc-dialog",
  "remove-toc",
  "update-toc-page",
  "table-of-figures",
  "update-figures",
  // Mail merge — recipients dialogs, the merge-field seeds, the preview
  // pass, and the Finish & Merge document (attrs.recipients data source).
  "select-recipients",
  "edit-recipients",
  "merge-field",
  "address-block",
  "greeting-line",
  "preview-results",
  "first-record",
  "last-record",
  "start-merge",
  "finish-merge",
  // The Equation context tab's symbol grid drops glyphs at the caret.
  "insert-symbol",
  // Header/footer stories — open through the bridge (the same lifecycle as
  // the band double-click); the Page Number drop seeds a PAGE field.
  "header",
  "footer",
  "page-number",
  // The Header & Footer context tab — switch stories, flip the slot flags
  // (same sectionProperties writes the Insert drop-downs use), and close.
  "goto-header",
  "goto-footer",
  "header-option",
  "close-header-footer",
  // Symbol opens its grid dialog (insertion arrives via symbol:insert);
  // Bookmark prompts for a name and wraps the selection.
  "symbol",
  "bookmark",
  // The Table button's face opens the hover grid; its dropdown's Insert
  // Table opens the classic dialog shape. ("insert-table" itself is wired —
  // the face click is intercepted before the engine dispatch.)
  "table-dialog",
  // Paragraph opens the paragraph dialog prefilled from the caret paragraph
  // (the commit arrives via paragraph:ok → the paragraph-dialog-apply
  // command, which stamps every selected paragraph).
  "paragraph-dialog",
  // Font opens the font dialog prefilled from the selection's run marks
  // (the commit arrives via font:ok, stamped mark-by-mark by #onFontDialogOk).
  "font-dialog",
  // Table Properties opens the table dialog prefilled from the caret table's
  // attrs (the commit arrives via table-properties:ok → table-properties-apply,
  // which rewrites the table's w:jc alignment and w:tblInd indent).
  "table-properties",
  // Drawing Properties (Alt Text) opens the size-and-position dialog for the
  // selected drawing (inline pictures are declined in #onCommand — no state);
  // Crop enters the crop-mode gesture. Shared by the context menu and the
  // Picture Format context tab's Accessibility group / Size split.
  "drawing-properties",
  "drawing-crop",
  // Footnote prompts for the note text, references the caret and appends the
  // note body to documentExtras.footnotes. Equation drops a placeholder
  // math template (the gallery's fraction/script/radical/sum/integral) at
  // the caret.
  "insert-footnote",
  // Field opens the field dialog (Insert → Text group; the commit arrives via
  // field:ok). The context menu's Update/Edit Field act on the atom under the
  // caret (update = Word's F9), Update All Fields walks every field atom, and
  // toggle-field-codes is Alt+F9's projection switch; the checkbox flip is the
  // form-field variant.
  "insert-field",
  "update-field",
  "edit-field",
  "update-all-fields",
  "toggle-field-codes",
  "toggle-field-checkbox",
  // Chart Design's Edit Data opens the data-grid dialog (the commit arrives
  // via chart:ok → the chart-data-apply command).
  "chart-edit-data",
  "equation",
  // Page Color writes the doc-level w:background (doc.attrs.background) from
  // the color-picker's palette value. Page Borders stamps a w:pgBorders
  // preset (none/box/shadow/double/dashed) on the current section. Watermark
  // stamps/removes the preset header shape (every slot, behind-doc).
  // Paragraph Spacing stamps the styles' docDefaults paragraph spacing.
  "page-color",
  "page-border",
  "watermark",
  "paragraph-spacing",
  // View-surface toggles (Word's View → Ruler / Gridlines).
  "toggle-ruler",
  "toggle-gridlines",
  // Link opens the hyperlink dialog and marks the selection (Word's Insert
  // Link); the context menu's Open/Copy Hyperlink and Remove Hyperlink act on
  // the clicked link.
  "link",
  "unset-link",
  "open-link",
  "copy-link",
  "open-embedded-object",
  "download-embedded-object",
  // New Comment anchors the selection with a Word comment (range markers +
  // a documentExtras.comments entry) — composed in the comments pane, not a
  // prompt; Edit opens the pane (cards edit inline), Delete removes the
  // comment covering the selection, Previous/Next step through the ranges.
  // Show Comments toggles the pane (Word's Review → Show Comments).
  "new-comment",
  "comment",
  "edit-comment",
  "delete-comment",
  "previous-comment",
  "next-comment",
  "show-comments",
  // Review → Reviewing Pane toggles the revisions pane (the card actions
  // dispatch the engine's by-id accept/reject commands).
  "reviewing-pane",
  // Word's Display for Review: the four markup views re-project the document
  // (the host carries the state and passes it to the layout projection);
  // Specific People scopes the view to one reviewer's revisions, Markup
  // Colors swaps the revision palette (By author / By change type). The two
  // "…All Changes Shown" sweeps need the same filter as their argument.
  "display-for-review",
  "review-specific-people",
  "markup-colors",
  "show-markup",
  "accept-all-changes-shown",
  "reject-all-changes-shown",
  // Table Design → Draw Border: the pen pickers stamp the host's pen state;
  // the painter split arms the paint/erase mode (the sweep itself rides the
  // wired paint-cell-border / erase-cell-border commands).
  "pen-style",
  "pen-size",
  "pen-color",
  "border-painter",
  // Text Box / Shapes insert a standalone wps shape run (Shapes reads the
  // gallery preset from the split item's value).
  "text-box",
  "shapes",
  // Insert → Pages menu extras: a cover block at the document start, or two
  // page breaks (Word's blank half-page + full page). WordArt inserts a
  // preset-styled text box; Date & Time opens its dialog (static text or a
  // DATE field); Object's "Text from File" reads a .txt at the caret.
  "cover-page",
  "blank-page",
  "wordart",
  "date-time",
  "object",
  "insert-file-text",
  // View → the four document views write the `view` attribute (Print Layout /
  // Web Layout / Read Mode / Draft); Outline opens the document-structure
  // pane (Word's outline view maps to the navigation pane here).
  "print-layout",
  "web-layout",
  "read-mode",
  "draft",
  "outline",
  // #onChange (data-event)
  "open",
  "save-as",
  "save-as-template",
  "save-as-markdown",
  "save-as-rtf",
  "save-as-html",
  "save-as-txt",
  "save-as-odt",
  "new-from-template",
  "save-as-pdf",
  "print",
  // The filename menu's tail (Word's File menu): Properties opens the
  // properties task pane, Inspect Document its findings dialog, Share goes
  // through the Web Share API (clipboard fallback), Close resets the
  // document (all host-handled via docen:close first).
  "properties",
  "inspect-document",
  "share",
  "close",
]);
