import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { DocAttrStep } from "@tiptap/pm/transform";

import { t } from "../../ui/i18n/localize";

/** A bibliography source as the sources dialog edits it — office-open's
 *  SourceTypeOptions narrowed to the exposed fields (the document's own
 *  sources round-trip untouched through the open attrs value). */
export type BibliographySource = {
  tag?: string;
  sourceType?: string;
  title?: string;
  year?: string;
  publisher?: string;
  author?: { authors?: { last?: string; first?: string; corporate?: string }[] };
};

/** Formats an in-text citation string per active citation style (APA, MLA, Chicago, IEEE). */
export function formatInTextCitation(source: BibliographySource, style = "APA", index = 1): string {
  const normStyle = style.toUpperCase();
  if (normStyle === "IEEE") {
    return `[${index}]`;
  }
  const authors = source.author?.authors ?? [];
  let authorStr = "";
  if (authors.length > 0) {
    if (authors[0]?.corporate) {
      authorStr = authors[0].corporate;
    } else {
      const lasts = authors.map((a) => a.last || a.first || "").filter(Boolean);
      if (lasts.length === 1) {
        authorStr = lasts[0]!;
      } else if (lasts.length === 2) {
        const joiner = normStyle === "APA" ? " & " : " and ";
        authorStr = `${lasts[0]}${joiner}${lasts[1]}`;
      } else if (lasts.length > 2) {
        authorStr = `${lasts[0]} et al.`;
      }
    }
  }
  const head = authorStr || source.title || source.tag || "Unknown";
  const year = source.year;

  if (normStyle === "MLA") {
    return `(${head})`;
  }
  if (normStyle === "CHICAGO") {
    return `(${head}${year ? ` ${year}` : ""})`;
  }
  // Default APA 7th
  return `(${head}${year ? `, ${year}` : ""})`;
}

/** Formats a full bibliography entry line per active citation style (APA, MLA, Chicago, IEEE). */
export function formatBibliographyEntry(
  source: BibliographySource,
  style = "APA",
  index = 1,
): string {
  const normStyle = style.toUpperCase();
  const authors = source.author?.authors ?? [];
  const year = source.year;
  const title = source.title;
  const publisher = source.publisher;

  if (normStyle === "IEEE") {
    const names = authors
      .map((a) => {
        if (a.corporate) return a.corporate;
        const initial = a.first ? `${a.first[0]}. ` : "";
        return `${initial}${a.last ?? ""}`.trim();
      })
      .filter(Boolean);
    const authorText = names.length > 2 ? `${names[0]} et al.` : names.join(" and ");
    const parts = [
      `[${index}]`,
      authorText ? `${authorText},` : "",
      title ? `${title}.` : "",
      [publisher, year].filter(Boolean).join(", ") + (publisher || year ? "." : ""),
    ].filter(Boolean);
    return parts.join(" ");
  }

  if (normStyle === "MLA") {
    const names = authors
      .map((a, i) => {
        if (a.corporate) return a.corporate;
        if (i === 0) return [a.last, a.first].filter(Boolean).join(", ");
        return [a.first, a.last].filter(Boolean).join(" ");
      })
      .filter(Boolean);
    const authorText = names.length > 2 ? `${names[0]}, et al.` : names.join(" and ");
    const parts = [
      authorText ? `${authorText}.` : "",
      title ? `${title}.` : "",
      [publisher, year].filter(Boolean).join(", ") + (publisher || year ? "." : ""),
    ].filter(Boolean);
    return parts.join(" ");
  }

  if (normStyle === "CHICAGO") {
    const names = authors
      .map((a, i) => {
        if (a.corporate) return a.corporate;
        if (i === 0) return [a.last, a.first].filter(Boolean).join(", ");
        return [a.first, a.last].filter(Boolean).join(" ");
      })
      .filter(Boolean);
    const authorText = names.length > 2 ? `${names[0]}, et al.` : names.join(" and ");
    const parts = [
      authorText ? `${authorText}.` : "",
      year ? `${year}.` : "",
      title ? `${title}.` : "",
      publisher ? `${publisher}.` : "",
    ].filter(Boolean);
    return parts.join(" ");
  }

  // Default APA 7th: Authors (Year). Title. Publisher.
  const names = authors
    .map((a) => {
      if (a.corporate) return a.corporate;
      const initial = a.first ? ` ${a.first[0]}.` : "";
      return `${a.last ?? ""},${initial}`.trim();
    })
    .filter(Boolean);
  let authorText = "";
  if (names.length === 1) authorText = names[0]!;
  else if (names.length === 2) authorText = `${names[0]} & ${names[1]}`;
  else if (names.length > 2)
    authorText = `${names.slice(0, -1).join(", ")}, & ${names[names.length - 1]}`;
  else if (authors[0]?.corporate) authorText = authors[0].corporate;

  const parts = [
    authorText ? `${authorText}` : "",
    year ? `(${year}).` : "",
    title ? `${title}.` : "",
    publisher ? `${publisher}.` : "",
  ].filter(Boolean);
  return parts.join(" ");
}

/** The references commands' view of the host — resolved per call so the
 *  controller can be built before a document opens (the editor and the story
 *  bridge both arrive later). */
export interface ReferencesHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — citations target the active story (header/footer
   *  stories included). */
  bridge(): { activeEditor(): Editor; focus(): void } | undefined;
  /** The host element — the i18n language source for prompts. */
  element(): HTMLElement;
  /** Host callback to update all fields after style changes. */
  updateAllFields?(): void;
}

/**
 * The References tab's citation/bibliography commands, split out of the host
 * element: the source master list (doc.attrs.bibliography), the CITATION
 * field seed, and the Bibliography-styled block rebuild.
 */
export class ReferencesCommands {
  constructor(private readonly host: ReferencesHost) {}

  #target(): Editor | null | undefined {
    return this.host.bridge()?.activeEditor() ?? this.host.editor();
  }

  /** Mark Entry — prompt for the entry text (defaulting to the selection) and
   *  seed an `XE "…"` field at the selection's end. A cached-less fldSimple
   *  renders nothing (Word's invisible index marker) but round-trips verbatim
   *  through DOCX for Insert Index to collect. */
  markIndexEntry(target: Editor): void {
    const { empty, from, to } = target.state.selection;
    const selected = empty ? "" : target.state.doc.textBetween(from, to, " ");
    const entry = window.prompt(t("index.prompt", this.host.element()), selected)?.trim();
    if (!entry) return;
    const seed: JSONContent = {
      type: "inlinePassthrough",
      attrs: {
        data: JSON.stringify({
          simpleField: { instruction: `XE "${entry.replaceAll('"', "''")}"` },
        }),
      },
    } as JSONContent;
    target.view.dispatch(target.state.tr.insert(to, target.schema.nodeFromJSON(seed)));
    this.host.bridge()?.focus();
  }

  /** Mark Citation — prompt for the citation text (defaulting to the selection) and
   *  seed a `TA \l "…" \s "…" \c 1` field at the selection's end. */
  markCitation(target: Editor): void {
    const { empty, from, to } = target.state.selection;
    const selected = empty ? "" : target.state.doc.textBetween(from, to, " ");
    const citation = window.prompt(t("toa.prompt", this.host.element()), selected)?.trim();
    if (!citation) return;
    const seed: JSONContent = {
      type: "inlinePassthrough",
      attrs: {
        data: JSON.stringify({
          simpleField: {
            instruction: `TA \\l "${citation.replaceAll('"', "''")}" \\s "${citation.replaceAll('"', "''")}" \\c 1`,
          },
        }),
      },
    } as JSONContent;
    target.view.dispatch(target.state.tr.insert(to, target.schema.nodeFromJSON(seed)));
    this.host.bridge()?.focus();
  }

  /** The document's bibliography sources — doc.attrs.bibliography (the Source
   *  Manager's master list, word/bibliography.xml on save). */
  bibliographySources(): BibliographySource[] {
    const attrs = this.host.editor()?.state.doc.attrs as {
      bibliography?: { sources?: BibliographySource[]; style?: string } | null;
    };
    return [...(attrs.bibliography?.sources ?? [])];
  }

  /** The document's active citation style ("APA" | "MLA" | "Chicago" | "IEEE"). */
  citationStyle(): string {
    const attrs = this.host.editor()?.state.doc.attrs as {
      bibliography?: { sources?: BibliographySource[]; style?: string } | null;
    };
    return attrs.bibliography?.style || "APA";
  }

  /** Set active citation style and re-evaluate fields in the document. */
  setCitationStyle(style: string): void {
    const target = this.#target();
    if (!target) return;
    const attrs = target.state.doc.attrs as {
      bibliography?: { sources?: BibliographySource[]; style?: string } | null;
    };
    const current = attrs.bibliography ?? {};
    target.commands.command(({ tr }) => {
      tr.step(new DocAttrStep("bibliography", { ...current, style }));
      return true;
    });
    this.host.updateAllFields?.();
  }

  /** Sources dialog commit — replace the document's source list. An empty
   *  list clears the attr (null) so no empty part is emitted. */
  readonly onSourcesOk = (event: Event): void => {
    const { sources } = (event as CustomEvent<{ sources?: BibliographySource[] }>).detail ?? {};
    const target = this.#target();
    if (!target || !sources) return;
    const current = (target.state.doc.attrs as { bibliography?: { style?: string } }).bibliography;
    target.commands.command(({ tr }) => {
      tr.step(
        new DocAttrStep(
          "bibliography",
          sources.length > 0 || current?.style ? { ...current, sources } : null,
        ),
      );
      return true;
    });
  };

  /** Citation dialog insert — seed a cached CITATION field at the caret formatted
   *  per active citation style. */
  readonly onCitationOk = (event: Event): void => {
    const { tag } = (event as CustomEvent<{ tag?: string }>).detail ?? {};
    const target = this.#target();
    if (!target || !tag) return;
    const sources = this.bibliographySources();
    const sourceIndex = sources.findIndex((entry) => entry.tag === tag);
    const source = sourceIndex >= 0 ? sources[sourceIndex]! : { tag };
    const style = this.citationStyle();
    const cached = formatInTextCitation(source, style, sourceIndex >= 0 ? sourceIndex + 1 : 1);
    const seed: JSONContent = {
      type: "inlinePassthrough",
      attrs: {
        data: JSON.stringify({
          simpleField: { instruction: `CITATION "${tag}" \\l 1033`, cachedValue: cached },
        }),
      },
    } as JSONContent;
    const { from } = target.state.selection;
    target.view.dispatch(target.state.tr.insert(from, target.schema.nodeFromJSON(seed)));
    this.host.bridge()?.focus();
  };

  /** Bibliography — rebuild the Bibliography-styled block after the caret's
   *  paragraph from the document's sources (an existing block is replaced,
   *  not duplicated). The two style definitions join the document styles when
   *  absent. */
  insertBibliography(): void {
    const target = this.#target();
    if (!target) return;
    const sources = this.bibliographySources();
    if (sources.length === 0) {
      window.alert(t("bibliography.empty", this.host.element()));
      return;
    }
    const { state } = target;
    if (state.selection.$from.parent.type.name !== "paragraph") return;
    const style = this.citationStyle();
    const block: JSONContent[] = [
      {
        type: "paragraph",
        attrs: { style: "BibliographyHeading" },
        content: [{ type: "text", text: t("bibliography.heading", this.host.element()) }],
      },
      ...sources.map((source, i) => ({
        type: "paragraph",
        attrs: { style: "Bibliography" },
        content: [{ type: "text", text: formatBibliographyEntry(source, style, i + 1) }],
      })),
    ];
    const styles = { ...((state.doc.attrs.styles ?? {}) as Record<string, unknown>) };
    const paragraphStyles = (styles.paragraphStyles ?? []) as { id?: string }[];
    const missing = (["BibliographyHeading", "Bibliography"] as const).filter(
      (id) => !paragraphStyles.some((style) => style.id === id),
    );
    const afterCaret = state.selection.$from.after(state.selection.$from.depth);
    target
      .chain()
      .command(({ tr }) => {
        if (missing.length > 0) {
          const definitions = missing.map((id) => ({
            id,
            name: id === "Bibliography" ? "bibliography" : "Bibliography Heading",
            basedOn: "Normal",
            next: "Normal",
            ...(id === "BibliographyHeading" ? { bold: true } : {}),
          }));
          tr.step(
            new DocAttrStep("styles", {
              ...styles,
              paragraphStyles: [...paragraphStyles, ...definitions],
            }),
          );
        }
        // Drop the stale block first; map the insertion anchor through the
        // deletions so the fresh block lands after the caret's paragraph
        // even when the old block sat before the caret.
        state.doc.descendants((node, at) => {
          if (node.type.name !== "paragraph") return true;
          const raw = (node.attrs as Record<string, unknown>).style;
          const style = typeof raw === "string" ? raw : "";
          if (style === "Bibliography" || style === "BibliographyHeading")
            tr.delete(at, at + node.nodeSize);
          return true;
        });
        tr.insert(
          tr.mapping.map(afterCaret),
          block.map((node) => target.schema.nodeFromJSON(node)),
        );
        return true;
      })
      .run();
    this.host.bridge()?.focus();
  }
}
