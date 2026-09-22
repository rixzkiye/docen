import { decodeBase64, encodeBase64 } from "@office-open/core";

import { Node } from "../core";

/** JSON marker wrapping a binary leaf in passthrough data — see
 *  {@link encodePassthroughData}. */
const BINARY_MARKER = "$docenBinary";

/** JSON-safe encoding for a passthrough branch's data. A branch may carry
 *  binary leaves (an OLE `embed.data` Uint8Array, a VML picture's bytes, a
 *  subDoc body) that a plain `JSON.stringify` either corrupts — a typed array
 *  serializes to a `{"0": …}` index object — or drops (an ArrayBuffer becomes
 *  `{}`). Each binary leaf rides as `{ $docenBinary: <base64> }` and
 *  {@link decodePassthroughData} restores the exact bytes for compile, so an
 *  autosave/v-model JSON round-trip no longer degrades the branch. */
export function encodePassthroughData(value: unknown): string {
  return JSON.stringify(value, (_key, leaf: unknown) => {
    if (leaf instanceof Uint8Array) return { [BINARY_MARKER]: encodeBase64(leaf) };
    if (leaf instanceof ArrayBuffer) return { [BINARY_MARKER]: encodeBase64(new Uint8Array(leaf)) };
    if (ArrayBuffer.isView(leaf)) {
      const view = leaf as ArrayBufferView;
      return {
        [BINARY_MARKER]: encodeBase64(
          new Uint8Array(view.buffer, view.byteOffset, view.byteLength),
        ),
      };
    }
    // A Node Buffer reaches the replacer as its toJSON form
    // ({type:"Buffer",data:[…]}), not as the Uint8Array it really is.
    if (
      leaf !== null &&
      typeof leaf === "object" &&
      (leaf as { type?: unknown }).type === "Buffer" &&
      Array.isArray((leaf as { data?: unknown }).data)
    ) {
      return { [BINARY_MARKER]: encodeBase64(Uint8Array.from((leaf as { data: number[] }).data)) };
    }
    return leaf;
  });
}

/** Reverse of {@link encodePassthroughData}: restore the binary leaves a
 *  passthrough branch carries, so compile sees the same model office-open
 *  produced. Data written before the codec (plain JSON) parses unchanged. */
export function decodePassthroughData<T = unknown>(data: string): T {
  return JSON.parse(data, (_key, leaf: unknown) => {
    if (leaf !== null && typeof leaf === "object" && Object.keys(leaf as object).length === 1) {
      const marked = (leaf as Record<string, unknown>)[BINARY_MARKER];
      if (typeof marked === "string") return decodeBase64(marked);
    }
    return leaf;
  }) as T;
}

/**
 * Passthrough — block atom carrying an opaque {@link SectionChild} that has
 * no native Tiptap representation (rawXml, bookmarkStart/End, altChunk,
 * subDoc, customXml).
 *
 * The full SectionChild is stored as JSON in `attrs.data` so the DOCX→JSON→DOCX
 * round-trip stays byte-faithful: office-open's stringify handles the inner
 * structure verbatim. The node is not editable; the canvas paints its
 * placeholder.
 *
 * DOCX serialization is inlined in DocxManager (compile/resolve read/write
 * `attrs.data` directly), so no renderDocx/parseDocx is needed here.
 */
export const Passthrough = Node.create({
  name: "passthrough",
  group: "block",
  atom: true,

  addAttributes() {
    return {
      data: {
        default: "{}",
        rendered: false,
      },
    };
  },

  parseHTML() {
    return [{ tag: "div[data-passthrough]" }];
  },
});

/**
 * InlinePassthrough — inline atom carrying an opaque inline ParagraphChild that
 * has no native Tiptap representation (bookmarkStart/End, comment range markers,
 * proofErr, track-change markers, …). The full ParagraphChild rides in
 * `attrs.data` as JSON so DOCX→JSON→DOCX round-trips byte-faithful; the atom is
 * zero-width (bookmark/range markers carry no layout box), matching Word's
 * non-printing metadata. Mirrors the block-level Passthrough for inline children.
 */
export const InlinePassthrough = Node.create({
  name: "inlinePassthrough",
  group: "inline",
  inline: true,
  atom: true,

  addAttributes() {
    return {
      data: {
        default: "{}",
        rendered: false,
        parseHTML: (element: HTMLElement) =>
          element.getAttribute("data-inline-passthrough") ?? "{}",
      },
      // Citation identity for academic documents: the passthrough atom carries
      // the bibliography key beside its OOXML payload so a citation mark
      // survives DOCX round-trips (the payload alone has no key field).
      citationKey: {
        default: null,
        rendered: false,
        parseHTML: (element: HTMLElement) => element.getAttribute("data-citation-key"),
      },
    };
  },

  parseHTML() {
    return [{ tag: "span[data-inline-passthrough]" }];
  },
});
