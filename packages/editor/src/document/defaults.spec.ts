// @vitest-environment happy-dom
// Design → Set as Default: the persisted pair (theme + style set) and its
// re-application to a newly created document. The host element wires the
// command; the storage/apply logic here is driven with a fake storage and a
// real Editor so the behavior is asserted on the document model, not on a
// re-implementation.
import { docxExtensions, normalizeDocument } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import {
  DOCUMENT_DEFAULTS_STORAGE_KEY,
  applyDocumentDefaults,
  clearDocumentDefaults,
  documentDefaultsOf,
  newDocumentJSON,
  parseDocumentDefaults,
  readDocumentDefaults,
  writeDocumentDefaults,
  type DefaultsStorage,
} from "./defaults";
import { DocumentCommands } from "./extensions/commands";

function fakeStorage(seed: Record<string, string> = {}): DefaultsStorage & {
  data: Record<string, string>;
} {
  const data = { ...seed };
  return {
    data,
    getItem: (key) => data[key] ?? null,
    setItem: (key, value) => {
      data[key] = value;
    },
    removeItem: (key) => {
      delete data[key];
    },
  };
}

function makeEditor(json: unknown): Editor {
  return new Editor({
    element: null,
    extensions: [...docxExtensions, DocumentCommands],
    content: json as never,
  });
}

describe("parseDocumentDefaults", () => {
  it("round-trips a valid defaults document", () => {
    expect(
      parseDocumentDefaults(
        JSON.stringify({ version: 1, theme: "facet", themeKind: "theme", styleSet: "classic" }),
      ),
    ).toEqual({ theme: "facet", themeKind: "theme", styleSet: "classic" });
  });

  it("rejects corrupt JSON, foreign versions, and empty themes", () => {
    expect(parseDocumentDefaults(null)).toBeNull();
    expect(parseDocumentDefaults("{")).toBeNull();
    expect(parseDocumentDefaults(JSON.stringify({ version: 2, theme: "facet" }))).toBeNull();
    expect(parseDocumentDefaults(JSON.stringify({ version: 1, theme: "  " }))).toBeNull();
    expect(parseDocumentDefaults(JSON.stringify({ version: 1, theme: 7 }))).toBeNull();
  });
});

describe("document defaults storage", () => {
  it("writes, reads, and clears under the versioned key", () => {
    const storage = fakeStorage();
    writeDocumentDefaults({ theme: "ion", themeKind: "theme-color", styleSet: "modern" }, storage);
    expect(readDocumentDefaults(storage)).toEqual({
      theme: "ion",
      themeKind: "theme-color",
      styleSet: "modern",
    });
    // The stored document is versioned.
    expect(JSON.parse(storage.data[DOCUMENT_DEFAULTS_STORAGE_KEY]!)).toMatchObject({ version: 1 });
    clearDocumentDefaults(storage);
    expect(readDocumentDefaults(storage)).toBeNull();
  });

  it("falls back to factory defaults on corrupt storage", () => {
    const storage = fakeStorage({ [DOCUMENT_DEFAULTS_STORAGE_KEY]: "not json" });
    expect(readDocumentDefaults(storage)).toBeNull();
  });

  it("survives a throwing storage (private mode)", () => {
    const storage: DefaultsStorage = {
      getItem: () => {
        throw new Error("blocked");
      },
      setItem: () => {
        throw new Error("blocked");
      },
      removeItem: () => {
        throw new Error("blocked");
      },
    };
    expect(readDocumentDefaults(storage)).toBeNull();
    expect(() =>
      writeDocumentDefaults({ theme: "office", themeKind: "theme", styleSet: null }, storage),
    ).not.toThrow();
    expect(() => clearDocumentDefaults(storage)).not.toThrow();
  });
});

describe("Set as Default chain", () => {
  it("reads the current document's theme + style set", () => {
    expect(
      documentDefaultsOf({
        documentExtras: {
          settings: { theme: { id: "facet", kind: "theme-font" }, styleSet: "classic" },
        },
      }),
    ).toEqual({ theme: "facet", themeKind: "theme-font", styleSet: "classic" });
    // A document that never applied either reads as the factory state.
    expect(documentDefaultsOf(undefined)).toEqual({
      theme: "office",
      themeKind: "theme",
      styleSet: null,
    });
  });

  it("applies the stored style set to a fresh document (observable on the model)", () => {
    const editor = makeEditor(normalizeDocument(newDocumentJSON(null)));
    const before = (
      editor.state.doc.attrs as {
        styles?: { default?: { document?: { run?: { font?: string } } } };
      }
    ).styles?.default?.document?.run?.font;
    expect(before).not.toBe("Georgia");

    applyDocumentDefaults(editor, {
      theme: "office",
      themeKind: "theme",
      styleSet: "elegant",
    });
    const after = (
      editor.state.doc.attrs as {
        styles?: { default?: { document?: { run?: { font?: string } } } };
      }
    ).styles?.default?.document?.run?.font;
    expect(after).toBe("Georgia");
  });

  it("carries the stored theme into the new document's attrs", () => {
    const storage = fakeStorage();
    writeDocumentDefaults(
      { theme: "integral", themeKind: "theme-color", styleSet: "modern" },
      storage,
    );
    const stored = readDocumentDefaults(storage);
    const json = normalizeDocument(newDocumentJSON(stored));
    const settings = (json.attrs as { documentExtras?: { settings?: Record<string, unknown> } })
      .documentExtras?.settings;
    expect(settings?.theme).toEqual({ id: "integral", kind: "theme-color" });
    expect(settings?.styleSet).toBe("modern");
    // Applying the defaults to a fresh document mutates the model only — the
    // stored JSON stays the source the element reloads.
    const editor = makeEditor(json);
    applyDocumentDefaults(editor, stored);
    expect(
      (
        editor.state.doc.attrs as {
          styles?: { default?: { document?: { run?: { font?: string } } } };
        }
      ).styles?.default?.document?.run?.font,
    ).toBe("Calibri");
  });

  it("builds a blank paragraph document when no defaults are stored", () => {
    const json = newDocumentJSON(null);
    expect(json.attrs).toBeUndefined();
    expect(json.content).toEqual([{ type: "paragraph" }]);
  });
});
