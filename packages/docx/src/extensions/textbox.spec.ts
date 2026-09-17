import { unzipSync } from "@office-open/core";
import type { DocumentOptions, SectionChild } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import {
  compileDocument,
  docxExtensions,
  generateDOCXSync,
  parseDOCXSync,
  resolveDocument,
  type JSONContent,
} from "../index";

/**
 * VML text box (w:pict > v:shape > v:textbox) round-trips: the box's own data
 * (structured VML style + residual paragraph options) rides attrs verbatim on
 * both legs; the contained block stream resolves to editable nodes and
 * compiles back through the shared block walk.
 */

function roundTrip(children: SectionChild[]) {
  const doc: DocumentOptions = { sections: [{ children }] };
  const json = resolveDocument(doc, docxExtensions);
  const compiled = compileDocument(json, docxExtensions);
  return { json, child: compiled.sections[0].children[0] };
}

describe("textbox", () => {
  it("carries the box data verbatim and compiles content back", () => {
    const style = { width: "200pt", height: "100pt" } as const;
    const { json, child } = roundTrip([
      {
        textbox: {
          text: "caption",
          style,
          children: [{ paragraph: { text: "in-box" } }, { paragraph: { text: "more" } }],
        },
      },
    ]);
    const node = json.content?.[0] as {
      type: string;
      attrs?: { textbox?: unknown };
      content?: { type: string }[];
    };
    expect(node.type).toBe("textbox");
    expect(node.attrs?.textbox).toEqual({ text: "caption", style });
    expect(node.content?.map((c) => c.type)).toEqual(["paragraph", "paragraph"]);
    expect(child).toEqual({
      textbox: {
        text: "caption",
        style,
        children: [{ paragraph: "in-box" }, { paragraph: "more" }],
      },
    });
  });

  it("resolves a nested table through the block walk", () => {
    const { json, child } = roundTrip([
      {
        textbox: {
          style: { width: "300pt", height: "80pt" },
          children: [
            { paragraph: { text: "head" } },
            {
              table: {
                rows: [{ cells: [{ children: [{ paragraph: { text: "cell" } }] }] }],
              },
            },
          ],
        },
      },
    ]);
    const node = json.content?.[0] as { content?: { type: string }[] };
    expect(node.content?.map((c) => c.type)).toEqual(["paragraph", "table"]);
    const box = (
      child as {
        textbox: {
          children: {
            table: { rows: { cells: { children: { paragraph: string }[] }[] }[] };
          }[];
        };
      }
    ).textbox;
    // The table compiles with its projected widths re-added (shared table
    // behavior) — probe the structure, not the whole object.
    expect(box.children[1].table.rows[0].cells[0].children[0].paragraph).toBe("cell");
  });

  it("keeps an empty box resolvable with a placeholder paragraph", () => {
    const { json } = roundTrip([{ textbox: { style: { width: "100pt", height: "40pt" } } }]);
    const node = json.content?.[0] as { type: string; content?: unknown[] };
    expect(node.type).toBe("textbox");
    expect(node.content?.length).toBeGreaterThan(0);
  });

  // The engine's writer emits the box as w:p > w:r > w:pict (pict is an
  // EG_RunInnerContent); the reader now promotes that run-level textbox pict
  // back to the { textbox } branch, so the box is native on both legs.
  it("round-trips through the OPC package as a native textbox", () => {
    const json: JSONContent = {
      type: "doc",
      content: [
        {
          type: "textbox",
          attrs: { textbox: { style: { width: "200pt", height: "100pt" } } },
          content: [
            {
              type: "paragraph",
              content: [{ type: "text", text: "VML 文本框内容" }],
            },
          ],
        },
      ],
    };
    const gen1 = generateDOCXSync(json, { prepare: false }) as Uint8Array;
    const xml = new TextDecoder().decode(unzipSync(gen1)["word/document.xml"]);
    expect(xml).toContain("txbxContent");
    expect(xml).toContain("VML 文本框内容");

    const reparsed = parseDOCXSync(gen1);
    const box = reparsed.content?.find((n) => n.type === "textbox");
    expect(box?.content?.[0]?.content?.[0]?.text).toBe("VML 文本框内容");

    // Second cycle: the box stays the native branch, not a pict passthrough.
    const reparsed2 = parseDOCXSync(generateDOCXSync(reparsed, { prepare: false }) as Uint8Array);
    expect(reparsed2.content?.some((n) => n.type === "textbox")).toBe(true);
    expect(reparsed2.content?.some((n) => n.type === "inlinePassthrough")).toBe(false);
  });
});
