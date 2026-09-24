// @vitest-environment node
import { describe, expect, it } from "vitest";

import { serializeNodeScene } from "./node-scene";
import { rowTexts } from "./pdf-scene";

describe("rowTexts — Leafer row → PDF row string", () => {
  it("passes plain text rows through and tracks newlines between them", () => {
    expect(rowTexts("Title\nbody", [{ text: "Title" }, { text: "body" }])).toEqual([
      "Title",
      "body",
    ]);
  });

  it("re-inserts the spaces Leafer drops from justified word rows", () => {
    // Leafer's word mode stores one entry per word (multi-char) and renders
    // the gaps as advances — the source string still carries the spaces.
    const source = "Halo dunia yang indah sekali";
    const rows = [
      { data: [{ char: "Halo" }, { char: "dunia" }, { char: "yang" }] },
      { data: [{ char: "indah" }, { char: "sekali" }] },
    ];
    expect(rowTexts(source, rows)).toEqual(["Halo dunia yang ", "indah sekali"]);
  });

  it("re-inserts spaces around char-mode single-char entries", () => {
    const source = "ab cd ef";
    const rows = [
      { data: [{ char: "a" }, { char: "b" }, { char: "c" }, { char: "d" }] },
      { data: [{ char: "e" }, { char: "f" }] },
    ];
    expect(rowTexts(source, rows)).toEqual(["ab cd ", "ef"]);
  });

  it("keeps explicit space entries from overflow char-mode rows", () => {
    const rows = [
      { data: [{ char: "a" }, { char: "b" }, { char: " " }, { char: "c" }, { char: "d" }] },
    ];
    expect(rowTexts("ab cd", rows)).toEqual(["ab cd"]);
  });

  it("aligns a data row that follows a plain text row", () => {
    const source = "Title\nbody text here";
    const rows = [
      { text: "Title" },
      { data: [{ char: "body" }, { char: "text" }, { char: "here" }] },
    ];
    expect(rowTexts(source, rows)).toEqual(["Title", "body text here"]);
  });

  it("keeps case-transformed entries (length-preserving) aligned", () => {
    const source = "puji syukur";
    const rows = [{ data: [{ char: "PUJI" }, { char: "SYUKUR" }] }];
    expect(rowTexts(source, rows)).toEqual(["PUJI SYUKUR"]);
  });

  it("starts the next row after a hard break in the source", () => {
    const source = "ab\ncd";
    const rows = [
      { data: [{ char: "a" }, { char: "b" }] },
      { data: [{ char: "c" }, { char: "d" }] },
    ];
    expect(rowTexts(source, rows)).toEqual(["ab", "cd"]);
  });

  it("round-trips the serialized scene rows for a Leafer-shaped Text element", async () => {
    const source = "Puji syukur penulis panjatkan ke hadirat Allah";
    const root = {
      tag: "Group",
      worldTransform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      children: [
        {
          tag: "Text",
          worldTransform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
          text: source,
          textDrawData: {
            rows: [
              {
                x: 0,
                y: 14,
                width: 300,
                data: [
                  { char: "Puji" },
                  { char: "syukur" },
                  { char: "penulis" },
                  { char: "panjatkan" },
                ],
              },
              {
                x: 0,
                y: 34,
                width: 300,
                data: [{ char: "ke" }, { char: "hadirat" }, { char: "Allah" }],
              },
            ],
          },
          fontSize: 16,
          fontFamily: "Times New Roman",
          fill: "#000000",
        },
      ],
    };
    const scene = await serializeNodeScene(root as never, 595, 842);
    const text = scene.nodes[0];
    expect(text?.type).toBe("text");
    if (text?.type !== "text") throw new Error("expected a text node");
    expect(text.rows.map((row) => row.text)).toEqual([
      "Puji syukur penulis panjatkan ",
      "ke hadirat Allah",
    ]);
  });
});
