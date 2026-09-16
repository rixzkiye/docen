import { DocPartBehavior, DocPartGallery, type JSONContent } from "@docen/docx";
import { describe, expect, it } from "vitest";

import {
  autotextMatch,
  blocksFromGlossary,
  blocksOfDocAttrs,
  BUILDING_BLOCKS_VERSION,
  createBuildingBlock,
  DEFAULT_BLOCK_CATEGORY,
  DEFAULT_BLOCK_GALLERY,
  glossaryOfBlocks,
  groupBlocksByGallery,
  isDuplicateBlockName,
  parseBuildingBlocks,
  parseSlicePayload,
  sortBlocks,
  withBlocks,
  type BuildingBlock,
} from "./building-blocks";

const paragraph = (text: string, marks?: JSONContent["marks"]): JSONContent => ({
  type: "paragraph",
  content: [{ type: "text", text, ...(marks ? { marks } : {}) }],
});

const block = (name: string, overrides: Partial<BuildingBlock> = {}): BuildingBlock => ({
  id: `id-${name}`,
  name,
  gallery: DEFAULT_BLOCK_GALLERY,
  category: DEFAULT_BLOCK_CATEGORY,
  description: "",
  savedAt: 0,
  insertMode: "content",
  content: { openStart: 0, openEnd: 0, content: [paragraph(name)] },
  ...overrides,
});

describe("parseBuildingBlocks", () => {
  it("round-trips valid data and sanitizes defaults", () => {
    const data = parseBuildingBlocks({
      version: BUILDING_BLOCKS_VERSION,
      blocks: [
        {
          id: "a",
          name: " Signature ",
          gallery: "not-a-gallery",
          content: { openStart: 1, openEnd: 0, content: [paragraph("Hi")] },
        },
      ],
    });
    expect(data.blocks).toHaveLength(1);
    expect(data.blocks[0]).toMatchObject({
      id: "a",
      name: "Signature",
      gallery: DEFAULT_BLOCK_GALLERY,
      category: DEFAULT_BLOCK_CATEGORY,
      insertMode: "content",
      savedAt: 0,
    });
    expect(data.blocks[0]!.content.openStart).toBe(1);
  });

  it("returns an empty list for a foreign version or malformed shape", () => {
    for (const bad of [undefined, null, 42, "x", {}, { version: 99, blocks: [] }]) {
      expect(parseBuildingBlocks(bad)).toEqual({ version: BUILDING_BLOCKS_VERSION, blocks: [] });
    }
  });

  it("drops malformed blocks and duplicate ids", () => {
    const data = parseBuildingBlocks({
      version: BUILDING_BLOCKS_VERSION,
      blocks: [
        { id: "a", name: "Keep", content: { content: [paragraph("x")] } },
        { id: "a", name: "Duplicate id", content: { content: [paragraph("x")] } },
        { id: "b", name: "", content: { content: [paragraph("x")] } },
        { id: "c", name: "No content" },
        null,
      ],
    });
    expect(data.blocks.map((b) => b.name)).toEqual(["Keep"]);
  });
});

describe("documentExtras persistence", () => {
  it("writes both the docen list and the derived Word glossary part", () => {
    const extras = withBlocks({ settings: { zoom: 100 } }, [block("Signature")]);
    expect(extras.settings).toEqual({ zoom: 100 });
    expect(extras.docenBlocks).toMatchObject({ version: BUILDING_BLOCKS_VERSION });
    const glossary = extras.glossary as { parts: Array<Record<string, unknown>> };
    expect(glossary.parts).toHaveLength(1);
    expect(glossary.parts[0]).toMatchObject({
      name: "Signature",
      gallery: DEFAULT_BLOCK_GALLERY,
      guid: "id-Signature",
      behaviors: [DocPartBehavior.CONTENT],
    });
  });

  it("reads the docen list back from attrs (even an explicit empty one)", () => {
    const attrs = { documentExtras: withBlocks(undefined, [block("A"), block("B")]) };
    expect(blocksOfDocAttrs(attrs).map((b) => b.name)).toEqual(["A", "B"]);
    const emptied = { documentExtras: withBlocks(attrs.documentExtras, []) };
    expect(blocksOfDocAttrs(emptied)).toEqual([]);
    expect("glossary" in emptied.documentExtras).toBe(false);
  });

  it("imports Word glossary parts when no docen list exists", () => {
    const attrs = {
      documentExtras: {
        glossary: {
          parts: [
            {
              name: "Word Block",
              gallery: DocPartGallery.CUSTOM_AUTO_TEXT,
              category: "General",
              description: "from Word",
              guid: "{abc}",
              sections: [{ children: [{ paragraph: { text: "Hello" } }] }],
            },
          ],
        },
      },
    };
    const blocks = blocksOfDocAttrs(attrs);
    expect(blocks).toHaveLength(1);
    expect(blocks[0]).toMatchObject({
      id: "{abc}",
      name: "Word Block",
      gallery: DocPartGallery.CUSTOM_AUTO_TEXT,
      description: "from Word",
      insertMode: "content",
      savedAt: 0,
    });
    expect(JSON.stringify(blocks[0]!.content.content)).toContain("Hello");
  });

  it("falls back to the glossary when the docen list is corrupt", () => {
    const glossary = {
      parts: [{ name: "Real", sections: [{ children: [{ paragraph: { text: "x" } }] }] }],
    };
    for (const docenBlocks of [
      { version: 99, blocks: [{ name: "ignored" }] },
      { version: BUILDING_BLOCKS_VERSION, blocks: "not-an-array" },
    ]) {
      expect(
        blocksOfDocAttrs({ documentExtras: { docenBlocks, glossary } }).map((b) => b.name),
      ).toEqual(["Real"]);
    }
  });

  it("returns no blocks when neither key carries data", () => {
    expect(blocksOfDocAttrs(undefined)).toEqual([]);
    expect(blocksOfDocAttrs({ documentExtras: { glossary: "garbage" } })).toEqual([]);
  });
});

describe("glossary projection", () => {
  it("round-trips text, marks, categories and insert modes", () => {
    const source = block("Rich", {
      gallery: DocPartGallery.CUSTOM_AUTO_TEXT,
      category: "Letters",
      description: "Sign-off",
      insertMode: "paragraph",
      content: {
        openStart: 0,
        openEnd: 0,
        content: [paragraph("Best ", [{ type: "bold" }]), paragraph("regards,")],
      },
    });
    const glossary = glossaryOfBlocks([source]);
    const [restored] = blocksFromGlossary(glossary);
    expect(restored).toBeDefined();
    expect(restored).toMatchObject({
      id: "id-Rich",
      name: "Rich",
      gallery: DocPartGallery.CUSTOM_AUTO_TEXT,
      category: "Letters",
      description: "Sign-off",
      insertMode: "paragraph",
    });
    expect(restored!.content.content).toHaveLength(2);
    expect(JSON.stringify(restored!.content.content)).toContain("Best ");
    expect(JSON.stringify(restored!.content.content)).toContain('"bold"');
  });

  it("round-trips a table block", () => {
    const table: JSONContent = {
      type: "table",
      content: [
        {
          type: "tableRow",
          content: [
            { type: "tableCell", content: [{ type: "paragraph" }] },
            { type: "tableCell", content: [{ type: "paragraph" }] },
          ],
        },
      ],
    };
    const glossary = glossaryOfBlocks([
      block("Grid", { content: { openStart: 0, openEnd: 0, content: [table] } }),
    ]);
    const [restored] = blocksFromGlossary(glossary);
    expect(restored).toBeDefined();
    expect(restored!.content.content[0]?.type).toBe("table");
  });

  it("skips malformed glossary parts without throwing", () => {
    expect(
      blocksFromGlossary({ parts: [null, {}, { name: "" }, { name: "x", sections: [] }] }),
    ).toEqual([]);
    expect(blocksFromGlossary("nope")).toEqual([]);
  });

  it("keeps the docen list when the glossary projection skips unrepresentable content", () => {
    const odd = block("odd", {
      content: { openStart: 0, openEnd: 0, content: [{ type: "unknownBlockNode" }] },
    });
    const extras = withBlocks(undefined, [odd]);
    const data = extras.docenBlocks as { blocks: BuildingBlock[] };
    expect(data.blocks.map((b) => b.name)).toEqual(["odd"]);
    // The projection is best-effort — it must not throw and must not drop the
    // docen record (the reload path prefers it over the glossary anyway).
    expect(() => glossaryOfBlocks([odd])).not.toThrow();
  });
});

describe("createBuildingBlock / duplicates", () => {
  it("applies defaults and trims", () => {
    const created = createBuildingBlock(
      {
        name: "  Hello  ",
        description: "  d  ",
        content: { openStart: 0, openEnd: 0, content: [paragraph("x")] },
      },
      { id: "fixed", savedAt: 42 },
    );
    expect(created).toMatchObject({
      id: "fixed",
      name: "Hello",
      gallery: DEFAULT_BLOCK_GALLERY,
      category: DEFAULT_BLOCK_CATEGORY,
      description: "d",
      savedAt: 42,
      insertMode: "content",
    });
  });

  it("rejects case-insensitive duplicate names, ignoring the renamed block", () => {
    const blocks = [block("Signature"), block("Invoice")];
    expect(isDuplicateBlockName(blocks, "signature")).toBe(true);
    expect(isDuplicateBlockName(blocks, " invoice ")).toBe(true);
    expect(isDuplicateBlockName(blocks, "signature", "id-Signature")).toBe(false);
    expect(isDuplicateBlockName(blocks, "")).toBe(false);
  });
});

describe("autotextMatch (F3)", () => {
  const blocks = [block("Signature"), block("My Block"), block("Block"), block("signature line")];

  it("matches the exact typed name at the caret, case-insensitively", () => {
    expect(autotextMatch("Signature", blocks)?.block.name).toBe("Signature");
    expect(autotextMatch("signature", blocks)?.block.name).toBe("Signature");
    expect(autotextMatch("SIGNATURE", blocks)?.block.name).toBe("Signature");
    expect(autotextMatch("My Block", blocks)).toMatchObject({ back: 8 });
    expect(autotextMatch("Signature Line", blocks)?.block.name).toBe("signature line");
  });

  it("prefers the longest matching name", () => {
    expect(autotextMatch("My Block", blocks)?.block.name).toBe("My Block");
    expect(autotextMatch("Block", blocks)?.block.name).toBe("Block");
  });

  it("matches after a word boundary only", () => {
    expect(autotextMatch("xSignature", blocks)).toBeNull();
    expect(autotextMatch("(Signature", blocks)?.back).toBe(9);
    expect(autotextMatch("Hello Signature", blocks)?.back).toBe(9);
  });

  it("never matches a partial name or empty text", () => {
    expect(autotextMatch("Signat", blocks)).toBeNull();
    expect(autotextMatch("", blocks)).toBeNull();
    expect(autotextMatch("Signature", [])).toBeNull();
  });
});

describe("gallery grouping", () => {
  it("groups by gallery order, then category, then name", () => {
    const blocks = [
      block("Zed", { gallery: DocPartGallery.CUSTOM_AUTO_TEXT }),
      block("Beta", { gallery: DocPartGallery.CUSTOM_QUICK_PARTS, category: "B" }),
      block("Alpha", { gallery: DocPartGallery.CUSTOM_QUICK_PARTS, category: "A" }),
      block("Alpha", { gallery: DocPartGallery.CUSTOM_QUICK_PARTS, category: "B" }),
    ];
    expect(sortBlocks(blocks).map((b) => `${b.category}/${b.name}`)).toEqual([
      "A/Alpha",
      "B/Alpha",
      "B/Beta",
      "General/Zed",
    ]);
    const groups = groupBlocksByGallery(blocks);
    expect(groups.map((g) => g.gallery)).toEqual([
      DocPartGallery.CUSTOM_QUICK_PARTS,
      DocPartGallery.CUSTOM_AUTO_TEXT,
    ]);
    expect(groups[0]!.blocks).toHaveLength(3);
  });
});

describe("parseSlicePayload", () => {
  it("parses the clipboard-lane payload and tolerates corruption", () => {
    expect(
      parseSlicePayload(JSON.stringify({ openStart: 1, openEnd: 0, content: [paragraph("x")] })),
    ).toMatchObject({
      openStart: 1,
      openEnd: 0,
    });
    expect(parseSlicePayload("not json")).toBeNull();
    expect(parseSlicePayload(JSON.stringify({ content: [] }))).toBeNull();
    expect(parseSlicePayload(null)).toBeNull();
  });
});
