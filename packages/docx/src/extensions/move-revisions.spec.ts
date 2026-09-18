import type { DocumentOptions } from "@office-open/docx";
import { describe, expect, it } from "vitest";

import { compileDocument, docxExtensions, resolveDocument } from "../index";
import type { ProjectContext } from "../layout/project/context";
import { projectParagraph } from "../layout/project/paragraph";

describe("W9.1 Move Revisions Tracking", () => {
  it("compiles and resolves move range start/end nodes", () => {
    const doc = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    moveFromRangeStart: {
                      id: 1,
                      name: "move1",
                      author: "Alice",
                      date: "2026-09-18T08:00:00Z",
                    },
                  },
                  {
                    text: "Moved content",
                    movedFrom: { id: 1, author: "Alice", date: "2026-09-18T08:00:00Z" },
                  },
                  { moveFromRangeEnd: { id: 1 } },
                  {
                    moveToRangeStart: {
                      id: 2,
                      name: "move2",
                      author: "Alice",
                      date: "2026-09-18T08:00:00Z",
                    },
                  },
                  {
                    text: "Moved content",
                    movedTo: { id: 2, author: "Alice", date: "2026-09-18T08:00:00Z" },
                  },
                  { moveToRangeEnd: { id: 2 } },
                ],
              },
            },
          ],
        },
      ],
    };

    const resolved = resolveDocument(doc as unknown as DocumentOptions, docxExtensions);
    const para = resolved.content?.[0];
    expect(para?.type).toBe("paragraph");

    const types = para?.content?.map((c) => c.type);
    expect(types).toContain("moveFromRangeStart");
    expect(types).toContain("moveFromRangeEnd");
    expect(types).toContain("moveToRangeStart");
    expect(types).toContain("moveToRangeEnd");

    const compiled = compileDocument(resolved, docxExtensions);
    const compiledPara = compiled.sections[0]!.children[0] as {
      paragraph: { children: unknown[] };
    };
    expect(compiledPara.paragraph.children.some((c: any) => "moveFromRangeStart" in c)).toBe(true);
    expect(compiledPara.paragraph.children.some((c: any) => "moveFromRangeEnd" in c)).toBe(true);
    expect(compiledPara.paragraph.children.some((c: any) => "moveToRangeStart" in c)).toBe(true);
    expect(compiledPara.paragraph.children.some((c: any) => "moveToRangeEnd" in c)).toBe(true);
    expect(compiledPara.paragraph.children.some((c: any) => c.movedFrom?.id === 1)).toBe(true);
    expect(compiledPara.paragraph.children.some((c: any) => c.movedTo?.id === 2)).toBe(true);
  });

  it("projects moveFrom with green strikethrough and moveTo with green double-underline", () => {
    const ctx: ProjectContext = {
      styles: undefined,
      characterStyles: new Map(),
      numberings: { byNumId: new Map(), abstractNumberings: new Map() } as any,
      listCounters: new Map(),
      openComments: new Set(),
      footnoteOrdinals: new Map(),
      endnoteOrdinals: new Map(),
      revisionAuthorColors: new Map(),
      stateful: { hit: false },
      markup: { view: "all" },
    };

    const para = {
      children: [
        {
          text: "Original Moved",
          movedFrom: { id: 10, author: "Bob", date: "2026-09-18T08:00:00Z" },
        },
        {
          text: "Destination Moved",
          movedTo: { id: 11, author: "Bob", date: "2026-09-18T08:00:00Z" },
        },
      ],
    };

    const layout = projectParagraph(para as any, ctx);
    expect(layout.inline).toHaveLength(2);

    const inlineMoveFrom = layout.inline[0] as {
      style?: { strikethrough?: boolean; color?: string };
    };
    expect(inlineMoveFrom.style?.strikethrough).toBe(true);
    expect(inlineMoveFrom.style?.color).toBe("008000");

    const inlineMoveTo = layout.inline[1] as {
      style?: {
        underline?: boolean;
        underlineStyle?: string;
        underlineColor?: string;
        color?: string;
      };
    };
    expect(inlineMoveTo.style?.underline).toBe(true);
    expect(inlineMoveTo.style?.underlineStyle).toBe("double");
    expect(inlineMoveTo.style?.underlineColor).toBe("008000");
    expect(inlineMoveTo.style?.color).toBe("008000");
  });

  it("projects container-style movedFrom and movedTo elements correctly", () => {
    const ctx: ProjectContext = {
      styles: undefined,
      characterStyles: new Map(),
      numberings: { byNumId: new Map(), abstractNumberings: new Map() } as any,
      listCounters: new Map(),
      openComments: new Set(),
      footnoteOrdinals: new Map(),
      endnoteOrdinals: new Map(),
      revisionAuthorColors: new Map(),
      stateful: { hit: false },
      markup: { view: "all" },
    };

    const para = {
      children: [
        {
          movedFrom: {
            id: 20,
            author: "Alice",
            date: "2026-09-18T08:00:00Z",
            children: [{ text: "Container MoveFrom" }],
          },
        },
        {
          movedTo: {
            id: 21,
            author: "Alice",
            date: "2026-09-18T08:00:00Z",
            children: [{ text: "Container MoveTo" }],
          },
        },
      ],
    };

    const layout = projectParagraph(para as any, ctx);
    expect(layout.inline).toHaveLength(2);

    const inlineMoveFrom = layout.inline[0] as {
      style?: { strikethrough?: boolean; color?: string };
    };
    expect(inlineMoveFrom.style?.strikethrough).toBe(true);
    expect(inlineMoveFrom.style?.color).toBe("008000");

    const inlineMoveTo = layout.inline[1] as {
      style?: {
        underline?: boolean;
        underlineStyle?: string;
        underlineColor?: string;
        color?: string;
      };
    };
    expect(inlineMoveTo.style?.underline).toBe(true);
    expect(inlineMoveTo.style?.underlineStyle).toBe("double");
    expect(inlineMoveTo.style?.underlineColor).toBe("008000");
    expect(inlineMoveTo.style?.color).toBe("008000");
  });
});
