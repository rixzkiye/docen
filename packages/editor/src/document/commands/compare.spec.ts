import { describe, expect, it } from "vitest";

import { combineDocs, compareDocs, diffTokens, tokenize } from "./compare";

describe("compare and combine engine", () => {
  it("tokenizes sentences properly into words and symbols", () => {
    const tokens = tokenize("Hello, world! 123");
    expect(tokens).toEqual(["Hello", ",", " ", "world", "!", " ", "123"]);
  });

  it("computes diff tokens for additions and deletions", () => {
    const a = tokenize("The quick brown fox");
    const b = tokenize("The fast brown fox jumps");
    const diff = diffTokens(a, b);

    // Should detect "quick" deleted, "fast" inserted, and " jumps" inserted
    expect(diff.some((d) => d.type === "del" && d.text.includes("quick"))).toBe(true);
    expect(diff.some((d) => d.type === "ins" && d.text.includes("fast"))).toBe(true);
    expect(diff.some((d) => d.type === "ins" && d.text.includes("jumps"))).toBe(true);
    expect(diff.some((d) => d.type === "same" && d.text.includes("fox"))).toBe(true);
  });

  it("generates tracked changes with insertion and deletion marks", () => {
    const docA = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "The quick brown fox jumps." }],
        },
      ],
    };

    const docB = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "The very fast fox jumps high." }],
        },
      ],
    };

    const compared = compareDocs(docA, docB, { author: "Reviewer A" });
    expect(compared.type).toBe("doc");
    const p = compared.content?.[0];
    expect(p).toBeDefined();
    expect(p?.content).toBeDefined();

    // Check that there are deletion and insertion marks with author "Reviewer A"
    const hasDeletion = p?.content?.some((r: any) =>
      r.marks?.some((m: any) => m.type === "deletion" && m.attrs?.author === "Reviewer A"),
    );
    const hasInsertion = p?.content?.some((r: any) =>
      r.marks?.some((m: any) => m.type === "insertion" && m.attrs?.author === "Reviewer A"),
    );

    expect(hasDeletion).toBe(true);
    expect(hasInsertion).toBe(true);
  });

  it("handles combining multiple revisions into base", () => {
    const baseDoc = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Line 1" }],
        },
      ],
    };

    const rev1 = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Line 1 modified" }],
        },
      ],
    };

    const rev2 = {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "Line 1 modified by two" }],
        },
      ],
    };

    const combined = combineDocs(baseDoc, [
      { doc: rev1, author: "Author1" },
      { doc: rev2, author: "Author2" },
    ]);

    expect(combined.type).toBe("doc");
    expect(combined.content?.length).toBeGreaterThan(0);
  });
});
