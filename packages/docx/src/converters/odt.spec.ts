import { describe, expect, it } from "vitest";

import type { JSONContent } from "../core";
import { generateODT } from "./odt";

describe("ODT Converter", () => {
  it("generates a valid OpenDocument Text zip archive", async () => {
    const doc: JSONContent = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1 },
          content: [{ type: "text", text: "OpenDocument Title" }],
        },
        {
          type: "paragraph",
          attrs: { textAlign: "center" },
          content: [
            { type: "text", text: "Formatted in " },
            { type: "text", text: "bold", marks: [{ type: "bold" }] },
            { type: "text", text: " and " },
            { type: "text", text: "italic", marks: [{ type: "italic" }] },
          ],
        },
      ],
    };

    const bytes = await generateODT(doc);
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.length).toBeGreaterThan(100);

    // ZIP signature: 0x50, 0x4b, 0x03, 0x04 ("PK\x03\x04")
    expect(bytes[0]).toBe(0x50);
    expect(bytes[1]).toBe(0x4b);
    expect(bytes[2]).toBe(0x03);
    expect(bytes[3]).toBe(0x04);

    const str = new TextDecoder().decode(bytes);
    expect(str).toContain("application/vnd.oasis.opendocument.text");
    expect(str).toContain("content.xml");
    expect(str).toContain("styles.xml");
    expect(str).toContain("META-INF/manifest.xml");
    expect(str).toContain("OpenDocument Title");
    expect(str).toContain("bold");
  });
});
