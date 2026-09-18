import { unzipSync } from "@office-open/core";
import type { DocumentOptions, ParagraphChild } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { generateDOCXSync, parseDOCXSync, resolveDocument } from "../index";

/**
 * DocumentOptions-authored pictures (`{ picture: { type, data } }`) are the
 * persistence-model form of every image; the runtime resolver must rebuild the
 * `src` from the data bytes *or* from a base64/data-URL string (office-open
 * accepts all three), or the drawing silently disappears at generate time.
 * Header/footer pictures additionally need their `wp:docPr` id carried through
 * parse→generate or every save allocates a fresh id and the bytes drift.
 */

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const PAGE = {
  pageSize: { width: 11906, height: 16838 },
  pageMargin: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
};

function docWithHeaderLogo(): DocumentOptions {
  const picture = {
    picture: {
      type: "png",
      data: PNG,
      transformation: { width: "20px", height: "20px" },
      nonVisualProperties: { name: "logo-hdr" },
    },
  } as unknown as ParagraphChild;
  return {
    sections: [
      {
        properties: PAGE,
        headers: {
          default: [{ paragraph: { children: [{ text: "logo " }, picture] } }],
        },
        footers: {
          default: [{ paragraph: { children: [{ text: "foot " }, picture] } }],
        },
        children: [{ paragraph: { children: [{ text: "body " }, picture] } }],
      },
    ],
  };
}

function imageOf(json: JSONContent, slot: "sectionHeaders" | "sectionFooters"): JSONContent {
  const group = (json.attrs ?? {})[slot] as { default: JSONContent[] };
  const paragraph = group.default[0] as JSONContent;
  return paragraph.content?.[1] as JSONContent;
}

describe("DocumentOptions pictures", () => {
  it("resolves string picture data to an image src (body, header, footer)", () => {
    const json = resolveDocument(docWithHeaderLogo());
    for (const slot of ["sectionHeaders", "sectionFooters"] as const) {
      const image = imageOf(json, slot);
      expect(image.type, slot).toBe("image");
      expect(image.attrs?.src, slot).toBe(PNG);
    }
    const bodyImage = (json.content?.[0]?.content ?? [])[1] as JSONContent;
    expect(bodyImage.attrs?.src).toBe(PNG);
  });

  it("writes header/footer drawings and keeps saves byte-stable", () => {
    const gen1 = generateDOCXSync(resolveDocument(docWithHeaderLogo()), { prepare: false });
    const zip1 = unzipSync(gen1);
    expect(Object.keys(zip1).some((name) => name.startsWith("word/media/"))).toBe(true);
    for (const part of ["word/header1.xml", "word/footer1.xml"]) {
      expect(new TextDecoder().decode(zip1[part]), part).toContain("<a:blip");
    }

    const parsed = parseDOCXSync(gen1);
    const headerImage = imageOf(parsed, "sectionHeaders");
    expect((headerImage.attrs?.altText as { id?: number } | null)?.id).toBeDefined();

    const gen2 = generateDOCXSync(parsed, { prepare: false });
    const gen3 = generateDOCXSync(parseDOCXSync(gen2), { prepare: false });
    expect(Buffer.from(gen2).equals(Buffer.from(gen3))).toBe(true);

    const docPrs = (bytes: Uint8Array) =>
      [
        ...new TextDecoder()
          .decode(unzipSync(bytes)["word/header1.xml"])
          .matchAll(/<wp:docPr[^>]*>/g),
      ].map((match) => match[0]);
    expect(docPrs(gen3)).toEqual(docPrs(gen2));
  });
});
