// @vitest-environment happy-dom
// Insert → Online Pictures: the host insert path and the DOCX round-trip. The
// image the dialog hands over must land as a normal image node (alt + natural
// size clamped to the content width) and come back out of a save/parse cycle
// as an embedded picture.
import { docxExtensions, generateDOCX, parseDOCX, type JSONContent } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { describe, expect, it, vi } from "vitest";

import { InsertDomain, type InsertHostView } from "./host/insert";

const PNG_1PX =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function makeEditor(): Editor {
  return new Editor({
    element: null,
    extensions: docxExtensions,
    content: { type: "doc", content: [{ type: "paragraph" }] },
  });
}

function makeInsert(editor: Editor, contentWidthPx = 600): InsertDomain {
  const host: InsertHostView = {
    element: () => ({}) as HTMLElement,
    root: () => null,
    editor: () => editor,
    bridge: () => undefined,
    flow: () => ({ contentWidthPx }) as never,
    pages: () => [],
    hyphenation: () => ({}),
    setHyphenation: vi.fn(),
    merge: () => ({}) as never,
    getJSON: () => editor.getJSON() as JSONContent,
    setJSON: vi.fn(),
    renderDoc: vi.fn(),
    setTextSelection: vi.fn(),
    openBookmarkDialog: vi.fn(),
  };
  return new InsertDomain(host);
}

const imageNodes = (json: JSONContent | { content?: JSONContent[] }): JSONContent[] => {
  const found: JSONContent[] = [];
  const walk = (node: JSONContent): void => {
    if (node.type === "image") found.push(node);
    for (const child of node.content ?? []) walk(child);
  };
  if ((json as JSONContent).type) walk(json as JSONContent);
  return found;
};

describe("InsertDomain.insertOnlinePicture", () => {
  it("inserts the picture with alt text and the natural size", () => {
    const editor = makeEditor();
    makeInsert(editor).insertOnlinePicture({
      src: PNG_1PX,
      width: 320,
      height: 200,
      alt: "team photo",
    });
    const [image] = imageNodes(editor.getJSON());
    expect(image).toBeDefined();
    expect(image!.attrs).toMatchObject({
      src: PNG_1PX,
      alt: "team photo",
      width: 320,
      height: 200,
    });
  });

  it("clamps an oversized picture to the content width, keeping the aspect", () => {
    const editor = makeEditor();
    makeInsert(editor, 600).insertOnlinePicture({
      src: PNG_1PX,
      width: 1200,
      height: 600,
      alt: "",
    });
    const [image] = imageNodes(editor.getJSON());
    expect(image!.attrs).toMatchObject({ width: 600, height: 300 });
  });

  it("round-trips as a normal embedded image through DOCX", async () => {
    const editor = makeEditor();
    makeInsert(editor, 600).insertOnlinePicture({
      src: PNG_1PX,
      width: 300,
      height: 150,
      alt: "Online Picture",
    });
    const bytes = (await generateDOCX(editor.getJSON(), { prepare: false })) as Uint8Array;
    const parsed = (await parseDOCX(bytes)) as JSONContent;
    const [image] = imageNodes(parsed);
    expect(image).toBeDefined();
    expect(String(image!.attrs?.src)).toMatch(/^data:image\/png;base64,/);
    expect(image!.attrs).toMatchObject({ alt: "Online Picture", width: 300, height: 150 });
  });
});
