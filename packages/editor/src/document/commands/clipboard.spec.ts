// @vitest-environment happy-dom
import { docxExtensions, type JSONContent } from "@docen/docx";
import { Editor, type Editor as EditorType } from "@docen/docx/core";
import { UndoRedo } from "@tiptap/extensions";
import { describe, expect, it, vi } from "vitest";

import { ClipboardCommands, type ClipboardHost, type PasteSource } from "./clipboard";

const buildEditor = (content?: JSONContent[]): EditorType =>
  new Editor({
    element: null,
    extensions: [...docxExtensions, UndoRedo],
    content: {
      type: "doc",
      content: content ?? [{ type: "paragraph", content: [{ type: "text", text: "Initial" }] }],
    },
  });

describe("ClipboardCommands", () => {
  it("pastes plain text via pasteSpecial('text')", async () => {
    const editor = buildEditor();
    const host: ClipboardHost = {
      editor: () => editor,
      bridge: () => undefined,
      element: () => document.createElement("div"),
    };
    const cmds = new ClipboardCommands(host);

    vi.spyOn(navigator.clipboard, "read").mockResolvedValue([
      {
        types: ["text/plain", "text/html"],
        getType: async (type: string) => {
          if (type === "text/plain") return new Blob(["Hello plain text"], { type: "text/plain" });
          return new Blob(["<strong>Bold HTML</strong>"], { type: "text/html" });
        },
      } as unknown as ClipboardItem,
    ]);

    await cmds.pasteSpecial("text");
    expect(editor.getText()).toContain("Hello plain text");
  });

  it("pastes formatted HTML via pasteSpecial('html')", async () => {
    const editor = buildEditor();
    const host: ClipboardHost = {
      editor: () => editor,
      bridge: () => undefined,
      element: () => document.createElement("div"),
    };
    const cmds = new ClipboardCommands(host);

    vi.spyOn(navigator.clipboard, "read").mockResolvedValue([
      {
        types: ["text/plain", "text/html"],
        getType: async (type: string) => {
          if (type === "text/plain") return new Blob(["fallback text"], { type: "text/plain" });
          return new Blob(["<p>HTML paragraph</p>"], { type: "text/html" });
        },
      } as unknown as ClipboardItem,
    ]);

    await cmds.pasteSpecial("html");
    expect(editor.getText()).toContain("HTML paragraph");
  });

  it("pastes RTF via pasteSpecial('rtf')", async () => {
    const editor = buildEditor();
    const host: ClipboardHost = {
      editor: () => editor,
      bridge: () => undefined,
      element: () => document.createElement("div"),
    };
    const cmds = new ClipboardCommands(host);

    const rtf = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Calibri;}}\\pard\\b Rich RTF text\\b0\\par}";
    vi.spyOn(navigator.clipboard, "read").mockResolvedValue([
      {
        types: ["text/plain", "text/rtf"],
        getType: async (type: string) => {
          if (type === "text/plain") return new Blob(["plain text"], { type: "text/plain" });
          return new Blob([rtf], { type: "text/rtf" });
        },
      } as unknown as ClipboardItem,
    ]);

    await cmds.pasteSpecial("rtf");
    expect(editor.getText()).toContain("Rich RTF text");
  });

  it("replays RTF paste in text and match modes", () => {
    const editor = buildEditor();
    const host: ClipboardHost = {
      editor: () => editor,
      bridge: () => undefined,
      element: () => document.createElement("div"),
    };
    const cmds = new ClipboardCommands(host);
    const rtf = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Calibri;}}\\pard\\b Rich RTF text\\b0\\par}";
    const source: PasteSource = {
      kind: "rtf",
      raw: rtf,
      text: "Rich RTF text",
    };

    const anchor = {
      frame: document.createElement("div"),
      left: 0,
      top: 0,
      height: 20,
    };
    const bridgeMock = {
      activeEditor: () => editor,
      pasteAnchorRect: () => anchor,
    };
    (host as { bridge: () => unknown }).bridge = () => bridgeMock;

    cmds.showPasteOptions(source);
    cmds.replayPaste("text");
    expect(editor.getText()).toContain("Rich RTF text");

    cmds.replayPaste("match");
    expect(editor.getText()).toContain("Rich RTF text");
  });
});
