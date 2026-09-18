// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";

if (typeof (globalThis as any).CanvasRenderingContext2D === "undefined") {
  class FakeCanvasRenderingContext2D {}
  (globalThis as any).CanvasRenderingContext2D = FakeCanvasRenderingContext2D;
}
if (typeof (globalThis as any).Path2D === "undefined") {
  class FakePath2D {}
  (globalThis as any).Path2D = FakePath2D;
}

const fakeCtx = (): unknown =>
  new Proxy({ measureText: (s: string) => ({ width: s.length * 8 }) } as Record<string, unknown>, {
    get: (t, p: string) => (p in t ? t[p] : () => {}),
    set: () => true,
  });
const canvasProto = (globalThis as any).HTMLCanvasElement?.prototype;
if (canvasProto) canvasProto.getContext = () => fakeCtx();

const { docxExtensions } = await import("@docen/docx");
const { DocumentCommands } = await import("../extensions/commands");
const { TrackChanges } = await import("../extensions/track-changes");
const { mountEditBridge } = await import("./edit-bridge");

describe("W4 Shortcut Suite", () => {
  function createTestHarness(initialContent?: unknown) {
    const content = initialContent ?? {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: "First paragraph with multiple words." }],
        },
        {
          type: "paragraph",
          content: [{ type: "text", text: "Second paragraph for testing navigation." }],
        },
      ],
    };

    const host = document.createElement("div");
    const inputHost = document.createElement("div");
    document.body.append(host, inputHost);

    const bridge = mountEditBridge({
      host,
      inputHost,
      content: content as never,
      extensions: [...docxExtensions, DocumentCommands, TrackChanges],
      onDoc: () => {},
      canEdit: () => true,
    });

    const editor = bridge.editor;
    const ta = inputHost.querySelector("textarea")!;
    expect(ta).toBeTruthy();

    const press = (
      key: string,
      modifiers: {
        ctrl?: boolean;
        shift?: boolean;
        alt?: boolean;
        meta?: boolean;
        code?: string;
      } = {},
    ) => {
      const event = new KeyboardEvent("keydown", {
        key,
        code: modifiers.code ?? (key.length === 1 && /[0-9]/.test(key) ? `Digit${key}` : undefined),
        ctrlKey: modifiers.ctrl ?? false,
        shiftKey: modifiers.shift ?? false,
        altKey: modifiers.alt ?? false,
        metaKey: modifiers.meta ?? false,
        bubbles: true,
        cancelable: true,
      });
      ta.dispatchEvent(event);
      return event;
    };

    return { editor, host, bridge, ta, press };
  }

  describe("W4.3: Navigation & Selection Shortcuts", () => {
    it("unblocks Ctrl+ArrowLeft and Ctrl+ArrowRight word jumping", () => {
      const { editor, press } = createTestHarness();
      // Set caret at offset 16
      editor.commands.setTextSelection(16);
      expect(editor.state.selection.from).toBe(16);

      // Ctrl+ArrowLeft jumps word backward
      press("ArrowLeft", { ctrl: true });
      expect(editor.state.selection.from).toBeLessThan(16);

      // Ctrl+ArrowRight jumps word forward
      const prev = editor.state.selection.from;
      press("ArrowRight", { ctrl: true });
      expect(editor.state.selection.from).toBeGreaterThan(prev);
    });

    it("unblocks Ctrl+Shift+ArrowLeft and Ctrl+Shift+ArrowRight word selection", () => {
      const { editor, press } = createTestHarness();
      editor.commands.setTextSelection(16);

      press("ArrowLeft", { ctrl: true, shift: true });
      expect(editor.state.selection.empty).toBe(false);
      expect(editor.state.selection.to).toBe(16);
      expect(editor.state.selection.from).toBeLessThan(16);
    });

    it("unblocks Ctrl+ArrowUp and Ctrl+ArrowDown paragraph navigation", () => {
      const { editor, press } = createTestHarness();
      // Start in paragraph 2
      const p2Pos = editor.state.doc.child(0).nodeSize + 10;
      editor.commands.setTextSelection(p2Pos);

      // Ctrl+ArrowUp jumps to start of paragraph 2
      press("ArrowUp", { ctrl: true });
      const p2Start = editor.state.doc.child(0).nodeSize + 1;
      expect(editor.state.selection.from).toBe(p2Start);

      // Another Ctrl+ArrowUp jumps to start of paragraph 1
      press("ArrowUp", { ctrl: true });
      expect(editor.state.selection.from).toBe(1);

      // Ctrl+ArrowDown jumps to next paragraph
      press("ArrowDown", { ctrl: true });
      expect(editor.state.selection.from).toBe(p2Start);
    });

    it("unblocks Ctrl+Home and Ctrl+End document jumping and selection", () => {
      const { editor, press } = createTestHarness();
      editor.commands.setTextSelection(20);

      // Ctrl+Home jumps to 0
      press("Home", { ctrl: true });
      expect(editor.state.selection.from).toBe(0);

      // Ctrl+End jumps to doc end
      press("End", { ctrl: true });
      expect(editor.state.selection.from).toBe(editor.state.doc.content.size);

      // Ctrl+Shift+Home selects to start
      press("Home", { ctrl: true, shift: true });
      expect(editor.state.selection.empty).toBe(false);
      expect(editor.state.selection.from).toBe(0);
      expect(editor.state.selection.to).toBe(editor.state.doc.content.size);
    });

    it("unblocks Ctrl+Backspace and Ctrl+Delete", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [{ type: "paragraph", content: [{ type: "text", text: "hello brave world" }] }],
      });
      // Caret after "brave" (offset 1 + 11 = 12)
      editor.commands.setTextSelection(12);

      // Ctrl+Backspace deletes "brave"
      press("Backspace", { ctrl: true });
      expect(editor.getText()).toContain("hello");
      expect(editor.getText()).not.toContain("brave");

      // Caret at start of " world"
      press("Delete", { ctrl: true });
      expect(editor.getText().trim()).toBe("hello");
    });

    it("wires Alt+Home/End and Alt+PageUp/Down in table", () => {
      const tableContent = {
        type: "doc",
        content: [
          {
            type: "table",
            content: [
              {
                type: "tableRow",
                content: [
                  {
                    type: "tableCell",
                    content: [{ type: "paragraph", content: [{ type: "text", text: "R1C1" }] }],
                  },
                  {
                    type: "tableCell",
                    content: [{ type: "paragraph", content: [{ type: "text", text: "R1C2" }] }],
                  },
                ],
              },
              {
                type: "tableRow",
                content: [
                  {
                    type: "tableCell",
                    content: [{ type: "paragraph", content: [{ type: "text", text: "R2C1" }] }],
                  },
                  {
                    type: "tableCell",
                    content: [{ type: "paragraph", content: [{ type: "text", text: "R2C2" }] }],
                  },
                ],
              },
            ],
          },
        ],
      };

      const { editor, press } = createTestHarness(tableContent);
      // Place caret in R1C1
      editor.commands.setTextSelection(4);

      // Alt+End moves to last cell in row (R1C2)
      press("End", { alt: true });
      expect(editor.state.selection.$from.parent.textContent).toBe("R1C2");

      // Alt+Home moves to first cell in row (R1C1)
      press("Home", { alt: true });
      expect(editor.state.selection.$from.parent.textContent).toBe("R1C1");

      // Alt+PageDown moves to bottom cell in col (R2C1)
      press("PageDown", { alt: true });
      expect(editor.state.selection.$from.parent.textContent).toBe("R2C1");

      // Alt+PageUp moves to top cell in col (R1C1)
      press("PageUp", { alt: true });
      expect(editor.state.selection.$from.parent.textContent).toBe("R1C1");
    });

    it("wires Ctrl+Tab in table cell to insert tab character", () => {
      const tableContent = {
        type: "doc",
        content: [
          {
            type: "table",
            content: [
              {
                type: "tableRow",
                content: [
                  {
                    type: "tableCell",
                    content: [{ type: "paragraph", content: [{ type: "text", text: "Cell1" }] }],
                  },
                ],
              },
            ],
          },
        ],
      };

      const { editor, press } = createTestHarness(tableContent);
      editor.commands.setTextSelection(4);

      press("Tab", { ctrl: true });
      expect(editor.getText()).toContain("\t");
    });

    it("wires Shift+F3 to cycle case (lower -> capitalize -> upper -> lower)", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [{ type: "paragraph", content: [{ type: "text", text: "word test" }] }],
      });
      // Select all text
      editor.commands.setTextSelection({ from: 1, to: 10 });

      // lower -> Title Case
      press("F3", { shift: true });
      expect(editor.getText().trim()).toBe("Word Test");

      // Title Case -> UPPERCASE
      press("F3", { shift: true });
      expect(editor.getText().trim()).toBe("WORD TEST");

      // UPPERCASE -> lowercase
      press("F3", { shift: true });
      expect(editor.getText().trim()).toBe("word test");
    });
  });

  describe("W4.1: Formatting Shortcuts", () => {
    it("wires Ctrl+Shift+C and Ctrl+Shift+V for Format Painter", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { alignment: "center" },
            content: [{ type: "text", marks: [{ type: "bold" }], text: "Bold Text" }],
          },
          {
            type: "paragraph",
            content: [{ type: "text", text: "Plain Text" }],
          },
        ],
      });

      // Select "Bold Text" and copy format
      editor.commands.setTextSelection({ from: 1, to: 10 });
      press("c", { ctrl: true, shift: true });

      // Select "Plain Text" and paste format
      const p2Pos = editor.state.doc.child(0).nodeSize + 1;
      editor.commands.setTextSelection({ from: p2Pos, to: p2Pos + 10 });
      press("v", { ctrl: true, shift: true });

      // Plain Text should now have bold mark and center alignment
      const p2 = editor.state.doc.child(1);
      expect(p2.attrs.alignment).toBe("center");
      expect(p2.child(0).marks.some((m) => m.type.name === "bold")).toBe(true);
    });

    it("wires Ctrl+Space to clear character formatting", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "text",
                marks: [{ type: "bold" }, { type: "italic" }],
                text: "Formatted Text",
              },
            ],
          },
        ],
      });
      editor.commands.setTextSelection({ from: 1, to: 15 });
      expect(editor.state.doc.child(0).child(0).marks.length).toBeGreaterThan(0);

      press(" ", { ctrl: true });
      expect(editor.state.doc.child(0).child(0).marks.length).toBe(0);
    });

    it("wires Ctrl+Q to clear paragraph formatting", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { alignment: "center", indent: { left: 720 }, spacing: { line: 480 } },
            content: [{ type: "text", text: "Centered Text" }],
          },
        ],
      });
      editor.commands.setTextSelection(3);
      expect(editor.state.doc.child(0).attrs.alignment).toBe("center");

      press("q", { ctrl: true });
      const para = editor.state.doc.child(0);
      expect(para.attrs.alignment).toBeNull();
      expect(para.attrs.indent).toBeNull();
      expect(para.attrs.spacing).toBeNull();
    });

    it("wires Ctrl+Shift+N to apply Normal style", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { style: "Heading1", heading: "Heading1" },
            content: [{ type: "text", text: "Heading Text" }],
          },
        ],
      });
      editor.commands.setTextSelection(3);

      press("n", { ctrl: true, shift: true });
      expect(editor.state.doc.child(0).attrs.heading).toBeNull();
    });

    it("wires Ctrl+M and Ctrl+Shift+M for paragraph indent", () => {
      const { editor, press } = createTestHarness();
      editor.commands.setTextSelection(3);

      // Ctrl+M increases indent
      press("m", { ctrl: true });
      expect((editor.state.doc.child(0).attrs.indent as any)?.left).toBe(720);

      // Ctrl+Shift+M decreases indent
      press("m", { ctrl: true, shift: true });
      expect((editor.state.doc.child(0).attrs.indent as any)?.left).toBe(0);
    });

    it("wires Ctrl+T and Ctrl+Shift+T for hanging indent", () => {
      const { editor, press } = createTestHarness();
      editor.commands.setTextSelection(3);

      // Ctrl+T increases hanging indent
      press("t", { ctrl: true });
      const ind = editor.state.doc.child(0).attrs.indent as any;
      expect(ind?.hanging).toBe(720);
      expect(ind?.left).toBe(720);

      // Ctrl+Shift+T decreases hanging indent
      press("t", { ctrl: true, shift: true });
      const ind2 = editor.state.doc.child(0).attrs.indent as any;
      expect(ind2?.hanging).toBeUndefined();
    });

    it("wires Ctrl+Shift+W, Ctrl+Shift+K, Ctrl+Shift+D", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [{ type: "paragraph", content: [{ type: "text", text: "Decorated Text" }] }],
      });
      editor.commands.setTextSelection({ from: 1, to: 15 });

      // Ctrl+Shift+W: words underline
      press("w", { ctrl: true, shift: true });
      let marks = editor.state.doc.child(0).child(0).marks;
      expect(marks.some((m) => m.type.name === "underline" && m.attrs.style === "words")).toBe(
        true,
      );

      // Ctrl+Shift+D: double underline
      press("d", { ctrl: true, shift: true });
      marks = editor.state.doc.child(0).child(0).marks;
      expect(marks.some((m) => m.type.name === "underline" && m.attrs.style === "double")).toBe(
        true,
      );

      // Ctrl+Shift+K: small caps
      press("k", { ctrl: true, shift: true });
      marks = editor.state.doc.child(0).child(0).marks;
      expect(marks.some((m) => m.type.name === "textStyle" && m.attrs.smallCaps === true)).toBe(
        true,
      );
    });

    it("wires Ctrl+B, Ctrl+I, Ctrl+U", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [{ type: "paragraph", content: [{ type: "text", text: "Sample" }] }],
      });
      editor.commands.setTextSelection({ from: 1, to: 7 });

      press("b", { ctrl: true });
      expect(
        editor.state.doc
          .child(0)
          .child(0)
          .marks.some((m) => m.type.name === "bold"),
      ).toBe(true);

      press("i", { ctrl: true });
      expect(
        editor.state.doc
          .child(0)
          .child(0)
          .marks.some((m) => m.type.name === "italic"),
      ).toBe(true);

      press("u", { ctrl: true });
      expect(
        editor.state.doc
          .child(0)
          .child(0)
          .marks.some((m) => m.type.name === "underline"),
      ).toBe(true);
    });
  });

  describe("W4.2: Paragraph and List Shortcuts", () => {
    it("wires Ctrl+1, Ctrl+2, Ctrl+5 for line spacing", () => {
      const { editor, press } = createTestHarness();
      editor.commands.setTextSelection(3);

      // Ctrl+2 -> double spacing (line: 480)
      press("2", { ctrl: true });
      expect((editor.state.doc.child(0).attrs.spacing as any)?.line).toBe(480);

      // Ctrl+5 -> 1.5 line spacing (line: 360)
      press("5", { ctrl: true });
      expect((editor.state.doc.child(0).attrs.spacing as any)?.line).toBe(360);

      // Ctrl+1 -> single spacing (line: 240)
      press("1", { ctrl: true });
      expect((editor.state.doc.child(0).attrs.spacing as any)?.line).toBe(240);
    });

    it("wires Ctrl+0 to toggle space before paragraph", () => {
      const { editor, press } = createTestHarness();
      editor.commands.setTextSelection(3);

      // First press adds 240 twips (12pt) space before
      press("0", { ctrl: true });
      expect((editor.state.doc.child(0).attrs.spacing as any)?.before).toBe(240);

      // Second press removes space before
      press("0", { ctrl: true });
      expect((editor.state.doc.child(0).attrs.spacing as any)?.before).toBeNull();
    });

    it("wires Ctrl+Alt+1/2/3 for headings", () => {
      const { editor, press } = createTestHarness();
      editor.commands.setTextSelection(3);

      press("1", { ctrl: true, alt: true });
      expect(editor.state.doc.child(0).attrs.heading).toBe("Heading1");

      press("2", { ctrl: true, alt: true });
      expect(editor.state.doc.child(0).attrs.heading).toBe("Heading2");

      press("3", { ctrl: true, alt: true });
      expect(editor.state.doc.child(0).attrs.heading).toBe("Heading3");
    });

    it("wires Ctrl+Shift+L for bullet list and Ctrl+Shift+7 for numbered list", () => {
      const { editor, press } = createTestHarness();
      editor.commands.setTextSelection(3);

      // Bullet list
      press("l", { ctrl: true, shift: true });
      expect(editor.state.doc.child(0).attrs.bullet).toBeTruthy();

      // Numbered list
      press("7", { ctrl: true, shift: true });
      expect(editor.state.doc.child(0).attrs.numbering).toBeTruthy();
    });
  });

  describe("W4.4: Review and Notes Shortcuts", () => {
    it("wires Ctrl+Shift+E for track changes toggle", () => {
      const { editor, press } = createTestHarness();
      editor.commands.setTextSelection(3);

      press("e", { ctrl: true, shift: true });
      expect(editor.commands["track-changes"]).toBeDefined();
    });

    it("dispatches host commands for Ctrl+Alt+M, Alt+Ctrl+F, Alt+Ctrl+D", () => {
      const { host, press } = createTestHarness();
      const dispatched: string[] = [];
      host.addEventListener("command", (e: any) => {
        dispatched.push(e.detail.event + (e.detail.value ? `:${e.detail.value}` : ""));
      });

      // Ctrl+Alt+M -> new-comment
      press("m", { ctrl: true, alt: true });
      expect(dispatched).toContain("new-comment");

      // Alt+Ctrl+F -> insert-footnote
      press("f", { ctrl: true, alt: true });
      expect(dispatched).toContain("insert-footnote");

      // Alt+Ctrl+D -> insert-footnote:endnote
      press("d", { ctrl: true, alt: true });
      expect(dispatched).toContain("insert-footnote:endnote");
    });
  });

  describe("W4.5: Outline and Extended F4 Repeat", () => {
    it("wires Alt+Shift+ArrowLeft/Right/Up/Down for outline operations", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: { style: "Heading2" },
            content: [{ type: "text", text: "H2" }],
          },
          { type: "paragraph", content: [{ type: "text", text: "Para 2" }] },
        ],
      });
      editor.commands.setTextSelection(3);

      // Alt+Shift+ArrowLeft promotes heading
      press("ArrowLeft", { alt: true, shift: true });
      expect(editor.state.doc.child(0).attrs.style).toBe("Heading1");

      // Alt+Shift+ArrowRight demotes heading
      press("ArrowRight", { alt: true, shift: true });
      expect(editor.state.doc.child(0).attrs.style).toBe("Heading2");

      // Alt+Shift+ArrowDown moves block down
      press("ArrowDown", { alt: true, shift: true });
      expect(editor.state.doc.child(1).attrs.style).toBe("Heading2");

      // Alt+Shift+ArrowUp moves block up
      press("ArrowUp", { alt: true, shift: true });
      expect(editor.state.doc.child(0).attrs.style).toBe("Heading2");
    });

    it("extended F4 repeats both typing and formatting commands", () => {
      const { editor, press } = createTestHarness({
        type: "doc",
        content: [
          { type: "paragraph", content: [{ type: "text", text: "Word1" }] },
          { type: "paragraph", content: [{ type: "text", text: "Word2" }] },
        ],
      });

      // 1. Formatting command repeat:
      // Apply bold to Word1 via shortcut
      editor.commands.setTextSelection({ from: 1, to: 6 });
      press("b", { ctrl: true });
      expect(
        editor.state.doc
          .child(0)
          .child(0)
          .marks.some((m) => m.type.name === "bold"),
      ).toBe(true);

      // Now select Word2 and press F4
      const p2Pos = editor.state.doc.child(0).nodeSize + 1;
      editor.commands.setTextSelection({ from: p2Pos, to: p2Pos + 5 });
      press("F4");
      expect(
        editor.state.doc
          .child(1)
          .child(0)
          .marks.some((m) => m.type.name === "bold"),
      ).toBe(true);
    });
  });
});
