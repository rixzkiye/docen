// @vitest-environment happy-dom
// Protection host-domain coverage: the production logic in ./protection driven
// through a view backed by a REAL Editor and a REAL restrict-editing pane,
// plus the real edit bridge with the production SDT predicate. No
// re-implementation of the host algorithm.
import { docxExtensions } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { describe, expect, it } from "vitest";

import { DocenRestrictEditingPane } from "../ui/components/workspace/restrict-editing-pane";
import {
  applyProtectionMode,
  enforceProtection,
  isInsideEditableSdt,
  stopProtection,
  withProtection,
  type ProtectionHostView,
} from "./protection";

// leafer's module graph touches these globals at import time (the bridge
// import below pulls it in).
(globalThis as any).CanvasRenderingContext2D ??= class {};
(globalThis as any).Path2D ??= class {};
(globalThis as any).ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
};
const fakeCtx = () =>
  new Proxy({ measureText: (s: string) => ({ width: s.length * 8 }) } as Record<string, unknown>, {
    get: (t, p: string) => (p in t ? t[p] : () => {}),
    set: () => true,
  });
const canvasProto = (globalThis as any).HTMLCanvasElement?.prototype;
if (canvasProto) canvasProto.getContext = () => fakeCtx();

const { mountEditBridge } = await import("./canvas/edit-bridge");

const content = {
  type: "doc",
  content: [
    { type: "paragraph", content: [{ type: "text", text: "Outside text" }] },
    {
      type: "sdtBlock",
      attrs: { properties: { title: "Field" } },
      content: [{ type: "paragraph", content: [{ type: "text", text: "Inside field" }] }],
    },
    {
      type: "sdtBlock",
      attrs: { properties: { title: "Locked Field", cannotEdit: true } },
      content: [{ type: "paragraph", content: [{ type: "text", text: "Locked content" }] }],
    },
  ],
};

function makeEditor(json: unknown = content): Editor {
  return new Editor({ element: null, extensions: docxExtensions, content: json as never });
}

/** The host view DocenDocument builds, with a real editor + real pane. */
function makeView(json: unknown = content) {
  const editor = makeEditor(json);
  const pane = new DocenRestrictEditingPane();
  document.body.append(pane);
  const state: { mode?: string; docProtected: boolean; syncs: number } = {
    mode: undefined,
    docProtected: false,
    syncs: 0,
  };
  const view: ProtectionHostView = {
    editor: () => editor,
    settings: () =>
      ((editor.state.doc.attrs as { documentExtras?: { settings?: Record<string, unknown> } })
        .documentExtras?.settings ?? {}) as Record<string, unknown>,
    commitSettings: (settings) => {
      const attrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: Record<string, unknown> };
      const extras = attrs.documentExtras ?? {};
      editor.view.dispatch(
        editor.state.tr.setDocAttribute("documentExtras", { ...extras, settings }),
      );
    },
    setMode: (mode, docProtected) => {
      state.mode = mode;
      state.docProtected = docProtected;
      state.syncs++;
      // The host's #syncEditable formula: readOnly/comments make the editor
      // non-editable; forms stays editable (the bridge predicate gates input).
      const editable = !docProtected;
      if (editor.isEditable !== editable) editor.setEditable(editable);
    },
    pane: () => pane as never,
  };
  return { view, editor, pane, state };
}

function posIn(editor: Editor, text: string): number {
  let pos = -1;
  editor.state.doc.descendants((node, p) => {
    if (node.isText && node.text === text) pos = p + 1;
  });
  if (pos < 0) throw new Error(`text not found: ${text}`);
  return pos;
}

describe("protection domain (pane path)", () => {
  it("enforce persists readOnly, blocks editing; stop clears and restores", () => {
    const { view, editor, state } = makeView();
    enforceProtection(view, {
      type: "readOnly",
      formattingRestricted: true,
      passwordHash: "hash-1",
    });
    expect((editor.state.doc.attrs as any).documentExtras?.settings?.documentProtection).toEqual({
      edit: "readOnly",
      hash: "hash-1",
      formatting: true,
    });
    expect(editor.isEditable).toBe(false);
    expect(state.mode).toBe("readOnly");
    expect(state.docProtected).toBe(true);

    stopProtection(view);
    expect(
      (editor.state.doc.attrs as any).documentExtras?.settings?.documentProtection,
    ).toBeUndefined();
    expect(editor.isEditable).toBe(true);
    expect(state.mode).toBeUndefined();
    expect(state.docProtected).toBe(false);
  });
});

describe("protection domain (Options path)", () => {
  it("preserves hash/formatting when switching modes and syncs the pane", () => {
    const { view, editor, pane, state } = makeView();
    enforceProtection(view, {
      type: "readOnly",
      formattingRestricted: true,
      passwordHash: "keep-me",
    });

    // The Options commit's settings write + mode sync (production helpers).
    view.commitSettings(withProtection(view.settings(), "comments"));
    applyProtectionMode(view, "comments");
    expect((editor.state.doc.attrs as any).documentExtras?.settings?.documentProtection).toEqual({
      edit: "comments",
      hash: "keep-me",
      formatting: true,
    });
    expect(state.mode).toBe("comments");
    expect(editor.isEditable).toBe(false);
    expect(pane.protectionType).toBe("comments");
    expect(pane.isEnforced).toBe(true);
    expect(pane.formattingRestricted).toBe(true);

    applyProtectionMode(view, "forms");
    expect(state.mode).toBe("forms");
    expect(editor.isEditable).toBe(true);

    view.commitSettings(withProtection(view.settings(), "none"));
    applyProtectionMode(view, "none");
    expect(
      (editor.state.doc.attrs as any).documentExtras?.settings?.documentProtection,
    ).toBeUndefined();
    expect(pane.isEnforced).toBe(false);
  });

  it("withProtection leaves unrelated settings untouched", () => {
    const next = withProtection(
      { updateFields: true, documentProtection: { edit: "readOnly", hash: "h", formatting: true } },
      "forms",
    );
    expect(next.updateFields).toBe(true);
    expect(next.documentProtection).toEqual({ edit: "forms", hash: "h", formatting: true });
  });
});

describe("isInsideEditableSdt", () => {
  it("is true only inside an unlocked SDT that fully contains the selection", () => {
    const editor = makeEditor();
    editor.commands.setTextSelection(posIn(editor, "Outside text"));
    expect(isInsideEditableSdt(editor)).toBe(false);

    editor.commands.setTextSelection(posIn(editor, "Inside field"));
    expect(isInsideEditableSdt(editor)).toBe(true);

    editor.commands.setTextSelection(posIn(editor, "Locked content"));
    expect(isInsideEditableSdt(editor)).toBe(false);

    // A range crossing the control boundary is not "inside".
    editor.commands.setTextSelection({
      from: posIn(editor, "Inside field"),
      to: posIn(editor, "Outside text"),
    });
    expect(isInsideEditableSdt(editor)).toBe(false);
  });
});

describe("forms enforcement through the edit bridge", () => {
  function mountBridge(predicate: (ed: Editor) => boolean) {
    const host = document.createElement("div");
    const inputHost = document.createElement("div");
    document.body.append(host, inputHost);
    const bridge = mountEditBridge({
      host,
      inputHost,
      content: content as never,
      onDoc: () => {},
      extensions: docxExtensions,
      canEdit: predicate,
    });
    const ta = inputHost.querySelector("textarea")!;
    return { bridge, ta };
  }

  function type(ta: HTMLTextAreaElement, data: string): void {
    ta.dispatchEvent(
      new InputEvent("beforeinput", {
        data,
        inputType: "insertText",
        bubbles: true,
        cancelable: true,
      }),
    );
  }

  it("blocks typing outside the SDT and allows it inside (locked stays blocked)", () => {
    const { bridge, ta } = mountBridge(isInsideEditableSdt);
    const editor = bridge.editor;
    const before = editor.state.doc.textContent;

    editor.commands.setTextSelection(posIn(editor, "Outside text"));
    type(ta, "X");
    expect(editor.state.doc.textContent).toBe(before);

    editor.commands.setTextSelection(posIn(editor, "Locked content"));
    type(ta, "X");
    expect(editor.state.doc.textContent).toBe(before);

    editor.commands.setTextSelection(posIn(editor, "Inside field"));
    type(ta, "X");
    expect(editor.state.doc.textContent).not.toBe(before);
    expect(editor.state.doc.textContent).toContain("X");
  });
});
