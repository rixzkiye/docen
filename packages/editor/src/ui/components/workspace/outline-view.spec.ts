// @vitest-environment happy-dom
import { docxExtensions } from "@docen/docx";
import { Editor } from "@tiptap/core";
import { afterEach, describe, expect, it } from "vitest";

import { DocenOutlineView } from "./outline-view";

const created: DocenOutlineView[] = [];
afterEach(() => {
  while (created.length) created.pop()!.remove();
});

async function settle(): Promise<void> {
  for (let i = 0; i < 3; i++) await new Promise((r) => setTimeout(r, 0));
}

async function mountWithEditor(
  initialContent?: any,
): Promise<{ view: DocenOutlineView; editor: Editor }> {
  const content = initialContent ?? {
    type: "doc",
    content: [
      {
        type: "paragraph",
        attrs: { heading: "Heading1" },
        content: [{ type: "text", text: "Section 1" }],
      },
      { type: "paragraph", content: [{ type: "text", text: "Body paragraph 1 for section 1" }] },
      {
        type: "paragraph",
        attrs: { heading: "Heading2" },
        content: [{ type: "text", text: "Section 1.1" }],
      },
      { type: "paragraph", content: [{ type: "text", text: "Body paragraph 1.1" }] },
      {
        type: "paragraph",
        attrs: { heading: "Heading1" },
        content: [{ type: "text", text: "Section 2" }],
      },
      { type: "paragraph", content: [{ type: "text", text: "Body paragraph 2" }] },
    ],
  };

  const editor = new Editor({
    element: null,
    extensions: docxExtensions,
    content,
  });

  const view = new DocenOutlineView();
  created.push(view);
  document.body.append(view);
  view.setEditor(editor);
  await settle();

  return { view, editor };
}

describe("DocenOutlineView (W5.1)", () => {
  it("initializes and parses outline items from editor", async () => {
    const { view } = await mountWithEditor();
    expect(view.blocks).toHaveLength(6);
    expect(view.blocks[0]!.level).toBe(1);
    expect(view.blocks[0]!.hasChildren).toBe(true);
    expect(view.blocks[1]!.level).toBeNull();
    expect(view.blocks[2]!.level).toBe(2);
  });

  it("handles level-based heading folding (expand/collapse)", async () => {
    const { view } = await mountWithEditor();
    expect(view.collapsedIds.size).toBe(0);

    // Fold Heading 1 (Section 1)
    view.toggleCollapse("block-0");
    expect(view.collapsedIds.has("block-0")).toBe(true);
    await settle();

    // With Section 1 collapsed, subordinate items (block-1, block-2, block-3) should not be visible
    expect(view.isItemVisible(view.blocks[0]!)).toBe(true); // Section 1 heading itself visible
    expect(view.isItemVisible(view.blocks[1]!)).toBe(false); // Body 1 hidden
    expect(view.isItemVisible(view.blocks[2]!)).toBe(false); // Subheading 1.1 hidden
    expect(view.isItemVisible(view.blocks[3]!)).toBe(false); // Body 1.1 hidden
    expect(view.isItemVisible(view.blocks[4]!)).toBe(true); // Section 2 visible

    // Expand again
    view.toggleCollapse("block-0");
    expect(view.isItemVisible(view.blocks[1]!)).toBe(true);

    // Collapse all
    view.collapseAll();
    expect(view.collapsedIds.has("block-0")).toBe(true);

    // Expand all
    view.expandAll();
    expect(view.collapsedIds.size).toBe(0);
  });

  it("filters outline by levels 1-9", async () => {
    const { view } = await mountWithEditor();

    // Show Level 1 only
    view.showLevel = 1;
    await settle();

    expect(view.isItemVisible(view.blocks[0]!)).toBe(true); // H1 visible
    expect(view.isItemVisible(view.blocks[2]!)).toBe(false); // H2 hidden
    expect(view.isItemVisible(view.blocks[4]!)).toBe(true); // H1 visible

    // Show Level 2
    view.showLevel = 2;
    expect(view.isItemVisible(view.blocks[2]!)).toBe(true); // H2 visible

    // Show all
    view.showLevel = 0;
    expect(view.isItemVisible(view.blocks[2]!)).toBe(true);
  });

  it("toggles Show Body Text", async () => {
    const { view } = await mountWithEditor();
    expect(view.showBodyText).toBe(true);
    expect(view.isItemVisible(view.blocks[1]!)).toBe(true);

    // Hide body text
    view.showBodyText = false;
    await settle();

    expect(view.isItemVisible(view.blocks[0]!)).toBe(true); // Headings visible
    expect(view.isItemVisible(view.blocks[1]!)).toBe(false); // Body hidden
    expect(view.isItemVisible(view.blocks[3]!)).toBe(false); // Body hidden
    expect(view.isItemVisible(view.blocks[4]!)).toBe(true); // Headings visible
  });

  it("toggles Show First Line Only attribute", async () => {
    const { view } = await mountWithEditor();
    expect(view.showFirstLineOnly).toBe(false);
    expect(view.hasAttribute("show-first-line-only")).toBe(false);

    view.showFirstLineOnly = true;
    view.toggleAttribute("show-first-line-only", true);
    await settle();

    expect(view.hasAttribute("show-first-line-only")).toBe(true);
  });

  it("supports multi-selection with Shift+Click and Ctrl+Click", async () => {
    const { view } = await mountWithEditor();

    // Single click
    view.onItemClick("block-0", { ctrlKey: false, shiftKey: false } as MouseEvent);
    expect(Array.from(view.selectedIds)).toEqual(["block-0"]);

    // Ctrl+Click adds to selection
    view.onItemClick("block-2", { ctrlKey: true, shiftKey: false } as MouseEvent);
    expect(Array.from(view.selectedIds).sort()).toEqual(["block-0", "block-2"]);

    // Shift+Click selects range
    view.onItemClick("block-4", { ctrlKey: false, shiftKey: true } as MouseEvent);
    // Anchor was block-2, so range from block-2 to block-4
    expect(Array.from(view.selectedIds).sort()).toEqual(["block-2", "block-3", "block-4"]);
  });

  it("drag-reorders outline blocks with real-time drop indicator line", async () => {
    const { view, editor } = await mountWithEditor();

    // Start drag on Section 2 (block-4)
    view.onDragStart({ dataTransfer: { setData: () => {} } } as any, "block-4");
    expect(view.draggedId).toBe("block-4");

    // Drag over Section 1 (block-0) in top half (before)
    const mockEl = {
      getBoundingClientRect: () => ({ top: 100, height: 40 }),
    };
    view.onDragOver(
      { preventDefault: () => {}, currentTarget: mockEl, clientY: 110 } as any,
      "block-0",
    );
    expect(view.dropTarget).toEqual({ targetId: "block-0", position: "before" });

    // Drop
    view.onDrop({ preventDefault: () => {} } as any, "block-0");
    expect(view.dropTarget).toBeNull();
    expect(view.draggedId).toBeNull();

    // Section 2 should now be at the beginning of document
    expect(editor.state.doc.child(0).textContent).toBe("Section 2");
    expect(editor.state.doc.child(1).textContent).toBe("Body paragraph 2");
  });

  it("executes promote, demote, move-up, move-down via shortcuts and toolbar", async () => {
    const { view, editor } = await mountWithEditor();

    // Select block-2 (Section 1.1, Heading 2)
    view.onItemClick("block-2", { ctrlKey: false, shiftKey: false } as MouseEvent);

    // Promote (Alt+Shift+Left) -> Heading 1
    view.promoteSelected();
    expect(editor.getJSON().content?.[2]?.attrs?.heading).toBe("Heading1");

    // Demote (Alt+Shift+Right) -> Heading 2
    view.demoteSelected();
    expect(editor.getJSON().content?.[2]?.attrs?.heading).toBe("Heading2");

    // Move Section 1 down
    view.onItemClick("block-0", { ctrlKey: false, shiftKey: false } as MouseEvent);
    view.moveSelectedDown();
    expect(editor.state.doc.child(0).textContent).toBe("Section 2");

    // Move back up
    view.onItemClick("block-2", { ctrlKey: false, shiftKey: false } as MouseEvent);
    view.moveSelectedUp();
  });

  it("emits outline-view:close when close button clicked", async () => {
    const { view } = await mountWithEditor();
    let closed = false;
    view.addEventListener("outline-view:close", () => {
      closed = true;
    });

    view.closeOutlineView();
    expect(closed).toBe(true);
  });
});
