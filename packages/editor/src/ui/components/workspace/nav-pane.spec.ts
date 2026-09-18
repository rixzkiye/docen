// @vitest-environment happy-dom
import { docxExtensions } from "@docen/docx";
import { Editor } from "@tiptap/core";
import { afterEach, describe, expect, it } from "vitest";

import { DocenNavPane } from "./nav-pane";

const created: DocenNavPane[] = [];
afterEach(() => {
  while (created.length) created.pop()!.remove();
});

async function settle(): Promise<void> {
  for (let i = 0; i < 3; i++) await new Promise((r) => setTimeout(r, 0));
}

function createSampleDocContent(): any {
  return {
    type: "doc",
    content: [
      {
        type: "paragraph",
        attrs: { heading: "Heading1" },
        content: [{ type: "text", text: "Introduction" }],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "This is the introduction body paragraph." }],
      },
      {
        type: "paragraph",
        attrs: { heading: "Heading2" },
        content: [{ type: "text", text: "Background" }],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "Some background details on this topic." }],
      },
      {
        type: "paragraph",
        attrs: { heading: "Heading3" },
        content: [{ type: "text", text: "Deep Dive" }],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "Technical specifications and notes." }],
      },
      {
        type: "paragraph",
        attrs: { heading: "Heading1" },
        content: [{ type: "text", text: "Conclusion" }],
      },
      {
        type: "paragraph",
        content: [{ type: "text", text: "Final closing thoughts and summary." }],
      },
    ],
  };
}

async function mountNavPane(): Promise<DocenNavPane> {
  const pane = new DocenNavPane();
  created.push(pane);
  document.body.append(pane);
  await settle();
  return pane;
}

describe("DocenNavPane (<docen-nav-pane>) (W5.3)", () => {
  it("mounts with default Headings tab and switches tabs", async () => {
    const pane = await mountNavPane();
    expect(pane.tab).toBe("headings");

    let emittedTab = "";
    pane.addEventListener("navigation:tab", (e: any) => {
      emittedTab = e.detail.tab;
    });

    pane.setTab("pages");
    expect(pane.tab).toBe("pages");
    expect(emittedTab).toBe("pages");

    pane.setTab("results");
    expect(pane.tab).toBe("results");
    expect(emittedTab).toBe("results");
  });

  it("extracts headings hierarchy and computes hasChildren correctly", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: createSampleDocContent(),
    });

    const pane = await mountNavPane();
    pane.bindEditor(editor);
    await settle();

    expect(pane.headings).toHaveLength(4);
    // Heading 1: Introduction (level 1)
    expect(pane.headings[0]!.title).toBe("Introduction");
    expect(pane.headings[0]!.level).toBe(1);
    expect(pane.headings[0]!.hasChildren).toBe(true);

    // Heading 2: Background (level 2)
    expect(pane.headings[1]!.title).toBe("Background");
    expect(pane.headings[1]!.level).toBe(2);
    expect(pane.headings[1]!.hasChildren).toBe(true);

    // Heading 3: Deep Dive (level 3)
    expect(pane.headings[2]!.title).toBe("Deep Dive");
    expect(pane.headings[2]!.level).toBe(3);
    expect(pane.headings[2]!.hasChildren).toBe(false);

    // Heading 1: Conclusion (level 1)
    expect(pane.headings[3]!.title).toBe("Conclusion");
    expect(pane.headings[3]!.level).toBe(1);
    expect(pane.headings[3]!.hasChildren).toBe(false);

    const tree = pane.shadowRoot?.querySelector('[role="tree"]');
    expect(tree).not.toBeNull();
    const items = pane.shadowRoot?.querySelectorAll('[role="treeitem"]');
    expect(items).toHaveLength(4);
    expect(items?.[0]?.getAttribute("aria-level")).toBe("1");
    expect(items?.[0]?.getAttribute("aria-expanded")).toBe("true");
  });

  it("collapses and expands heading branches in treeview", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: createSampleDocContent(),
    });

    const pane = await mountNavPane();
    pane.bindEditor(editor);
    await settle();

    expect(pane.visibleHeadings).toHaveLength(4);

    // Collapse Introduction (which contains Background and Deep Dive)
    const introId = pane.headings[0]!.id;
    pane.toggleHeadingCollapse(introId);

    // Background & Deep Dive should be hidden; Conclusion should remain
    expect(pane.visibleHeadings).toHaveLength(2);
    expect(pane.visibleHeadings.map((h) => h.title)).toEqual(["Introduction", "Conclusion"]);

    // Expand Introduction
    pane.toggleHeadingCollapse(introId);
    expect(pane.visibleHeadings).toHaveLength(4);
  });

  it("filters headings live on search query and highlights matching text", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: createSampleDocContent(),
    });

    const pane = await mountNavPane();
    pane.bindEditor(editor);
    await settle();

    // Search for "Deep"
    pane.setSearchQuery("Deep");
    pane.setTab("headings");
    await settle();

    // Matching item "Deep Dive" and its ancestors "Background" and "Introduction" must be visible
    expect(pane.visibleHeadings.map((h) => h.title)).toEqual([
      "Introduction",
      "Background",
      "Deep Dive",
    ]);
    const deepItem = pane.visibleHeadings.find((h) => h.title === "Deep Dive");
    expect(deepItem?.hasMatch).toBe(true);
    expect(deepItem?.partMatch).toBe("Deep");

    // Clear search
    pane.clearSearch();
    await settle();
    expect(pane.visibleHeadings).toHaveLength(4);
  });

  it("provides live Results tab with snippets and match selection", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: createSampleDocContent(),
    });

    const pane = await mountNavPane();
    pane.bindEditor(editor);
    await settle();

    pane.setSearchQuery("introduction");
    expect(pane.searchResults.length).toBeGreaterThanOrEqual(2); // in heading and in body text

    let resultSelected: any = null;
    pane.addEventListener("nav:result-select", (e: any) => {
      resultSelected = e.detail;
    });

    const firstResult = pane.searchResults[0]!;
    pane.onResultClick(firstResult);

    expect(resultSelected).not.toBeNull();
    expect(resultSelected.from).toBe(firstResult.from);
    expect(editor.state.selection.from).toBe(firstResult.from);
  });

  it("supports Pages tab with page numbers and click selection", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: createSampleDocContent(),
    });

    const pane = await mountNavPane();
    pane.bindEditor(editor);
    pane.setTab("pages");
    await settle();

    expect(pane.pages.length).toBeGreaterThanOrEqual(1);

    let pageSelected: any = null;
    pane.addEventListener("nav:page-select", (e: any) => {
      pageSelected = e.detail;
    });

    pane.onPageClick(pane.pages[0]!);
    expect(pageSelected.page).toBe(1);
    expect(pane.activePageNumber).toBe(1);
  });

  it("reorders heading sections in document via drag-and-drop", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: createSampleDocContent(),
    });

    const pane = await mountNavPane();
    pane.bindEditor(editor);
    await settle();

    const introId = pane.headings[0]!.id;
    const conclusionId = pane.headings[3]!.id;

    let reordered: any = null;
    pane.addEventListener("nav:reorder-heading", (e: any) => {
      reordered = e.detail;
    });

    // Move Conclusion before Introduction
    const ok = pane.reorderHeading(conclusionId, introId, "before");
    expect(ok).toBe(true);
    expect(reordered).toEqual({
      sourceId: conclusionId,
      targetId: introId,
      position: "before",
    });

    // ProseMirror doc now starts with Conclusion!
    const firstNode = editor.state.doc.child(0);
    expect(firstNode.textContent).toBe("Conclusion");
  });

  it("supports keyboard navigation in tree (Arrow keys, Space, Enter)", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: createSampleDocContent(),
    });

    const pane = await mountNavPane();
    pane.bindEditor(editor);
    await settle();

    // Initially focused on first item
    expect(pane.effectiveFocusedId).toBe(pane.headings[0]!.id);

    // ArrowDown moves focus to second item
    pane.onTreeKeyDown(new KeyboardEvent("keydown", { key: "ArrowDown" }));
    expect(pane.effectiveFocusedId).toBe(pane.headings[1]!.id);

    // ArrowUp moves back
    pane.onTreeKeyDown(new KeyboardEvent("keydown", { key: "ArrowUp" }));
    expect(pane.effectiveFocusedId).toBe(pane.headings[0]!.id);

    // Space toggles collapse
    pane.onTreeKeyDown(new KeyboardEvent("keydown", { key: " " }));
    expect(pane.visibleHeadings).toHaveLength(2);

    // Space toggles expand again
    pane.onTreeKeyDown(new KeyboardEvent("keydown", { key: " " }));
    expect(pane.visibleHeadings).toHaveLength(4);

    // Enter selects heading
    let selectedHeading: any = null;
    pane.addEventListener("nav:select-heading", (e: any) => {
      selectedHeading = e.detail;
    });

    pane.onTreeKeyDown(new KeyboardEvent("keydown", { key: "Enter" }));
    expect(selectedHeading.id).toBe(pane.headings[0]!.id);
  });

  it("handles two-way scroll sync from canvas to heading tree", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: createSampleDocContent(),
    });

    const pane = await mountNavPane();
    pane.bindEditor(editor);
    await settle();

    // Create mock scroll container
    const scrollContainer = document.createElement("div");
    const h1El = document.createElement("h1");
    h1El.id = pane.headings[0]!.id;
    h1El.setAttribute("data-heading-id", pane.headings[0]!.id);
    const h2El = document.createElement("h2");
    h2El.id = pane.headings[3]!.id;
    h2El.setAttribute("data-heading-id", pane.headings[3]!.id);

    scrollContainer.append(h1El, h2El);
    document.body.append(scrollContainer);

    pane.bindScrollContainer(scrollContainer);

    // Mock getBoundingClientRect
    scrollContainer.getBoundingClientRect = () => ({ top: 0, bottom: 500 }) as any;
    h1El.getBoundingClientRect = () => ({ top: -200, bottom: -180 }) as any;
    h2El.getBoundingClientRect = () => ({ top: 10, bottom: 30 }) as any;

    // Trigger scroll
    scrollContainer.dispatchEvent(new Event("scroll"));
    await settle();

    expect(pane.activeHeadingId).toBe(pane.headings[3]!.id);
    scrollContainer.remove();
  });
});
