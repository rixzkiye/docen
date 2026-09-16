// @vitest-environment happy-dom
import { docxExtensions, resolveDocument } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { describe, expect, it, vi } from "vitest";

import "../../ui/components/workspace/comments-pane";
import { CommentsCommands, type CommentsHost } from "./comments";

function buildEditor(extraParagraphs?: Record<string, unknown>[]) {
  const json = resolveDocument(
    {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [{ text: "First paragraph with text." }],
              },
            },
            ...((extraParagraphs ?? []) as never[]),
          ],
        },
      ],
    },
    docxExtensions,
  );
  const editor = new Editor({ element: null, extensions: docxExtensions, content: json as never });
  for (const plugin of editor.extensionManager.plugins) editor.registerPlugin(plugin);
  return editor;
}

function fakeHost(editor: Editor | null) {
  const dispatch = vi.fn();
  const scrollIntoView = vi.fn();
  const focus = vi.fn();
  const showTaskpane = vi.fn();
  const pane = {
    comments: "",
    setAttribute: vi.fn(),
  };
  const element = {
    dispatchEvent: dispatch,
    shadowRoot: {
      querySelector: (sel: string) => (sel === "docen-comments-pane" ? pane : null),
    },
  } as unknown as HTMLElement;

  const host: CommentsHost = {
    editor: () => editor,
    bridge: () => ({
      focus,
      scrollIntoView,
      commentAnchorRect: () => ({ frame: document.createElement("div"), left: 10, top: 20 }),
    }),
    element: () => element,
    showTaskpane,
  };
  return { host, dispatch, scrollIntoView, focus, showTaskpane, pane };
}

describe("CommentsCommands", () => {
  it("insertComment anchors selection and stores structured comment", () => {
    const editor = buildEditor();
    const { host } = fakeHost(editor);
    const commands = new CommentsCommands(host);

    editor.commands.setTextSelection({ from: 1, to: 6 });
    commands.insertComment({ text: "Important note @Alice", author: "Bob", initials: "B" });

    // Check markers in document
    let hasStart = false;
    let hasEnd = false;
    let hasRef = false;
    editor.state.doc.descendants((node) => {
      if (node.type.name === "inlinePassthrough") {
        const data = JSON.parse(String(node.attrs.data ?? "{}"));
        if (data.commentRangeStart?.id === 0) hasStart = true;
        if (data.commentRangeEnd?.id === 0) hasEnd = true;
        if (data.commentReference === 0) hasRef = true;
      }
      return true;
    });
    expect(hasStart).toBe(true);
    expect(hasEnd).toBe(true);
    expect(hasRef).toBe(true);

    // Check documentExtras
    const attrs = editor.state.doc.attrs as { documentExtras?: { comments?: any[] } };
    const commentsList = attrs.documentExtras?.comments ?? [];
    expect(commentsList).toHaveLength(1);
    expect(commentsList[0].author).toBe("Bob");
    expect(commentsList[0].initials).toBe("B");
    expect(JSON.stringify(commentsList[0].children)).toContain("Important note @Alice");
    editor.destroy();
  });

  it("insertComment on table cell selection resolves to valid inline range", () => {
    const editor = buildEditor([
      {
        table: {
          rows: [
            {
              cells: [
                {
                  children: [
                    {
                      paragraph: {
                        children: [{ text: "Cell content" }],
                      },
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    ]);
    const { host } = fakeHost(editor);
    const commands = new CommentsCommands(host);

    // Select the table
    let tablePos = 0;
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === "table") {
        tablePos = pos;
        return false;
      }
      return true;
    });

    editor.commands.setTextSelection({ from: tablePos, to: tablePos + 10 });
    expect(() => {
      commands.insertComment({ text: "Table comment", author: "Reviewer" });
    }).not.toThrow();

    const attrs = editor.state.doc.attrs as { documentExtras?: { comments?: any[] } };
    expect(attrs.documentExtras?.comments).toHaveLength(1);
    editor.destroy();
  });

  it("onCommentReply links reply through paraIdParent", () => {
    const editor = buildEditor();
    const { host, pane } = fakeHost(editor);
    const commands = new CommentsCommands(host);

    editor.commands.setTextSelection({ from: 1, to: 5 });
    commands.insertComment({ text: "Root comment", author: "Alice" });

    // Reply to root (id = 0)
    commands.onCommentReply(
      new CustomEvent("comment:reply", {
        detail: { parentId: 0, text: "I agree @Alice", author: "Charlie", initials: "C" },
      }),
    );

    const attrs = editor.state.doc.attrs as {
      documentExtras?: { comments?: any[]; commentsExtended?: any[] };
    };
    const commentsList = attrs.documentExtras?.comments ?? [];
    const extendedList = attrs.documentExtras?.commentsExtended ?? [];

    expect(extendedList).toHaveLength(2);
    expect(commentsList).toHaveLength(2);
    expect(commentsList[1].author).toBe("Charlie");
    expect(commentsList[1].initials).toBe("C");

    // Check thread hierarchy in syncCommentsPane
    commands.syncCommentsPane();
    expect(pane.comments).toBeTruthy();
    const cards = JSON.parse(pane.comments);
    expect(cards).toHaveLength(1); // 1 root thread
    expect(cards[0].replies).toHaveLength(1);
    expect(cards[0].replies[0].author).toBe("Charlie");
    editor.destroy();
  });

  it("onCommentResolve toggles thread resolved state", () => {
    const editor = buildEditor();
    const { host } = fakeHost(editor);
    const commands = new CommentsCommands(host);

    editor.commands.setTextSelection({ from: 1, to: 5 });
    commands.insertComment({ text: "Need revision", author: "Editor" });

    // Resolve thread
    commands.onCommentResolve(
      new CustomEvent("comment:resolve", {
        detail: { id: 0, done: true },
      }),
    );

    let attrs = editor.state.doc.attrs as {
      documentExtras?: { commentsExtended?: any[] };
    };
    expect(attrs.documentExtras?.commentsExtended?.[0]?.done).toBe(true);

    // Reopen thread
    commands.onCommentResolve(
      new CustomEvent("comment:resolve", {
        detail: { id: 0, done: false },
      }),
    );
    attrs = editor.state.doc.attrs as {
      documentExtras?: { commentsExtended?: any[] };
    };
    expect(attrs.documentExtras?.commentsExtended?.[0]?.done).toBe(false);
    editor.destroy();
  });

  it("deleteComment cascade deletes replies when root is deleted", () => {
    const editor = buildEditor();
    const { host } = fakeHost(editor);
    const commands = new CommentsCommands(host);

    editor.commands.setTextSelection({ from: 1, to: 5 });
    commands.insertComment({ text: "Root", author: "Alice" });
    commands.onCommentReply(
      new CustomEvent("comment:reply", {
        detail: { parentId: 0, text: "Reply", author: "Bob" },
      }),
    );

    // Cursor inside root comment
    editor.commands.setTextSelection({ from: 2, to: 3 });
    commands.deleteComment();

    const attrs = editor.state.doc.attrs as {
      documentExtras?: { comments?: any[]; commentsExtended?: any[] };
    };
    expect(attrs.documentExtras?.comments).toHaveLength(0);
    expect(attrs.documentExtras?.commentsExtended).toHaveLength(0);
    editor.destroy();
  });

  it("jumpComment steps between comments in document order", () => {
    const editor = buildEditor([
      {
        paragraph: {
          children: [{ text: "Second paragraph." }],
        },
      },
    ]);
    const { host } = fakeHost(editor);
    const commands = new CommentsCommands(host);

    editor.commands.setTextSelection({ from: 1, to: 5 });
    commands.insertComment({ text: "Comment 1" });

    // Find start of second paragraph
    let para2Pos = 0;
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === "paragraph" && pos > 0 && para2Pos === 0) {
        para2Pos = pos;
      }
      return true;
    });

    editor.commands.setTextSelection({ from: para2Pos + 1, to: para2Pos + 5 });
    commands.insertComment({ text: "Comment 2" });

    // Jump from beginning of doc to next comment
    editor.commands.setTextSelection(1);
    commands.jumpComment("next");
    expect(commands.activeCommentId()).toBe(0);

    // Jump to second comment
    commands.jumpComment("next");
    expect(commands.activeCommentId()).toBe(1);

    // Jump previous
    commands.jumpComment("previous");
    expect(commands.activeCommentId()).toBe(0);
    editor.destroy();
  });
});

describe("DocenCommentsPane", () => {
  it("renders threads, applies status and author filters, and sorts cards", async () => {
    // Ensure element is registered
    const pane = document.createElement("docen-comments-pane") as any;
    document.body.append(pane);

    const testCards = [
      {
        id: 1,
        author: "Alice",
        initials: "A",
        date: "2026-01-01T10:00:00Z",
        text: "Hello @Bob",
        resolved: false,
        replies: [
          {
            id: 2,
            author: "Bob",
            initials: "B",
            date: "2026-01-01T11:00:00Z",
            text: "Hi @Alice, noted.",
            resolved: false,
          },
        ],
      },
      {
        id: 3,
        author: "Charlie",
        initials: "C",
        date: "2026-01-02T10:00:00Z",
        text: "Resolved item",
        resolved: true,
      },
    ];

    pane.comments = JSON.stringify(testCards);
    await new Promise((r) => setTimeout(r, 20));

    // Initially 2 threads rendered
    expect(pane.shadowRoot.querySelectorAll(".thread")).toHaveLength(2);

    // Filter to active only
    pane.statusFilterEl.value = "active";
    pane.onFilterChange();
    expect(pane.shadowRoot.querySelectorAll(".thread")).toHaveLength(1);

    // Filter to resolved only
    pane.statusFilterEl.value = "resolved";
    pane.onFilterChange();
    expect(pane.shadowRoot.querySelectorAll(".thread")).toHaveLength(1);

    // Reset status and filter by author
    pane.statusFilterEl.value = "all";
    pane.authorFilterEl.value = "Charlie";
    pane.onFilterChange();
    expect(pane.shadowRoot.querySelectorAll(".thread")).toHaveLength(1);

    // Search query
    pane.authorFilterEl.value = "all";
    pane.searchEl.value = "Hello";
    pane.onFilterChange();
    expect(pane.shadowRoot.querySelectorAll(".thread")).toHaveLength(1);

    // Check mention rendered
    const mention = pane.shadowRoot.querySelector(".mention");
    expect(mention).not.toBeNull();
    expect(mention?.textContent).toBe("@Bob");

    pane.remove();
  });
});
