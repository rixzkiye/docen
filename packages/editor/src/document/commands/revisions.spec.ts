import { docxExtensions, resolveDocument } from "@docen/docx";
import { Editor } from "@docen/docx/core";
import { describe, expect, it, vi } from "vitest";

import { RevisionsCommands, type RevisionsHost } from "./revisions";

/**
 * Balloon click routing (Word's markup area): a comment balloon dispatches
 * comment:select and opens the comments pane; a revision balloon dispatches
 * the reviewing pane's revision:select and opens the revisions pane.
 */

const buildEditor = () => {
  const json = resolveDocument(
    {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    text: "styled",
                    revision: {
                      id: 5,
                      author: "Ada",
                      date: "2026-01-01T00:00:00Z",
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    },
    docxExtensions,
  );
  const editor = new Editor({ element: null, extensions: docxExtensions, content: json as never });
  return editor;
};

function fakeHost(editor: Editor | null) {
  const dispatch = vi.fn();
  const scrollIntoView = vi.fn();
  const togglePane = vi.fn();
  const host: RevisionsHost = {
    editor: () => editor,
    bridge: () => ({ scrollIntoView }),
    element: () => ({ dispatchEvent: dispatch }) as unknown as HTMLElement,
    togglePane,
  };
  return { host, dispatch, scrollIntoView, togglePane };
}

describe("RevisionsCommands balloon routing", () => {
  it("routes a revision balloon to revision:select and the reviewing pane", () => {
    const editor = buildEditor();
    const { host, dispatch, togglePane } = fakeHost(editor);
    new RevisionsCommands(host).onBalloonSelect({ kind: "revision", id: 5 });
    expect(dispatch).toHaveBeenCalledTimes(1);
    const event = dispatch.mock.calls[0]![0] as CustomEvent<{ index?: number }>;
    expect(event.type).toBe("revision:select");
    expect(event.detail).toEqual({ index: 0 });
    expect(togglePane).toHaveBeenCalledWith("revisions");
    editor.destroy();
  });

  it("routes a comment balloon to comment:select and the comments pane", () => {
    const editor = buildEditor();
    const { host, dispatch, togglePane } = fakeHost(editor);
    new RevisionsCommands(host).onBalloonSelect({ kind: "comment", id: 7 });
    expect(dispatch).toHaveBeenCalledTimes(1);
    const event = dispatch.mock.calls[0]![0] as CustomEvent<{ id?: number }>;
    expect(event.type).toBe("comment:select");
    expect(event.detail).toEqual({ id: 7 });
    expect(togglePane).toHaveBeenCalledWith("comments");
    editor.destroy();
  });

  it("ignores a revision balloon whose record is gone", () => {
    const editor = buildEditor();
    const { host, dispatch, togglePane } = fakeHost(editor);
    new RevisionsCommands(host).onBalloonSelect({ kind: "revision", id: 99 });
    expect(dispatch).not.toHaveBeenCalled();
    expect(togglePane).not.toHaveBeenCalled();
    editor.destroy();
  });

  it("does nothing before a document opens", () => {
    const { host, dispatch, togglePane } = fakeHost(null);
    new RevisionsCommands(host).onBalloonSelect({ kind: "revision", id: 5 });
    expect(dispatch).not.toHaveBeenCalled();
    expect(togglePane).not.toHaveBeenCalled();
  });
});
