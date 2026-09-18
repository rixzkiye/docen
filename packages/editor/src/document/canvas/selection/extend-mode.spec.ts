import { Schema } from "@tiptap/pm/model";
import { EditorState, TextSelection } from "@tiptap/pm/state";
import { describe, expect, it } from "vitest";

import { ExtendLevel, ExtendModeManager } from "./extend-mode";

describe("ExtendModeManager (W2.4 F8 extend mode)", () => {
  const schema = new Schema({
    nodes: {
      doc: { content: "paragraph+" },
      paragraph: { content: "text*", toDOM: () => ["p", 0] },
      text: { inline: true },
    },
  });

  const createDoc = (text: string) => {
    return schema.node("doc", null, [schema.node("paragraph", null, [schema.text(text)])]);
  };

  it("cycles through Word F8 extend levels: Active -> Word -> Sentence -> Paragraph -> Document", () => {
    const doc = createDoc("Hello beautiful world. This is sentence two.");
    let state = EditorState.create({
      doc,
      schema,
      selection: TextSelection.create(doc, 8), // inside "beautiful"
    });

    const mockView = {
      dispatch: (tr: any) => {
        state = state.apply(tr);
      },
    };
    const mockEditor = {
      get state() {
        return state;
      },
      get view() {
        return mockView;
      },
    } as any;

    const mgr = new ExtendModeManager();
    expect(mgr.isActive).toBe(false);
    expect(mgr.statusLabel()).toBe("");

    // 1st F8: Arm extend mode
    const lvl1 = mgr.step(mockEditor);
    expect(lvl1).toBe(ExtendLevel.Active);
    expect(mgr.isActive).toBe(true);
    expect(mgr.statusLabel()).toBe("EXT");
    expect(mgr.anchor).toBe(8);

    // 2nd F8: Select Word ("beautiful")
    const lvl2 = mgr.step(mockEditor);
    expect(lvl2).toBe(ExtendLevel.Word);
    expect(state.selection.from).toBe(7);
    expect(state.selection.to).toBe(16);
    expect(doc.textBetween(state.selection.from, state.selection.to)).toBe("beautiful");

    // 3rd F8: Select Sentence
    const lvl3 = mgr.step(mockEditor);
    expect(lvl3).toBe(ExtendLevel.Sentence);
    expect(state.selection.from).toBe(1);
    expect(state.selection.to).toBe(23);
    expect(doc.textBetween(state.selection.from, state.selection.to)).toBe(
      "Hello beautiful world.",
    );

    // 4th F8: Select Paragraph
    const lvl4 = mgr.step(mockEditor);
    expect(lvl4).toBe(ExtendLevel.Paragraph);
    expect(state.selection.from).toBe(1);
    expect(state.selection.to).toBe(45);

    // 5th F8: Select Document
    const lvl5 = mgr.step(mockEditor);
    expect(lvl5).toBe(ExtendLevel.Document);
    expect(state.selection.from).toBe(1);
    expect(state.selection.to).toBe(45);

    // Shift+F8: Shrink back down
    const shrunk = mgr.shrink(mockEditor);
    expect(shrunk).toBe(ExtendLevel.Paragraph);
    expect(state.selection.to).toBe(45);

    // Escape: Cancel
    mgr.cancel();
    expect(mgr.isActive).toBe(false);
    expect(mgr.statusLabel()).toBe("");
  });
});
