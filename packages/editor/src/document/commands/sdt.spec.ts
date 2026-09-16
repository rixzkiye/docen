import { docxExtensions } from "@docen/docx";
import { Editor } from "@tiptap/core";
import { describe, expect, it } from "vitest";

import { SdtCommands, SdtHostCommands } from "./host/sdt";

function makeTestEditor(initialText = "Hello world") {
  return new Editor({
    extensions: docxExtensions,
    content: {
      type: "doc",
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: initialText }],
        },
      ],
    },
  });
}

describe("Content Controls (SDT) - Wave H", () => {
  it("inserts inline checkbox content control and toggles it", () => {
    const editor = makeTestEditor();
    const host = {
      editor: () => editor,
      bridge: () => undefined,
      element: () => document.createElement("div"),
      rerender: () => {},
    };

    const sdt = new SdtCommands(host);
    sdt.insertSdt("checkbox");

    const json = editor.getJSON() as Record<string, any>;
    const para = json.content?.[0];
    const sdtNode = para?.content?.find((c: any) => c.type === "sdtInline");
    expect(sdtNode).toBeDefined();
    expect(sdtNode?.attrs?.properties?.tag).toBe("Checkbox");
    expect(sdtNode?.content?.[0]?.text).toBe("☐ ");

    // Move selection into checkbox
    editor.commands.setTextSelection(3);
    const toggled = sdt.toggleCheckboxAtCaret();
    expect(toggled).toBe(true);

    const afterToggle = editor.getJSON() as Record<string, any>;
    const toggledSdt = afterToggle.content?.[0]?.content?.find((c: any) => c.type === "sdtInline");
    expect(toggledSdt?.attrs?.properties?.checkbox?.checked).toBe(1);
    expect(toggledSdt?.content?.[0]?.text).toBe("☒ ");
  });

  it("inserts dropdown and date content controls with properties", () => {
    const editor = makeTestEditor();
    const host = {
      editor: () => editor,
      bridge: () => undefined,
      element: () => document.createElement("div"),
      rerender: () => {},
    };

    const sdt = new SdtCommands(host);
    sdt.insertSdt("date");
    let json = editor.getJSON() as Record<string, any>;
    let sdtNode = json.content?.[0]?.content?.find((c: any) => c.type === "sdtInline");
    expect(sdtNode?.attrs?.properties?.date?.dateFormat).toBe("YYYY-MM-DD");

    sdt.insertSdt("dropdown");
    json = editor.getJSON() as Record<string, any>;
    const ddNode = json.content?.[0]?.content?.filter((c: any) => c.type === "sdtInline")?.[1];
    expect(ddNode?.attrs?.properties?.dropDownList?.listItems?.length).toBeGreaterThan(0);
  });

  it("dispatches SDT ribbon commands through SdtHostCommands", () => {
    const editor = makeTestEditor();
    const host = {
      editor: () => editor,
      bridge: () => undefined,
      element: () => document.createElement("div"),
      rerender: () => {},
    };

    const sdt = new SdtCommands(host);
    let propertiesOpened = false;
    const hostCommands = new SdtHostCommands(sdt, () => {
      propertiesOpened = true;
    });

    expect(hostCommands.dispatch("sdt-checkbox")).toBe(true);
    expect(hostCommands.dispatch("toggle-design-mode")).toBe(true);
    expect(sdt.isDesignMode()).toBe(true);
    expect(hostCommands.dispatch("sdt-properties")).toBe(true);
    expect(propertiesOpened).toBe(true);
  });
});
