import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";

export type SdtType =
  | "plainText"
  | "richText"
  | "checkbox"
  | "dropdown"
  | "combo"
  | "date"
  | "picture"
  | "buildingBlock";

export interface SdtHost {
  editor(): Editor | null | undefined;
  bridge(): { activeEditor(): Editor; focus(): void } | undefined;
  element(): HTMLElement;
  rerender(): void;
}

export class SdtCommands {
  #designMode = false;

  constructor(private readonly host: SdtHost) {}

  #target(): Editor | null | undefined {
    return this.host.bridge()?.activeEditor() ?? this.host.editor();
  }

  isDesignMode(): boolean {
    return this.#designMode;
  }

  toggleDesignMode(): boolean {
    this.#designMode = !this.#designMode;
    this.host.rerender();
    return this.#designMode;
  }

  /**
   * Insert a content control (Structured Document Tag) at the current caret.
   * Matches Word's Developer tab control insertion behavior.
   */
  insertSdt(type: SdtType): void {
    const editor = this.#target();
    if (!editor) return;

    let sdtNode: JSONContent;

    switch (type) {
      case "checkbox":
        sdtNode = {
          type: "sdtInline",
          attrs: {
            properties: {
              tag: "Checkbox",
              alias: "Check Box",
              checkbox: {
                checked: 0,
                checkedState: { font: "Wingdings", val: "F0FE" },
                uncheckedState: { font: "Wingdings", val: "F0A8" },
              },
            },
          },
          content: [{ type: "text", text: "☐ " }],
        };
        break;

      case "date":
        sdtNode = {
          type: "sdtInline",
          attrs: {
            properties: {
              tag: "Date",
              alias: "Date Picker",
              date: {
                dateFormat: "YYYY-MM-DD",
                lid: "1033",
              },
            },
          },
          content: [{ type: "text", text: "Click here to enter a date." }],
        };
        break;

      case "dropdown":
      case "combo":
        sdtNode = {
          type: "sdtInline",
          attrs: {
            properties: {
              tag: "Dropdown",
              alias: "Drop-Down List",
              dropDownList: {
                listItems: [
                  { displayText: "Choose an item.", value: "" },
                  { displayText: "Option 1", value: "opt1" },
                  { displayText: "Option 2", value: "opt2" },
                ],
              },
            },
          },
          content: [{ type: "text", text: "Choose an item." }],
        };
        break;

      case "picture":
        sdtNode = {
          type: "sdtBlock",
          attrs: {
            properties: {
              tag: "Picture",
              alias: "Picture Content Control",
            },
          },
          content: [
            {
              type: "paragraph",
              content: [{ type: "text", text: "Click here to insert a picture." }],
            },
          ],
        };
        break;

      case "plainText":
        sdtNode = {
          type: "sdtInline",
          attrs: {
            properties: {
              tag: "PlainText",
              alias: "Plain Text",
            },
          },
          content: [{ type: "text", text: "Click or tap here to enter text." }],
        };
        break;

      case "buildingBlock":
        sdtNode = {
          type: "sdtBlock",
          attrs: {
            properties: {
              tag: "BuildingBlock",
              alias: "Building Block Gallery",
            },
          },
          content: [{ type: "paragraph", content: [{ type: "text", text: "Quick Part" }] }],
        };
        break;

      case "richText":
      default:
        sdtNode = {
          type: "sdtBlock",
          attrs: {
            properties: {
              tag: "RichText",
              alias: "Rich Text",
            },
          },
          content: [
            {
              type: "paragraph",
              content: [{ type: "text", text: "Click or tap here to enter text." }],
            },
          ],
        };
        break;
    }

    const { from } = editor.state.selection;
    const node = editor.schema.nodeFromJSON(sdtNode);
    editor.view.dispatch(editor.state.tr.insert(from, node));
    this.host.bridge()?.focus();
  }

  /**
   * Toggle a checkbox content control between checked (☒) and unchecked (☐).
   */
  toggleCheckboxAtCaret(): boolean {
    const editor = this.#target();
    if (!editor) return false;

    const { $from } = editor.state.selection;
    // Check if within sdtInline with checkbox properties
    for (let depth = $from.depth; depth > 0; depth--) {
      const node = $from.node(depth);
      if (node.type.name === "sdtInline") {
        const props = (node.attrs.properties as Record<string, unknown> | undefined) ?? {};
        if (props.checkbox || node.text?.includes("☐") || node.text?.includes("☒")) {
          const currentChecked =
            (props.checkbox as { checked?: number } | undefined)?.checked === 1;
          const newChecked = !currentChecked;
          const pos = $from.before(depth);

          const newProps = {
            ...props,
            checkbox: {
              ...(props.checkbox as object),
              checked: newChecked ? 1 : 0,
            },
          };

          const newText = newChecked ? "☒ " : "☐ ";
          const tr = editor.state.tr;
          tr.setNodeMarkup(pos, undefined, { ...node.attrs, properties: newProps });
          // Replace inner text
          tr.replaceWith(pos + 1, pos + node.nodeSize - 1, editor.schema.text(newText));
          editor.view.dispatch(tr);
          return true;
        }
      }
    }
    return false;
  }
}

export class SdtHostCommands {
  readonly events = [
    "sdt-rich-text",
    "sdt-plain-text",
    "sdt-picture",
    "sdt-checkbox",
    "sdt-combo-box",
    "sdt-dropdown",
    "sdt-date",
    "sdt-building-block",
    "toggle-design-mode",
    "sdt-properties",
  ] as const;

  constructor(
    private readonly sdt: SdtCommands,
    private readonly openProperties?: () => void,
  ) {}

  dispatch(event: string): boolean {
    switch (event) {
      case "sdt-rich-text":
        this.sdt.insertSdt("richText");
        return true;
      case "sdt-plain-text":
        this.sdt.insertSdt("plainText");
        return true;
      case "sdt-picture":
        this.sdt.insertSdt("picture");
        return true;
      case "sdt-checkbox":
        this.sdt.insertSdt("checkbox");
        return true;
      case "sdt-combo-box":
        this.sdt.insertSdt("combo");
        return true;
      case "sdt-dropdown":
        this.sdt.insertSdt("dropdown");
        return true;
      case "sdt-date":
        this.sdt.insertSdt("date");
        return true;
      case "sdt-building-block":
        this.sdt.insertSdt("buildingBlock");
        return true;
      case "toggle-design-mode":
        this.sdt.toggleDesignMode();
        return true;
      case "sdt-properties":
        this.openProperties?.();
        return true;
      default:
        return false;
    }
  }
}
