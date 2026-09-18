import type { FormFieldOptions, ParagraphChild } from "@office-open/docx";

import { Node } from "../core";
import type { ParseInlineRule } from "./types";
import { attrNative } from "./utils";

type FormFieldBranch = Extract<ParagraphChild, { formField: unknown }>;

export function extractFormFieldText(ff: FormFieldOptions): string {
  if (ff.checkBox) {
    return ff.checkBox.checked ? "☒" : "☐";
  }
  if (ff.dropDownList) {
    const idx = ff.dropDownList.result ?? ff.dropDownList.default ?? 0;
    return ff.dropDownList.entries?.[idx] ?? "";
  }
  if (ff.textInput) {
    return ff.textInput.value ?? ff.textInput.default ?? "";
  }
  return "";
}

export const parseFormFieldInline: ParseInlineRule<FormFieldBranch> = {
  match: (child): child is FormFieldBranch => "formField" in child,
  convert: (child) => {
    const ff = child.formField as FormFieldOptions;
    const text = extractFormFieldText(ff);
    return {
      type: "formField",
      attrs: { formField: ff },
      content: text ? [{ type: "text", text }] : [{ type: "text", text: "" }],
    };
  },
};

export const FormField = Node.create({
  name: "formField",
  group: "inline",
  inline: true,
  content: "inline*",
  isolating: true,
  defining: true,

  addAttributes() {
    return {
      formField: attrNative({}),
    };
  },

  parseHTML() {
    return [
      {
        tag: "span.docx-form-field",
        getAttrs: (el) => {
          const raw = (el as HTMLElement).getAttribute("data-form-field");
          return { formField: raw ? JSON.parse(raw) : {} };
        },
      },
    ];
  },

  renderHTML({ HTMLAttributes }) {
    return [
      "span",
      {
        class: "docx-form-field",
        "data-form-field": JSON.stringify(HTMLAttributes.formField ?? {}),
      },
      0,
    ];
  },

  parseDocxInline: parseFormFieldInline,
});
