// Built-in document templates for the "New from Template" gallery — data only,
// no UI: each entry pairs gallery metadata (an i18n name/description key) with
// a fresh Tiptap model-JSON builder. Instantiation is the normal #loadDoc
// pipeline (the model is plain Tiptap JSON, normalized with the engine's
// document defaults like any hand-built doc), so nothing here touches the DOM.

import type { JSONContent } from "@docen/docx";

/** Locales a built-in template body is written in. */
export type TemplateLocale = "en" | "zh-CN";

/** One built-in template: gallery metadata + the model-JSON builder. `build()`
 *  returns a fresh tree on every call, so an instantiated document never
 *  aliases the template data (editing it must not mutate the template). */
export interface DocxTemplate {
  /** Stable id the gallery's create action carries (no i18n — data key). */
  readonly id: string;
  /** i18n key for the gallery row's name. */
  readonly nameKey: string;
  /** i18n key for the gallery row's one-line description. */
  readonly descriptionKey: string;
  build(locale?: TemplateLocale): JSONContent;
}

/** Map an element/attribute language tag to a template-body locale. */
export function templateLocale(lang: string | null | undefined): TemplateLocale {
  return (lang ?? "").toLowerCase().startsWith("zh") ? "zh-CN" : "en";
}

/** Body copy for the built-in templates, per locale. Data, not UI chrome —
 *  the gallery's own labels live in the editor i18n tables. */
interface TemplateStrings {
  readonly reportTitle: string;
  readonly reportByline: string;
  readonly reportIntroHeading: string;
  readonly reportIntroBody: string;
  readonly reportAnalysisHeading: string;
  readonly reportAnalysisBody: string;
  readonly reportConclusionHeading: string;
  readonly reportConclusionBody: string;
  readonly letterSender: string;
  readonly letterAddress: string;
  readonly letterDate: string;
  readonly letterRecipient: string;
  readonly letterSalutation: string;
  readonly letterBody: string;
  readonly letterSecondBody: string;
  readonly letterClosing: string;
  readonly letterName: string;
}

const TEMPLATE_STRINGS: Record<TemplateLocale, TemplateStrings> = {
  en: {
    reportTitle: "Report Title",
    reportByline: "Author · Department · Date",
    reportIntroHeading: "Introduction",
    reportIntroBody:
      "Introduce the topic and state the purpose of the report. Replace this placeholder text with your own content.",
    reportAnalysisHeading: "Analysis",
    reportAnalysisBody:
      "Present the findings, data, and reasoning here. Add tables or figures as needed.",
    reportConclusionHeading: "Conclusion",
    reportConclusionBody: "Summarize the key points and outline recommended next steps.",
    letterSender: "Your Name",
    letterAddress: "Street Address · City · Postal Code",
    letterDate: "Date",
    letterRecipient: "Recipient Name\nCompany\nStreet Address · City · Postal Code",
    letterSalutation: "Dear Recipient Name,",
    letterBody:
      "Open with the purpose of your letter. Replace this placeholder text with your own content.",
    letterSecondBody: "Continue with the details, and close with any action you request.",
    letterClosing: "Sincerely,",
    letterName: "Your Name",
  },
  "zh-CN": {
    reportTitle: "报告标题",
    reportByline: "作者 · 部门 · 日期",
    reportIntroHeading: "引言",
    reportIntroBody: "介绍主题并说明报告目的。请将此处占位文字替换为实际内容。",
    reportAnalysisHeading: "分析",
    reportAnalysisBody: "在此呈现调查结果、数据与论证，可按需插入表格或图表。",
    reportConclusionHeading: "结论",
    reportConclusionBody: "总结要点，并列出建议的后续行动。",
    letterSender: "您的姓名",
    letterAddress: "街道地址 · 城市 · 邮政编码",
    letterDate: "日期",
    letterRecipient: "收件人姓名\n公司名称\n街道地址 · 城市 · 邮政编码",
    letterSalutation: "尊敬的收件人：",
    letterBody: "开篇说明写信目的。请将此处占位文字替换为实际内容。",
    letterSecondBody: "继续补充具体事项，并在结尾提出您期望的行动。",
    letterClosing: "此致",
    letterName: "您的姓名",
  },
};

/** One paragraph node — plain text when given, else an empty textblock. */
function p(text?: string, attrs?: Record<string, unknown>): JSONContent {
  const node: JSONContent = { type: "paragraph" };
  if (attrs) node.attrs = attrs;
  if (text) node.content = [{ type: "text", text }];
  return node;
}

/** Multi-line placeholder text as one paragraph per line (a single text node
 *  with hard breaks would need the hardBreak atom; separate blocks are simpler
 *  for address-style placeholders). */
function lines(text: string, attrs?: Record<string, unknown>): JSONContent[] {
  return text.split("\n").map((line) => p(line, attrs));
}

/** The report's placeholder TOC — the same field switches the editor's own
 *  TOC insert stamps (heading levels 1-3, hyperlinked entries), with an empty
 *  entry paragraph until Update Field runs. */
function placeholderToc(): JSONContent {
  return {
    type: "tocField",
    attrs: { options: { headingStyleRange: "1-3", hyperlink: true } },
    content: [p()],
  };
}

/** The built-in gallery, in display order. Bodies are localized at build time
 *  via {@link templateLocale}; add new entries here — the dialog renders this
 *  list, it does not hardcode templates. */
export const BUILTIN_TEMPLATES: readonly DocxTemplate[] = [
  {
    id: "blank",
    nameKey: "template.name.blank",
    descriptionKey: "template.desc.blank",
    build: () => ({ type: "doc", content: [p()] }),
  },
  {
    id: "report",
    nameKey: "template.name.report",
    descriptionKey: "template.desc.report",
    build: (locale = "en") => {
      const s = TEMPLATE_STRINGS[locale];
      return {
        type: "doc",
        content: [
          p(s.reportTitle, { heading: "Title" }),
          p(s.reportByline),
          placeholderToc(),
          p(s.reportIntroHeading, { heading: "Heading1" }),
          p(s.reportIntroBody),
          p(s.reportAnalysisHeading, { heading: "Heading1" }),
          p(s.reportAnalysisBody),
          p(s.reportConclusionHeading, { heading: "Heading1" }),
          p(s.reportConclusionBody),
        ],
      };
    },
  },
  {
    id: "letter",
    nameKey: "template.name.letter",
    descriptionKey: "template.desc.letter",
    build: (locale = "en") => {
      const s = TEMPLATE_STRINGS[locale];
      return {
        type: "doc",
        content: [
          ...lines(s.letterSender),
          ...lines(s.letterAddress),
          p(),
          ...lines(s.letterDate),
          p(),
          ...lines(s.letterRecipient),
          p(),
          p(s.letterSalutation),
          p(s.letterBody),
          p(s.letterSecondBody),
          p(),
          ...lines(s.letterClosing),
          ...lines(s.letterName),
        ],
      };
    },
  },
];

/** Look up a built-in template by its stable id. */
export function findTemplate(id: string): DocxTemplate | undefined {
  return BUILTIN_TEMPLATES.find((tpl) => tpl.id === id);
}
