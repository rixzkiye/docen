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
  readonly resumeName: string;
  readonly resumeContact: string;
  readonly resumeSummaryHeading: string;
  readonly resumeSummaryBody: string;
  readonly resumeExperienceHeading: string;
  readonly resumeJobTitle: string;
  readonly resumeJobCompany: string;
  readonly resumeJobDesc: string;
  readonly resumeEducationHeading: string;
  readonly resumeDegree: string;
  readonly resumeSkillsHeading: string;
  readonly resumeSkillsBody: string;
  readonly invoiceTitle: string;
  readonly invoiceNumber: string;
  readonly invoiceDate: string;
  readonly invoiceBillTo: string;
  readonly invoiceClient: string;
  readonly invoiceItem: string;
  readonly invoiceQty: string;
  readonly invoicePrice: string;
  readonly invoiceTotalHeader: string;
  readonly invoiceSubtotal: string;
  readonly invoiceTax: string;
  readonly invoiceTotal: string;
  readonly invoiceTerms: string;
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
    resumeName: "Alex Morgan",
    resumeContact:
      "alex.morgan@example.com · +1 (555) 019-2834 · San Francisco, CA · linkedin.com/in/alexmorgan",
    resumeSummaryHeading: "Professional Summary",
    resumeSummaryBody:
      "Results-driven software engineer with over 8 years of experience designing, developing, and deploying scalable distributed systems and cloud native applications.",
    resumeExperienceHeading: "Work Experience",
    resumeJobTitle: "Senior Software Engineer — Cloud Infrastructure",
    resumeJobCompany: "Acme Corporation · 2021 – Present",
    resumeJobDesc:
      "Led architectural design of high-throughput real-time processing pipelines reducing latency by 45%. Mentored 6 engineers and drove engineering excellence across teams.",
    resumeEducationHeading: "Education",
    resumeDegree: "B.S. in Computer Science — University of California, Berkeley (2016 – 2020)",
    resumeSkillsHeading: "Skills & Proficiencies",
    resumeSkillsBody:
      "TypeScript, Rust, Go, React, Distributed Systems, WebAssembly, CI/CD, Cloud Architecture",
    invoiceTitle: "INVOICE",
    invoiceNumber: "Invoice #: INV-2026-001\nDate: September 18, 2026\nDue Date: October 18, 2026",
    invoiceDate: "Date: September 18, 2026",
    invoiceBillTo: "Billed To:",
    invoiceClient: "Contoso Ltd.\n100 Enterprise Way\nSuite 400\nSeattle, WA 98101",
    invoiceItem: "Item Description",
    invoiceQty: "Qty",
    invoicePrice: "Unit Price",
    invoiceTotalHeader: "Amount",
    invoiceSubtotal: "Subtotal: $8,400.00",
    invoiceTax: "Tax (10%): $840.00",
    invoiceTotal: "Total Due: $9,240.00",
    invoiceTerms: "Payment Terms: Due within 30 days of invoice date. Thank you for your business!",
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
    resumeName: "张伟",
    resumeContact: "zhangwei@example.com · 138-0000-0000 · 北京市海淀区 · github.com/zhangwei",
    resumeSummaryHeading: "个人总结",
    resumeSummaryBody:
      "具备8年以上全栈及分布式系统研发经验的资深工程师，擅长高性能架构设计、大型前端与文档处理系统工程化。",
    resumeExperienceHeading: "工作经历",
    resumeJobTitle: "资深软件工程师 — 基础架构部",
    resumeJobCompany: "某知名科技有限公司 · 2021 – 至今",
    resumeJobDesc:
      "负责核心文档渲染与协同引擎研发，优化首屏排版性能达50%以上；带领敏捷开发团队高质量完成多次重大版本迭代。",
    resumeEducationHeading: "教育背景",
    resumeDegree: "计算机科学与技术 学士 — 清华大学 (2016 – 2020)",
    resumeSkillsHeading: "专业技能",
    resumeSkillsBody: "TypeScript、Rust、Go、React、分布式系统、WebAssembly、云原生微服务",
    invoiceTitle: "账 单 / 发 票",
    invoiceNumber: "账单编号：INV-2026-001\n开票日期：2026年9月18日\n到期日期：2026年10月18日",
    invoiceDate: "开票日期：2026年9月18日",
    invoiceBillTo: "客户信息：",
    invoiceClient: "北京某某科技有限公司\n北京市朝阳区科技园区8号楼\n邮编：100000",
    invoiceItem: "项目描述",
    invoiceQty: "数量",
    invoicePrice: "单价",
    invoiceTotalHeader: "金额",
    invoiceSubtotal: "小计：¥8,400.00",
    invoiceTax: "税额 (10%)：¥840.00",
    invoiceTotal: "应付总额：¥9,240.00",
    invoiceTerms: "付款说明：请在开票之日起30天内完成付款。感谢您的合作！",
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
  {
    id: "resume",
    nameKey: "template.name.resume",
    descriptionKey: "template.desc.resume",
    build: (locale = "en") => {
      const s = TEMPLATE_STRINGS[locale];
      return {
        type: "doc",
        content: [
          p(s.resumeName, { heading: "Title" }),
          p(s.resumeContact),
          p(),
          p(s.resumeSummaryHeading, { heading: "Heading1" }),
          p(s.resumeSummaryBody),
          p(),
          p(s.resumeExperienceHeading, { heading: "Heading1" }),
          p(s.resumeJobTitle, { heading: "Heading2" }),
          p(s.resumeJobCompany),
          p(s.resumeJobDesc),
          p(),
          p(s.resumeEducationHeading, { heading: "Heading1" }),
          p(s.resumeDegree),
          p(),
          p(s.resumeSkillsHeading, { heading: "Heading1" }),
          p(s.resumeSkillsBody),
        ],
      };
    },
  },
  {
    id: "invoice",
    nameKey: "template.name.invoice",
    descriptionKey: "template.desc.invoice",
    build: (locale = "en") => {
      const s = TEMPLATE_STRINGS[locale];
      return {
        type: "doc",
        content: [
          p(s.invoiceTitle, { heading: "Title" }),
          ...lines(s.invoiceNumber),
          p(),
          p(s.invoiceBillTo, { heading: "Heading2" }),
          ...lines(s.invoiceClient),
          p(),
          {
            type: "table",
            content: [
              {
                type: "tableRow",
                content: [
                  { type: "tableCell", content: [p(s.invoiceItem)] },
                  { type: "tableCell", content: [p(s.invoiceQty)] },
                  { type: "tableCell", content: [p(s.invoicePrice)] },
                  { type: "tableCell", content: [p(s.invoiceTotalHeader)] },
                ],
              },
              {
                type: "tableRow",
                content: [
                  { type: "tableCell", content: [p("Web Application Development")] },
                  { type: "tableCell", content: [p("40 hrs")] },
                  { type: "tableCell", content: [p("$125.00")] },
                  { type: "tableCell", content: [p("$5,000.00")] },
                ],
              },
              {
                type: "tableRow",
                content: [
                  { type: "tableCell", content: [p("Cloud Architecture Consulting")] },
                  { type: "tableCell", content: [p("15 hrs")] },
                  { type: "tableCell", content: [p("$160.00")] },
                  { type: "tableCell", content: [p("$2,400.00")] },
                ],
              },
              {
                type: "tableRow",
                content: [
                  { type: "tableCell", content: [p("System Optimization & Testing")] },
                  { type: "tableCell", content: [p("10 hrs")] },
                  { type: "tableCell", content: [p("$100.00")] },
                  { type: "tableCell", content: [p("$1,000.00")] },
                ],
              },
            ],
          },
          p(),
          p(s.invoiceSubtotal),
          p(s.invoiceTax),
          p(s.invoiceTotal, { heading: "Heading2" }),
          p(),
          p(s.invoiceTerms),
        ],
      };
    },
  },
];

/** Look up a built-in template by its stable id. */
export function findTemplate(id: string): DocxTemplate | undefined {
  return BUILTIN_TEMPLATES.find((tpl) => tpl.id === id);
}
