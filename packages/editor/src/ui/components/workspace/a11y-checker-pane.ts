import type { JSONContent } from "@docen/docx";
import { FASTElement, css, customElement, html, observable, ref } from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

export type A11ySeverity = "error" | "warning" | "tip";

export interface A11yIssue {
  id: string;
  severity: A11ySeverity;
  rule: string;
  message: string;
  nodeIndex?: number;
  nodeType: string;
  params?: {
    from?: number;
    to?: number;
    text?: string;
  };
}

/**
 * Scan a document for accessibility issues according to Word Parity rules.
 */
export function scanA11yIssues(doc: JSONContent | null | undefined): A11yIssue[] {
  if (!doc || !doc.content) return [];
  const issues: A11yIssue[] = [];

  let lastHeadingLevel = 0;
  let nodeIdx = 0;

  function traverse(node: JSONContent): void {
    const currentIdx = nodeIdx++;

    // 1. Missing alt text on images
    if (node.type === "image") {
      const alt = (node.attrs?.alt as string | undefined)?.trim();
      if (!alt) {
        issues.push({
          id: `img-alt-${currentIdx}`,
          severity: "error",
          rule: "alt-text",
          message: "Missing alternative text on image",
          nodeIndex: currentIdx,
          nodeType: "image",
        });
      }
    }

    // 2. Heading level jumps
    if (node.type === "heading" || (node.type === "paragraph" && node.attrs?.level)) {
      const level = Number(node.attrs?.level ?? 1);
      if (lastHeadingLevel > 0 && level > lastHeadingLevel + 1) {
        issues.push({
          id: `heading-jump-${currentIdx}`,
          severity: "warning",
          rule: "heading-order",
          message: `Heading level jumped from H${lastHeadingLevel} to H${level}`,
          params: { from: lastHeadingLevel, to: level },
          nodeIndex: currentIdx,
          nodeType: "heading",
        });
      }
      lastHeadingLevel = level;
    }

    // 3. Table without header row
    if (node.type === "table") {
      const firstRow = node.content?.[0];
      const hasHeader =
        firstRow && firstRow.content?.some((c) => c.type === "tableHeader" || c.attrs?.isHeader);
      if (!hasHeader) {
        issues.push({
          id: `table-header-${currentIdx}`,
          severity: "warning",
          rule: "table-header",
          message: "Table does not specify a header row",
          nodeIndex: currentIdx,
          nodeType: "table",
        });
      }
    }

    // 4. Ambiguous hyperlink text
    if (node.type === "text" && Array.isArray(node.marks)) {
      const linkMark = node.marks.find((m) => m.type === "link");
      if (linkMark && node.text) {
        const lower = node.text.trim().toLowerCase();
        if (["click here", "here", "read more", "link", "more"].includes(lower)) {
          issues.push({
            id: `link-text-${currentIdx}`,
            severity: "tip",
            rule: "link-text",
            message: `Ambiguous link text "${node.text}"`,
            params: { text: node.text },
            nodeIndex: currentIdx,
            nodeType: "link",
          });
        }
      }
    }

    if (node.content && Array.isArray(node.content)) {
      for (const child of node.content) {
        traverse(child);
      }
    }
  }

  for (const child of doc.content) {
    traverse(child);
  }

  return issues;
}

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    height: 100%;
    padding: 12px;
    box-sizing: border-box;
    font-size: 12px;
    color: var(--docen-color-foreground, #242424);
    overflow-y: auto;
    gap: 12px;
  }
  .header-box {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
    padding-bottom: 8px;
  }
  .title {
    font-weight: 600;
    font-size: 13px;
  }
  .status-all-good {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 12px;
    background: #def6e3;
    color: #0e700e;
    border-radius: 4px;
    font-weight: 500;
  }
  .issue-group {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .group-title {
    font-weight: 600;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .issue-card {
    padding: 8px 10px;
    border: 1px solid var(--docen-color-divider, #e0e0e0);
    border-radius: 4px;
    background: var(--docen-color-canvas, #ffffff);
    display: flex;
    flex-direction: column;
    gap: 4px;
    cursor: pointer;
  }
  .issue-card:hover {
    background: var(--colorNeutralBackground1Hover, #f5f5f5);
  }
  .issue-card.error {
    border-left: 3px solid #d13438;
  }
  .issue-card.warning {
    border-left: 3px solid #ffaa44;
  }
  .issue-card.tip {
    border-left: 3px solid #0078d4;
  }
  .issue-msg {
    font-weight: 500;
  }
  .issue-sub {
    font-size: 11px;
    color: var(--docen-color-foreground-secondary, #616161);
  }
`;

const template = html<DocenA11yCheckerPane>`
  <div class="header-box">
    <span class="title">${(x) => t("a11y.checkerTitle", x)}</span>
    <fluent-button size="small" @click="${(x) => x.refresh()}">
      ${(x) => t("a11y.recheck", x)}
    </fluent-button>
  </div>

  ${(x) =>
    x.issues.length === 0
      ? html<DocenA11yCheckerPane>`
          <div class="status-all-good"><span>✓</span> ${(p) => t("a11y.noIssuesFound", p)}</div>
        `
      : html<DocenA11yCheckerPane>` <div class="issue-group" ${ref("listEl")}></div> `}
`;

@customElement({ name: "docen-a11y-checker-pane", template, styles })
export class DocenA11yCheckerPane extends FASTElement {
  @observable issues: A11yIssue[] = [];
  @observable listEl?: HTMLElement;

  #doc?: JSONContent | null;
  #unsubscribe?: () => void;

  override connectedCallback(): void {
    super.connectedCallback();
    this.#unsubscribe = observeLang(() => this.#renderList());
  }

  override disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  formatMessage(issue: A11yIssue): string {
    if (issue.rule === "alt-text") {
      return t("a11y.missingAltText", this);
    }
    if (issue.rule === "heading-order") {
      const from = String(issue.params?.from ?? 1);
      const to = String(issue.params?.to ?? 2);
      const tpl = t("a11y.headingJump", this);
      return tpl.replace("{from}", from).replace("{to}", to);
    }
    if (issue.rule === "table-header") {
      return t("a11y.missingTableHeader", this);
    }
    if (issue.rule === "link-text") {
      const text = String(issue.params?.text ?? "");
      const tpl = t("a11y.ambiguousLink", this);
      return tpl.replace("{text}", text);
    }
    return issue.message;
  }

  check(doc: JSONContent | null | undefined): void {
    this.#doc = doc;
    this.issues = scanA11yIssues(doc);
    this.#renderList();
  }

  refresh(): void {
    this.$emit("a11y:refresh");
    this.check(this.#doc);
  }

  #renderList(): void {
    if (!this.listEl) return;
    this.listEl.replaceChildren();

    for (const issue of this.issues) {
      const card = document.createElement("div");
      card.className = `issue-card ${issue.severity}`;

      const msg = document.createElement("div");
      msg.className = "issue-msg";
      msg.textContent = this.formatMessage(issue);

      const sub = document.createElement("div");
      sub.className = "issue-sub";
      sub.textContent = `${issue.nodeType} • ${issue.rule}`;

      card.appendChild(msg);
      card.appendChild(sub);

      card.addEventListener("click", () => {
        this.$emit("a11y:select-issue", { issue });
      });

      this.listEl.appendChild(card);
    }
  }
}
