import type { LayoutBlock, LayoutInline, LayoutTable } from "@docen/layout";

export interface LayoutDocInput {
  sections: Array<{ blocks: LayoutBlock[] }>;
}

/**
 * A11y DOM Mirror — renders an off-screen semantic HTML tree strictly derived from LayoutDoc.
 * Single source of truth is LayoutDoc (no DOM measurement, no layout feedback into engine).
 * Also manages an aria-live region for announcements (page navigation, mode changes, revisions).
 */
export class A11yMirror {
  readonly root: HTMLElement;
  readonly contentContainer: HTMLElement;
  readonly liveRegion: HTMLElement;

  constructor() {
    this.root = document.createElement("div");
    this.root.className = "docen-a11y-mirror-root";
    // Visually hidden but accessible to screen readers
    Object.assign(this.root.style, {
      position: "absolute",
      width: "1px",
      height: "1px",
      padding: "0",
      margin: "-1px",
      overflow: "hidden",
      clip: "rect(0, 0, 0, 0)",
      whiteSpace: "normal",
      border: "0",
    });

    this.contentContainer = document.createElement("div");
    this.contentContainer.setAttribute("role", "document");
    this.contentContainer.setAttribute("aria-label", "Document Content");
    this.root.appendChild(this.contentContainer);

    this.liveRegion = document.createElement("div");
    this.liveRegion.setAttribute("role", "status");
    this.liveRegion.setAttribute("aria-live", "polite");
    this.liveRegion.setAttribute("aria-atomic", "true");
    this.root.appendChild(this.liveRegion);
  }

  /**
   * Announce an accessibility message to screen readers via aria-live.
   */
  announce(message: string): void {
    if (!message) return;
    this.liveRegion.textContent = "";
    // Trigger mutation for screen readers
    setTimeout(() => {
      this.liveRegion.textContent = message;
    }, 50);
  }

  /**
   * Update the hidden semantic DOM tree from a LayoutDoc.
   */
  update(doc: LayoutDocInput | null | undefined): void {
    this.contentContainer.replaceChildren();
    if (!doc) return;

    for (let sIdx = 0; sIdx < doc.sections.length; sIdx++) {
      const section = doc.sections[sIdx];
      const sectionEl = document.createElement("section");
      sectionEl.setAttribute("aria-label", `Section ${sIdx + 1}`);

      for (const block of section.blocks) {
        const el = this.renderBlock(block);
        if (el) sectionEl.appendChild(el);
      }

      this.contentContainer.appendChild(sectionEl);
    }
  }

  private renderBlock(block: LayoutBlock): HTMLElement | null {
    if (block.kind === "paragraph") {
      // Check if paragraph has text
      const inlineEls = block.inline.map((inItem) => this.renderInline(inItem)).filter(Boolean);
      const p = document.createElement("p");
      p.setAttribute("role", "paragraph");
      for (const inEl of inlineEls) {
        if (inEl) p.appendChild(inEl);
      }
      return p;
    }

    if (block.kind === "table") {
      return this.renderTable(block);
    }

    if (block.kind === "group") {
      const grp = document.createElement("div");
      grp.setAttribute("role", "group");
      for (const b of block.blocks) {
        const el = this.renderBlock(b);
        if (el) grp.appendChild(el);
      }
      return grp;
    }

    return null;
  }

  private renderTable(table: LayoutTable): HTMLElement {
    const tableEl = document.createElement("table");
    tableEl.setAttribute("role", "table");
    const tbody = document.createElement("tbody");

    for (let r = 0; r < table.rows.length; r++) {
      const row = table.rows[r];
      const tr = document.createElement("tr");
      tr.setAttribute("role", "row");

      for (let c = 0; c < row.cells.length; c++) {
        const cell = row.cells[c];
        const isHeader = r === 0; // standard fallback
        const cellEl = document.createElement(isHeader ? "th" : "td");
        cellEl.setAttribute("role", isHeader ? "columnheader" : "cell");

        for (const block of cell.blocks) {
          const bEl = this.renderBlock(block);
          if (bEl) cellEl.appendChild(bEl);
        }

        tr.appendChild(cellEl);
      }

      tbody.appendChild(tr);
    }

    tableEl.appendChild(tbody);
    return tableEl;
  }

  private renderInline(inline: LayoutInline): Node | null {
    if (inline.kind === "text") {
      if (inline.link?.url) {
        const a = document.createElement("a");
        a.href = inline.link.url;
        if (inline.link.tooltip) a.title = inline.link.tooltip;
        a.textContent = inline.text;
        return a;
      }
      return document.createTextNode(inline.text);
    }

    if (inline.kind === "math") {
      const mathSpan = document.createElement("span");
      mathSpan.setAttribute("role", "math");
      mathSpan.setAttribute("aria-label", inline.label || "Mathematical equation");
      mathSpan.textContent = inline.label || "Math formula";
      return mathSpan;
    }

    if (inline.kind === "picture") {
      const img = document.createElement("img");
      img.setAttribute("role", "img");
      img.alt = "Document image";
      if (inline.src) img.src = inline.src;
      return img;
    }

    if (inline.kind === "break") {
      return document.createElement("br");
    }

    return null;
  }
}
