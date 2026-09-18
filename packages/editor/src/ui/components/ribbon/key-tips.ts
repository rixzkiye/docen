import { FASTElement, css, customElement, html } from "@microsoft/fast-element";

export interface KeyTipBadge {
  id: string;
  key: string;
  target: HTMLElement;
  rect: { left: number; top: number; width: number; height: number };
  action: () => void;
}

const TAB_KEYS: Record<string, string> = {
  home: "H",
  insert: "I",
  layout: "P",
  page: "P",
  references: "S",
  review: "R",
  view: "V",
  draw: "D",
  design: "G",
  mailings: "M",
  developer: "L",
};

const COMMON_CONTROL_KEYS: Record<string, string> = {
  bold: "B",
  italic: "I",
  underline: "U",
  "grow-font": "FG",
  "shrink-font": "FK",
  "font-color": "FC",
  highlight: "HC",
  "copy-format": "FP",
  "format-painter": "FP",
  "font-name": "FF",
  "font-size": "FS",
  "align-left": "AL",
  "align-center": "AC",
  "align-right": "AR",
  justify: "AJ",
  "bullet-list": "BL",
  "ordered-list": "NL",
  "multilevel-list": "ML",
  style: "SS",
  "page-break": "B",
  table: "T",
  "insert-table": "T",
  picture: "P",
  comment: "C",
  "new-comment": "C",
  link: "K",
};

const styles = css`
  :host {
    position: fixed;
    inset: 0;
    z-index: 10002;
    pointer-events: none;
    display: none;
  }

  :host([data-active]) {
    display: block;
  }

  .badge {
    position: fixed;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 14px;
    height: 18px;
    padding: 0 4px;
    box-sizing: border-box;
    background: #111111;
    color: #ffffff;
    border: 1px solid #ffffff;
    border-radius: 3px;
    font-family: "Segoe UI", system-ui, sans-serif;
    font-size: 11px;
    font-weight: 700;
    line-height: 1;
    text-transform: uppercase;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.45);
    transform: translate(-50%, -50%);
    pointer-events: none;
    user-select: none;
    -webkit-user-select: none;
  }
`;

const template = html<DocenKeyTips>` <div class="keytips-container"></div> `;

/**
 * `<docen-key-tips>` — Office-style Ribbon Key Tips (Alt) overlay.
 * Renders high-contrast badge labels over ribbon tab headers (H, I, P, S, R, V)
 * and active ribbon controls, supporting full keyboard navigation.
 */
@customElement({ name: "docen-key-tips", template, styles })
export class DocenKeyTips extends FASTElement {
  #active = false;
  #mode: "tabs" | "controls" = "tabs";
  #badges: KeyTipBadge[] = [];
  #buffer = "";
  #focusedElement: HTMLElement | null = null;
  #container?: HTMLElement;

  get isActive(): boolean {
    return this.#active;
  }

  get mode(): "tabs" | "controls" {
    return this.#mode;
  }

  get badges(): readonly KeyTipBadge[] {
    return this.#badges;
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#container = this.shadowRoot?.querySelector(".keytips-container") as HTMLElement;
    window.addEventListener("keydown", this.#onKeyDown);
  }

  disconnectedCallback(): void {
    this.dismiss();
    window.removeEventListener("keydown", this.#onKeyDown);
    super.disconnectedCallback();
  }

  toggle(): void {
    if (this.#active) {
      this.dismiss();
    } else {
      this.activate();
    }
  }

  activate(): void {
    this.#active = true;
    this.#mode = "tabs";
    this.#buffer = "";
    this.setAttribute("data-active", "");
    this.#renderTabBadges();
  }

  dismiss(): void {
    this.#active = false;
    this.#mode = "tabs";
    this.#buffer = "";
    this.#badges = [];
    this.removeAttribute("data-active");
    if (this.#container) this.#container.replaceChildren();
  }

  #onKeyDown = (event: KeyboardEvent): void => {
    if (!this.isConnected) return;
    // Single Alt press toggles key tips
    if (event.key === "Alt" && !event.ctrlKey && !event.metaKey && !event.shiftKey) {
      // Don't prevent default on bare Alt so browser menu heuristics don't completely break,
      // but toggle key tips.
      event.preventDefault();
      this.toggle();
      return;
    }

    if (!this.#active) return;

    // When key tips are active, intercept navigation and letter keys:
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      if (this.#mode === "controls") {
        this.#mode = "tabs";
        this.#buffer = "";
        this.#renderTabBadges();
      } else {
        this.dismiss();
      }
      return;
    }

    // Tab / Shift+Tab keyboard navigation
    if (event.key === "Tab") {
      event.preventDefault();
      event.stopPropagation();
      this.#navigateFocus(event.shiftKey ? -1 : 1);
      return;
    }

    // Arrow navigation
    if (event.key === "ArrowLeft" || event.key === "ArrowUp") {
      event.preventDefault();
      event.stopPropagation();
      this.#navigateFocus(-1);
      return;
    }
    if (event.key === "ArrowRight" || event.key === "ArrowDown") {
      event.preventDefault();
      event.stopPropagation();
      this.#navigateFocus(1);
      return;
    }

    // Enter or Space activates focused control
    if (event.key === "Enter" || event.key === " ") {
      if (this.#focusedElement) {
        event.preventDefault();
        event.stopPropagation();
        this.#focusedElement.click();
        this.dismiss();
        return;
      }
    }

    // Handle alphanumeric badge selection
    if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      event.stopPropagation();
      this.#handleLetter(event.key.toUpperCase());
    }
  };

  #handleLetter(char: string): void {
    this.#buffer += char;

    // Check exact matches
    const exact = this.#badges.find((b) => b.key === this.#buffer);
    if (exact) {
      exact.action();
      return;
    }

    // Check prefix matches
    const prefixMatches = this.#badges.filter((b) => b.key.startsWith(this.#buffer));
    if (prefixMatches.length === 0) {
      // No match, reset buffer
      this.#buffer = char;
      const reExact = this.#badges.find((b) => b.key === this.#buffer);
      if (reExact) {
        reExact.action();
      }
    } else {
      // Partial match: filter rendered badges to those matching prefix
      this.#filterBadgesToPrefix(this.#buffer);
    }
  }

  #filterBadgesToPrefix(prefix: string): void {
    if (!this.#container) return;
    for (const badgeEl of this.#container.querySelectorAll<HTMLElement>(".badge")) {
      const k = badgeEl.dataset.key ?? "";
      badgeEl.style.display = k.startsWith(prefix) ? "inline-flex" : "none";
    }
  }

  #renderTabBadges(): void {
    if (!this.#container) {
      this.#container = this.shadowRoot?.querySelector(".keytips-container") as HTMLElement;
    }
    if (!this.#container) return;
    this.#container.replaceChildren();
    this.#badges = [];

    const root = this.getRootNode() as Document | ShadowRoot;
    const tablist = root.querySelector("fluent-tablist");
    const tabs = root.querySelectorAll<HTMLElement>("docen-ribbon-tab");

    tabs.forEach((tab) => {
      const tabId = (tab.id || tab.getAttribute("value") || "").toLowerCase();
      const badgeKey = TAB_KEYS[tabId] ?? tab.textContent?.trim().charAt(0).toUpperCase() ?? "H";
      const rect = tab.getBoundingClientRect();
      if (rect.width === 0 && rect.height === 0) return;

      const badge: KeyTipBadge = {
        id: tabId,
        key: badgeKey,
        target: tab,
        rect: {
          left: rect.left + rect.width / 2,
          top: rect.bottom - 4,
          width: rect.width,
          height: rect.height,
        },
        action: () => {
          // Switch to tab
          if (tablist) {
            (tablist as unknown as { activeid: string }).activeid = tab.id;
          }
          tab.click();
          this.#mode = "controls";
          this.#buffer = "";
          // Re-render badges for newly active tab's panel
          requestAnimationFrame(() => this.#renderControlBadges());
        },
      };

      this.#badges.push(badge);
      this.#createBadgeElement(badge);
    });
  }

  #renderControlBadges(): void {
    if (!this.#container) return;
    this.#container.replaceChildren();
    this.#badges = [];

    const root = this.getRootNode() as Document | ShadowRoot;
    const ribbon = root.querySelector("docen-ribbon");
    if (!ribbon) return;

    // Find active ribbon panel
    const activePanel = ribbon.querySelector("docen-ribbon-panel[data-active]");
    if (!activePanel) return;

    // Find interactive controls inside the active panel
    const controls = activePanel.querySelectorAll<HTMLElement>(
      "docen-ribbon-button, docen-ribbon-split-button, docen-ribbon-toggle-button, docen-ribbon-combobox, docen-ribbon-menu, docen-ribbon-gallery, docen-color-picker, button",
    );

    const usedKeys = new Set<string>();
    let autoIndex = 1;

    controls.forEach((el) => {
      const rect = el.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) return;

      const cmd =
        el.getAttribute("event") || el.getAttribute("data-cmd") || el.getAttribute("value") || "";

      let key = COMMON_CONTROL_KEYS[cmd];
      if (!key || usedKeys.has(key)) {
        // Generate clean unique key
        const label = el.getAttribute("label") || el.getAttribute("title") || "";
        const c1 = (label.charAt(0) || cmd.charAt(0) || "A").toUpperCase();
        if (!usedKeys.has(c1)) {
          key = c1;
        } else {
          key = `${c1}${autoIndex++}`;
        }
      }
      usedKeys.add(key);

      const badge: KeyTipBadge = {
        id: cmd || key,
        key,
        target: el,
        rect: {
          left: rect.left + rect.width / 2,
          top: rect.top + rect.height / 2,
          width: rect.width,
          height: rect.height,
        },
        action: () => {
          this.dismiss();
          el.focus();
          el.click();
        },
      };

      this.#badges.push(badge);
      this.#createBadgeElement(badge);
    });
  }

  #createBadgeElement(badge: KeyTipBadge): void {
    if (!this.#container) return;
    const el = document.createElement("span");
    el.className = "badge";
    el.textContent = badge.key;
    el.dataset.key = badge.key;
    el.style.left = `${Math.round(badge.rect.left)}px`;
    el.style.top = `${Math.round(badge.rect.top)}px`;
    this.#container.append(el);
  }

  #navigateFocus(step: number): void {
    if (this.#badges.length === 0) return;
    let idx = this.#badges.findIndex((b) => b.target === this.#focusedElement);
    if (idx === -1) {
      idx = step > 0 ? 0 : this.#badges.length - 1;
    } else {
      idx = (idx + step + this.#badges.length) % this.#badges.length;
    }
    const next = this.#badges[idx]?.target;
    if (next) {
      this.#focusedElement = next;
      next.focus();
    }
  }
}

if (!customElements.get("docen-key-tips")) {
  customElements.define("docen-key-tips", DocenKeyTips);
}

export default DocenKeyTips;
