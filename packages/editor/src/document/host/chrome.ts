/**
 * Ribbon/title-bar chrome domain — split out of the host element (see
 * document/index.ts): the Quick Access Toolbar, auto-save, the header and pane
 * renders, and every ribbon greying/menu/context-tab sync pass. The host keeps
 * thin delegating members so its call sites stay unchanged.
 */

import { resolveTableLook, type JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import { redoDepth, undoDepth } from "@tiptap/pm/history";
import { NodeSelection } from "@tiptap/pm/state";

import { drawingSelectionKind } from "../../drawing";
import { t, type DocenAddin, type RibbonMenuItem } from "../../ui";
import { applyTheme, mergeRibbonSchema, resolveTheme, type RibbonTab } from "../../ui";
import { blocksOfDocAttrs, groupBlocksByGallery } from "../building-blocks";
import type { EditBridge } from "../canvas/edit-bridge";
import type { StoryKind } from "../canvas/edit-bridge";
import { escapeHtml } from "../chrome";
import type { MailMergeCommands } from "../commands/mail-merge";
import {
  chartMenuValueOf,
  floatingDrawingAt,
  formatToggleStatesOf,
  inlineDrawingAt,
  inlineImageAt,
  positionMenuValueOf,
  tableAncestry,
  textDirectionMenuValueOf,
  WIRED_DISPATCH,
  wrapMenuValueOf,
} from "../extensions/commands";
import { collectRevisions } from "../extensions/track-changes";
import { LOCAL_HANDLED, type SaveFormat } from "../file-formats";
import type { TaskPaneId } from "../index";
import { mathAtomAt } from "../math-atom";
import { equationContextTab } from "../ribbon";
import {
  DEFAULT_RIBBON_TAB,
  buildContextualTab,
  chartDesignTab,
  formatMeasureTwip,
  headerFooterContextTab,
  pictureFormatTab,
  renderRibbonFromSchema,
  ribbonActions,
  ribbonTabs,
  shapeFormatTab,
  tableContextTabs,
  useCmUnits,
} from "../ribbon";
import type { DocenSettings } from "../settings";
import type { StoriesDomain } from "./stories";

/** Quick Access Toolbar candidates (Word's customize-QAT menu): each id is a
 *  routed command (`event`), shown in the title bar while checked. The shown
 *  set persists in localStorage (`docen:qat`); the order here is the bar's. */
/** Split-button faces whose command greys from the text selection, not the
 *  dropdown value. */
const FACE_ONLY_SPLITS: ReadonlySet<string> = new Set(["autofit", "columns"]);

/** Arrange events that need a floating object (Word greys the rest). */
const FLOATING_ONLY = new Set([
  "align-objects",
  "bring-forward",
  "send-backward",
  "bring-to-front",
  "send-to-back",
]);

/** Arrange events that also serve an inline drawing — wrap and position
 *  convert it to a floating one (Word's galleries convert on click). */
const FLOATING_OR_INLINE = new Set(["wrap", "position"]);

/** Rotate also serves an inline picture; inline shapes and charts don't
 *  rotate. */
const FLOATING_OR_INLINE_IMAGE = new Set(["rotate"]);

const QAT_CANDIDATES: readonly { id: string; icon: string; labelKey: string }[] = [
  { id: "new", icon: "new", labelKey: "header.new" },
  { id: "open", icon: "open", labelKey: "header.open" },
  { id: "save", icon: "save", labelKey: "header.save" },
  { id: "print", icon: "print", labelKey: "header.print" },
  { id: "undo", icon: "undo", labelKey: "header.undo" },
  { id: "redo", icon: "redo", labelKey: "header.redo" },
  { id: "repeat", icon: "repeat", labelKey: "header.repeat" },
  { id: "spell-check", icon: "spell-check", labelKey: "ribbon.cmd.spell-check" },
];
const QAT_DEFAULT: readonly string[] = ["save", "undo", "redo"];
const QAT_STORAGE_KEY = "docen:qat";
const AUTOSAVE_ON_KEY = "docen:autosave";
/** localStorage quota is ~5MB per origin — skip the write (keep the last
 *  backup) rather than throwing mid-typing. */
const AUTOSAVE_MAX_CHARS = 4_000_000;

/** A version-history snapshot (auto-save writes them; the dialog restores). */
export type VersionSnapshot = {
  id: string;
  timestamp: string;
  author: string;
  isAutosave: boolean;
  doc: JSONContent;
};

/** The chrome domain's view of the host — only what its bodies touch. */
export interface ChromeHostView {
  element(): HTMLElement;
  root(): ShadowRoot | null;
  editor(): Editor | undefined;
  bridge(): EditBridge | undefined;
  settings(): DocenSettings;
  addins(): readonly DocenAddin[];
  addAddin(addin: DocenAddin): void;
  removeAddin(id: string): void;
  getAttribute(name: string): string | null;
  hasAttribute(name: string): boolean;
  markdown(): boolean;
  markupView(): "simple" | "all" | "none" | "original";
  markupAuthors(): string[] | null;
  markupColors(): "author" | "changeType";
  balloons(): "all" | "comments" | "revisions" | "none";
  snapshots(): VersionSnapshot[];
  setSnapshots(value: VersionSnapshot[]): void;
  merge(): MailMergeCommands;
  stories(): StoriesDomain;
  storyKind(): StoryKind | null;
  getTaskpaneState(id: TaskPaneId): boolean;
  setTaskpane(id: TaskPaneId, open: boolean): void;
  updateStatus(): void;
  dispatch(event: Event): boolean;
}

/**
 * The ribbon/title-bar chrome pass, split out of the host element. Owns the
 * QAT/auto-save persistence caches and the markup-display diff caches; the
 * display state itself (markup view, balloons) stays on the host.
 */
export class ChromeDomain {
  /** The previous arrange-greying selection signature (null = re-read). */
  #arrangeGrey: string | null = null;
  /** The previous format toggles' [event, lit] signature — the DOM sweep is
   *  skipped while it is unchanged; "" forces a re-run after a re-stamp. */
  #formatButtonsKey = "";
  /** Ids the addins are currently registered for (the attribute diff base). */
  #addinAttrIds = new Set<string>();
  /** Contextual-tab ids the current ribbon DOM carries. */
  #contextTabIds = new Set<string>();
  /** Auto-save debounce timer (per-filename switch + backup). */
  #autosaveTimer?: number;

  constructor(private readonly host: ChromeHostView) {}

  /** Reconnect-safe teardown: drop the auto-save debounce and the context-tab
   *  diff so a reconnect re-stamps the ribbon. */
  dispose(): void {
    clearTimeout(this.#autosaveTimer);
    this.#contextTabIds.clear();
  }

  /** The shown QAT ids, persisted across sessions; falls back to Word's
   *  default trio when nothing (or anything stale) is stored. */
  qatIds(): string[] {
    try {
      const raw = localStorage.getItem(QAT_STORAGE_KEY);
      if (raw) {
        const ids = JSON.parse(raw) as unknown;
        if (
          Array.isArray(ids) &&
          ids.length > 0 &&
          ids.every((id) => QAT_CANDIDATES.some((c) => c.id === id))
        ) {
          return ids as string[];
        }
      }
    } catch {
      // Private mode / disabled storage — defaults only.
    }
    return [...QAT_DEFAULT];
  }

  toggleQat(id: string): void {
    const cur = this.qatIds();
    const next = cur.includes(id) ? cur.filter((x) => x !== id) : [...cur, id];
    try {
      localStorage.setItem(QAT_STORAGE_KEY, JSON.stringify(next));
    } catch {
      // Storage unavailable — the bar still toggles for this session.
    }
    // Re-stamp the header so the bar and the menu's checkmarks follow.
    this.renderChrome();
  }

  autosaveEnabled(): boolean {
    try {
      return localStorage.getItem(AUTOSAVE_ON_KEY) === "1";
    } catch {
      return false;
    }
  }

  setAutosave(on: boolean): void {
    try {
      localStorage.setItem(AUTOSAVE_ON_KEY, on ? "1" : "0");
      if (on) this.scheduleAutosave();
      else {
        clearTimeout(this.#autosaveTimer);
        localStorage.removeItem(this.autosaveKey());
      }
    } catch {
      // Storage unavailable — the switch still flips for this session.
    }
  }

  autosaveKey(): string {
    return `docen:autosave:${this.host.getAttribute("filename") ?? "document"}`;
  }

  loadPersistedHistory(): void {
    try {
      const raw = localStorage.getItem(`${this.autosaveKey()}:history`);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          this.host.setSnapshots(parsed.slice(0, 20));
        }
      }
    } catch {
      // Storage unavailable
    }
  }

  scheduleAutosave(): void {
    if (!this.autosaveEnabled() || !this.host.editor()) return;
    clearTimeout(this.#autosaveTimer);
    this.#autosaveTimer = window.setTimeout(() => {
      try {
        const curDoc = this.host.editor()?.getJSON();
        if (curDoc) {
          const json = JSON.stringify(curDoc);
          if (json.length <= AUTOSAVE_MAX_CHARS) localStorage.setItem(this.autosaveKey(), json);
          this.host.snapshots().unshift({
            id: `v-${Date.now()}`,
            timestamp: new Date().toISOString(),
            author: this.host.settings().identity.name || "User",
            isAutosave: true,
            doc: curDoc,
          });
          if (this.host.snapshots().length > 20) this.host.snapshots().pop();
          try {
            localStorage.setItem(
              `${this.autosaveKey()}:history`,
              JSON.stringify(this.host.snapshots()),
            );
          } catch {}
        }
      } catch {
        // Quota exceeded — keep the last good backup, retry on the next change.
      }
    }, 1500);
  }

  renderHeader(): string {
    // The store is the identity source; an explicit `user` attribute overrides
    // it for display (Options edits the store, the attribute remains a host knob).
    const identity = this.host.settings().identity;
    const user = identity.name;
    const avatar = this.host.getAttribute("avatar") ?? "";
    const filename =
      this.host.getAttribute("filename") ?? t("header.doc-name", this.host.element());
    const initial = identity.initials || user.trim().charAt(0).toUpperCase();
    const avatarMarkup = avatar
      ? `<img class="avatar avatar-img" src="${escapeHtml(avatar)}" alt="" />`
      : initial
        ? `<span class="avatar">${escapeHtml(initial)}</span>`
        : "";
    const autosave = t("header.autosave", this.host.element());
    const qatIds = this.qatIds();
    // Undo/redo render as icon-only split buttons (Word's QAT shape): the
    // primary runs one step, the 14px caret opens the history flyout
    // (#fillHistory fills it on open). An empty stack hides the caret through
    // data-history-empty (documentStyles; Word shows no flyout for a fresh
    // document) — #updateStatus keeps the flag in step with the depths, the
    // header doesn't rebuild per transaction.
    const editorForDepth = this.host.bridge()?.activeEditor() ?? this.host.editor();
    const depthOf = (kind: "undo" | "redo"): number =>
      !editorForDepth
        ? 0
        : kind === "undo"
          ? undoDepth(editorForDepth.state)
          : redoDepth(editorForDepth.state);
    const qatButtons = QAT_CANDIDATES.filter((c) => qatIds.includes(c.id))
      .map((c) => {
        if (c.id === "undo" || c.id === "redo") {
          return `<docen-ribbon-split-button icon="${c.icon}" label="${t(c.labelKey, this.host.element())}" event="${c.id}" icon-only data-history="${c.id}" items="[]"${depthOf(c.id) === 0 ? " data-history-empty" : ""}></docen-ribbon-split-button>`;
        }
        return `<docen-ribbon-button icon="${c.icon}" label="${t(c.labelKey, this.host.element())}" event="${c.id}" icon-only></docen-ribbon-button>`;
      })
      .join("");
    const qatMenuItems = QAT_CANDIDATES.map((c) => {
      const on = qatIds.includes(c.id);
      // `checked` drives Fluent's checkmark glyph and the change event;
      // aria-checked is kept in sync by the element internals.
      return `<fluent-menu-item role="menuitemcheckbox" ${on ? "checked" : ""} data-qat="${c.id}">${t(c.labelKey, this.host.element())}</fluent-menu-item>`;
    }).join("");
    return `
          <div slot="start" style="display:flex;align-items:center;gap:4px">
            <span style="font-weight:600;font-size:13px;padding-inline:6px">${t("header.brand", this.host.element())}</span>
            <span class="autosave-label">${autosave}</span>
            <fluent-switch data-event="autosave" ${this.autosaveEnabled() ? "checked" : ""} aria-label="${autosave}"></fluent-switch>
            ${qatButtons}
            <!-- The customize caret and the file menu detach from the QAT
                 cluster (the row's 4px gap would read the caret as the last
                 button's split dropdown, and butt it against the file name). -->
            <fluent-menu style="margin-inline-start:4px">
              <fluent-menu-button
                slot="trigger"
                appearance="subtle"
                icon-only
                class="qat-customize"
                title="${t("header.qat-customize", this.host.element())}"
              ></fluent-menu-button>
              <fluent-menu-list>${qatMenuItems}</fluent-menu-list>
            </fluent-menu>
            <fluent-menu style="margin-inline-start:10px">
              <fluent-menu-button
                slot="trigger"
                appearance="subtle"
                style="max-width:36vw;overflow:hidden;white-space:nowrap"
                title="${escapeHtml(filename)}"
              >${escapeHtml(filename)}</fluent-menu-button>
              <fluent-menu-list>
                <fluent-menu-item data-event="new">${t("header.new", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="new-from-template">${t("header.new-from-template", this.host.element())}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="open">${t("header.open", this.host.element())}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="save-as">${t("header.save-as", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="version-history">${t("header.version-history", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-template">${t("header.save-as-template", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-markdown">${t("header.save-as-markdown", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-rtf">${t("header.save-as-rtf", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-html">${t("header.save-as-html", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-txt">${t("header.save-as-txt", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-odt">${t("header.save-as-odt", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="save-as-pdf">${t("header.save-as-pdf", this.host.element())}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="print">${t("header.print", this.host.element())}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="share">${t("header.share", this.host.element())}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="properties">${t("header.properties", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="inspect-document">${t("header.inspect", this.host.element())}</fluent-menu-item>
                <fluent-divider role="separator" aria-orientation="horizontal" orientation="horizontal"></fluent-divider>
                <fluent-menu-item data-event="options">${t("header.options", this.host.element())}</fluent-menu-item>
                <fluent-menu-item data-event="close">${t("header.close", this.host.element())}</fluent-menu-item>
              </fluent-menu-list>
            </fluent-menu>
          </div>
          <docen-command-search slot="search"></docen-command-search>
          <div slot="end" style="display:flex;align-items:center;gap:4px">
            <span style="display:inline-flex;align-items:center;gap:6px;padding-inline:6px">${avatarMarkup}${escapeHtml(user)}</span>
          </div>`;
  }

  /** Stamp the header + ribbon markup for the active locale (re-run on lang change). */
  renderChrome(): void {
    const root = this.host.root();
    // FAST fires @attr change callbacks during element upgrade, BEFORE the
    // template is stamped (connectedCallback runs after) — the shadowRoot
    // exists but is empty, so the title-bar query is null. Bail until stamped;
    // connectedCallback's explicit call does the first render.
    const titleBar = root?.querySelector("docen-title-bar");
    if (!root || !titleBar) return;
    const styles = this.host.editor()?.state.doc.attrs?.styles ?? null;
    titleBar.innerHTML = this.renderHeader();
    // Built-in tabs (Home/Insert/… with the live style gallery) come from
    // ribbonTabs; external add-ins layer their own tabs on top via
    // mergeRibbonSchema. The default add-in contributes no ribbon, so without
    // extra add-ins this is just the built-in set.
    const tabs = [
      ...ribbonTabs(styles, { revisionAuthors: this.revisionAuthors() }),
      ...mergeRibbonSchema(this.host.addins()),
    ];
    const ribbonEl = root.querySelector("docen-ribbon")!;
    // Pass the workspace as the i18n scope so labels resolve against
    // `<docen-workspace lang>` (forwarded from `<docen-document lang>`)
    // rather than `<html lang>`. `closest()` can't reach the workspace from
    // inside this fragment (shadow boundary + not yet inserted), so the
    // workspace element must be handed in explicitly.
    ribbonEl.replaceChildren(
      renderRibbonFromSchema(
        tabs,
        ribbonActions(),
        root.querySelector("docen-workspace") ?? document.documentElement,
      ),
    );
    // Feed the full ribbon schema (built-in tabs + addin contributions) to the
    // command search so it can flatten and index every command. Re-runs on
    // lang/addin change since #renderChrome is the single chrome re-stamp.
    const searchEl = root.querySelector("docen-command-search") as
      | (HTMLElement & { setTabs(tabs: readonly unknown[], scope: Element | null): void })
      | null;
    // Pass the workspace as the i18n scope so command labels resolve against
    // `<docen-workspace lang>` (forwarded from `<docen-document lang>`) — the
    // same scope the ribbon uses just above.
    searchEl?.setTabs(tabs, root.querySelector("docen-workspace"));
    this.applyRibbonGreying();
    this.syncEditModeMenu();
    this.syncStoryMenus();
    this.syncMarkupMenus();
    // The ribbon DOM was rebuilt from scratch — drop the stale context-tab
    // tracking, then re-append them if the selection is inside a table.
    this.#contextTabIds.clear();
    this.syncContextTabs();
    this.syncCellSize();
    this.syncDrawingSize();
    this.syncFormatButtons();
    this.syncDrawingMenus();
    this.syncQuickPartsMenu();
    this.renderPanes();
  }

  /** Fill a QAT history flyout with one entry per available step. PM's history
   *  keeps no per-item labels, so entries read "Edit N"; picking entry N
   *  arrives as the undo/redo command carrying its step count (#onCommand,
   *  Word's flyout shape). */
  fillHistory(kind: "undo" | "redo"): void {
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    const split = this.host
      .root()
      ?.querySelector(`docen-ribbon-split-button[data-history="${kind}"]`);
    if (!editor || !split) return;
    const depth = kind === "undo" ? undoDepth(editor.state) : redoDepth(editor.state);
    const label = t("header.history-item", this.host.element());
    const items: Array<{ text: string; value: string }> = [];
    for (let steps = depth; steps >= 1; steps--) {
      items.push({ text: label.replace("{0}", String(steps)), value: String(steps) });
    }
    split.setAttribute("items", JSON.stringify(items));
  }

  /** Stamp pane titles + status text for the active locale (re-run on lang change). */
  renderPanes(): void {
    const root = this.host.root();
    if (!root) return;
    const navPane = root.querySelector('docen-task-pane[position="start"]');
    if (navPane) navPane.setAttribute("title", t("pane.navigation", this.host.element()));
    const propsPane = root.querySelector('docen-task-pane[position="end"]');
    if (propsPane) propsPane.setAttribute("title", t("pane.properties", this.host.element()));
    // The end-rail panes — the static template's title attrs are English
    // literals, so every pane's title is stamped here for the locale.
    for (const [part, key] of [
      ["comments-pane", "pane.comments"],
      ["revisions-pane", "pane.revisions"],
      ["clipboard-pane", "pane.clipboard"],
      ["proofing-pane", "pane.proofing"],
      ["thesaurus-pane", "pane.thesaurus"],
      ["styles-pane", "pane.styles"],
    ] as const) {
      root
        .querySelector(`docen-task-pane[part="${part}"]`)
        ?.setAttribute("title", t(key, this.host.element()));
    }
    // Status bar is dynamic (page count / caret page / zoom) — re-stamp it so a
    // locale change re-localizes the text too.
    this.host.updateStatus();
  }

  /** Sync external add-ins with the `addins` JSON attribute: register new ids,
   *  remove ids no longer present. JSON can't carry functions, so only ribbon
   *  data contributions cross this boundary; command handlers stay in JS
   *  (addAddin with a full object). */
  applyAddinsAttr(): void {
    const raw = this.host.getAttribute("addins");
    const next = new Set<string>();
    if (raw) {
      let parsed: unknown;
      try {
        parsed = JSON.parse(raw);
      } catch {
        return;
      }
      if (Array.isArray(parsed)) {
        for (const item of parsed) {
          if (
            item &&
            typeof item === "object" &&
            typeof (item as { id?: unknown }).id === "string"
          ) {
            const id = (item as { id: string }).id;
            next.add(id);
            if (!this.#addinAttrIds.has(id)) this.host.addAddin(item as DocenAddin);
          }
        }
      }
    }
    // Remove add-ins that fell out of the attribute (covers editing it to drop
    // a tab at runtime, or removing the attribute entirely).
    for (const id of this.#addinAttrIds) {
      if (!next.has(id)) this.host.removeAddin(id);
    }
    this.#addinAttrIds = next;
  }

  /** Apply the `theme` attribute: switch the Fluent theme
   *  (light/dark/high-contrast/teams-*). */
  applyThemeAttr(value: string): void {
    applyTheme(resolveTheme(value));
  }

  /** Grey out ribbon commands that have no handler (skeleton buttons). Runs
   *  after every ribbon re-stamp; fresh elements start un-disabled, so this is
   *  the single place `disabled` is applied. Only controls that support
   *  `disabled` (button/split-button/toggle-button) are greyed — combobox /
   *  color-picker lack it and live in wired tabs anyway. */
  applyRibbonGreying(): void {
    const ribbon = this.host.root()?.querySelector("docen-ribbon");
    if (!ribbon) return;
    const wired = this.wiredCommands();
    ribbon
      .querySelectorAll<HTMLElement>(
        "docen-ribbon-button[event], docen-ribbon-split-button[event], docen-ribbon-toggle-button[event], docen-ribbon-menu[event]",
      )
      .forEach((el) => {
        const event = el.getAttribute("event");
        if (!event) return;
        // A composite (split/menu) stays live while ANY drop-down variant
        // resolves to a wired command — greying the host would bury its live
        // items (the AutoFit split's face has no action, its three variants
        // do). A face-only split keeps its caret and opens the menu instead.
        const liveItems =
          el.tagName === "DOCEN-RIBBON-SPLIT-BUTTON" || el.tagName === "DOCEN-RIBBON-MENU"
            ? this.ribbonMenuItems(el).some(
                (item) => !item.disabled && wired.has(item.event ?? event),
              )
            : false;
        if (wired.has(event) || liveItems) {
          el.removeAttribute("disabled");
          // A face with no action of its own opens the drop-down instead:
          // either its own event is unwired while the variants are live, or
          // the handler only exists for the variants' values (Columns).
          const faceOnly =
            el.tagName === "DOCEN-RIBBON-SPLIT-BUTTON" && FACE_ONLY_SPLITS.has(event);
          if (faceOnly || (liveItems && !wired.has(event)))
            el.setAttribute("primary-opens-menu", "");
          else el.removeAttribute("primary-opens-menu");
        } else {
          el.setAttribute("disabled", "");
          el.removeAttribute("primary-opens-menu");
        }
      });
    // The merge controls need a recipient data source (Word grays them the
    // same way until Select Recipients has run).
    const hasSource = this.host.merge().recipients() !== null;
    for (const event of [
      "merge-field",
      "address-block",
      "greeting-line",
      "preview-results",
      "first-record",
      "last-record",
      "finish-merge",
    ]) {
      ribbon
        .querySelectorAll<HTMLElement>(`[event="${event}"]`)
        .forEach((el) => el.toggleAttribute("disabled", !hasSource));
    }
    // The ribbon DOM here is fresh (rebuilt or extended) — force the arrange
    // pass to re-read the selection instead of trusting the diff cache.
    this.#arrangeGrey = null;
    this.syncArrangeGreying();
    // Same for the format toggles: fresh elements start un-pressed, so the
    // signature cache must not skip the re-stamp below.
    this.#formatButtonsKey = "";
  }

  /** Word greys the Arrange group by selection: Align/z-order need a floating
   *  drawing, Wrap Text and Position also serve an inline drawing (they
   *  convert it), Rotate also an inline picture. The static pass can't see
   *  the selection, so this runs per transaction (and after every ribbon
   *  re-stamp). */
  syncArrangeGreying(): void {
    const ribbon = this.host.root()?.querySelector("docen-ribbon");
    if (!ribbon || !this.host.editor()) return;
    const state = this.host.editor()!.state;
    const floating = floatingDrawingAt(state) != null;
    const inlineImage = !floating && inlineImageAt(state) != null;
    const inlineDrawing = !floating && !inlineImage && inlineDrawingAt(state) != null;
    const key = `${floating ? "f" : ""}${inlineImage ? "i" : ""}${inlineDrawing ? "d" : ""}`;
    if (key === this.#arrangeGrey) return;
    this.#arrangeGrey = key;
    for (const el of ribbon.querySelectorAll<HTMLElement>(
      "docen-ribbon-button[event], docen-ribbon-menu[event], docen-ribbon-split-button[event]",
    )) {
      const event = el.getAttribute("event") ?? "";
      const live = FLOATING_ONLY.has(event)
        ? floating
        : FLOATING_OR_INLINE.has(event)
          ? floating || inlineImage || inlineDrawing
          : FLOATING_OR_INLINE_IMAGE.has(event)
            ? floating || inlineImage
            : null;
      // Events outside the two sets keep the static pass's decision.
      if (live == null) continue;
      el.toggleAttribute("disabled", !live);
    }
  }

  /** Re-stamp the Home tab's format toggles (Bold/Italic/…/alignment) against
   *  the caret/selection — Word's lit buttons. Runs per transaction after
   *  #syncArrangeGreying; toggleAttribute is a no-op on a same-value attr, so
   *  an unchanged state doesn't re-fire the component. The [event, lit] rows
   *  are the pass's only inputs, so an unchanged signature skips the DOM
   *  sweep (a ribbon re-stamp resets the cache — see #applyRibbonGreying). */
  syncFormatButtons(): void {
    const state = this.host.editor()?.state;
    if (!state) return;
    const rows: [string, boolean][] = [
      ...formatToggleStatesOf(state),
      ["show-marks", this.host.hasAttribute("show-marks")],
      // The format painter lights while armed (Word: the button stays lit
      // until the paint lands / sticky mode ends) — the host attribute is
      // its truth, same as show-marks.
      ["format-painter", this.host.hasAttribute("format-painter")],
      // Markdown input mode — the host flag is its truth (the Options
      // dialog writes it without a click, so the sync re-stamps both ways).
      ["markdown-input", this.host.markdown()],
    ];
    const key = rows.map(([event, on]) => (on ? `${event}|` : `${event},`)).join("");
    if (key === this.#formatButtonsKey) return;
    this.#formatButtonsKey = key;
    for (const [event, on] of rows) {
      for (const el of this.host.root()?.querySelectorAll<HTMLElement>(
        // A lit row may land on a plain toggle (bold) or on a split whose
        // primary carries the face ([pressed-face] guards the rest — underline
        // / bullet / ordered lists are the only opt-ins today).
        `docen-ribbon-toggle-button[event="${event}"], docen-ribbon-split-button[pressed-face][event="${event}"]`,
      ) ?? []) {
        el.toggleAttribute("pressed", on);
      }
    }
  }

  /** A composite control's parsed `items` attribute (menu variants), empty on
   *  malformed JSON so a typo greys the control rather than crashing. */
  ribbonMenuItems(el: HTMLElement): {
    checked?: boolean;
    disabled?: boolean;
    event?: string;
    value?: string;
  }[] {
    try {
      return JSON.parse(el.getAttribute("items") ?? "[]") as {
        checked?: boolean;
        disabled?: boolean;
        event?: string;
        value?: string;
      }[];
    } catch {
      return [];
    }
  }

  /** Re-stamp the drawing state menus' checked rows against the selection —
   *  Wrap Text / Position (picture and shape tabs), Chart Type / Legend, and
   *  the shape Text Direction menu report the drawing's current mode (Word's
   *  checked gallery row). Runs per transaction after #syncContextTabs (the
   *  menus only exist while a drawing tab is stamped) — a same-value
   *  setAttribute fires no attr-changed callback, so an unchanged state
   *  doesn't re-render the menu. */
  syncDrawingMenus(): void {
    const state = this.host.editor()?.state;
    if (!state) return;
    const chart = chartMenuValueOf(state);
    const rows: [event: string, value: string | null][] = [
      ["wrap", wrapMenuValueOf(state)],
      ["position", positionMenuValueOf(state)],
      ["chart-type", chart?.type ?? null],
      ["chart-legend", chart?.legend ?? null],
      ["shape-text-direction", textDirectionMenuValueOf(state)],
    ];
    for (const [event, value] of rows) {
      for (const el of this.host
        .root()
        ?.querySelectorAll<HTMLElement>(`docen-ribbon-menu[event="${event}"]`) ?? []) {
        const items = this.ribbonMenuItems(el).map((item) =>
          item.value == null ? item : { ...item, checked: item.value === value },
        );
        const json = JSON.stringify(items);
        if (json !== el.getAttribute("items")) el.setAttribute("items", json);
      }
    }
  }

  /** Re-stamp the Insert → Text → Quick Parts menu with the document's
   *  building blocks grouped by gallery (Word's Explore Quick Parts), greying
   *  Save Selection without a selection. Runs per transaction and after
   *  #renderChrome rebuilds the ribbon (the static seed carries no blocks). */
  syncQuickPartsMenu(): void {
    const menu = this.host
      .root()
      ?.querySelector<HTMLElement>('docen-ribbon-menu[event="quick-parts"]');
    if (!menu) return;
    const editor = this.host.bridge()?.activeEditor() ?? this.host.editor();
    const blocks = editor ? blocksOfDocAttrs(editor.state.doc.attrs) : [];
    const selection = editor?.state.selection;
    const editable = editor?.isEditable ?? false;
    const items: RibbonMenuItem[] = [
      {
        text: t("ribbon.opt.save-quick-part", this.host.element()),
        value: "save",
        event: "save-quick-part",
        disabled: !editable || !selection || selection.empty,
      },
      {
        text: t("ribbon.opt.building-blocks-organizer", this.host.element()),
        value: "organizer",
        event: "building-blocks-organizer",
        disabled: blocks.length === 0,
      },
    ];
    for (const group of groupBlocksByGallery(blocks)) {
      items.push({
        text: t(`buildingBlock.gallery.${group.gallery}`, this.host.element()),
        header: true,
      });
      for (const block of group.blocks) {
        items.push({ text: block.name, value: block.id, event: "quick-parts" });
      }
    }
    const json = JSON.stringify(items);
    if (json !== menu.getAttribute("items")) menu.setAttribute("items", json);
  }

  /** Re-stamp the tab-row "Editing" menu so its label + checked item match the
   *  editor's live editable state (initial render, after a switch, and on
   *  locale change — #renderChrome re-stamps the ribbon, so this runs after
   *  #applyRibbonGreying to override the static default items). */
  syncEditModeMenu(): void {
    const menu = this.host.root()?.querySelector('docen-ribbon-menu[event="edit-mode"]');
    if (!menu) return;
    const editable = this.host.editor()?.isEditable ?? true;
    menu.setAttribute(
      "label",
      t(editable ? "ribbon.opt.editing" : "ribbon.opt.viewing", this.host.element()),
    );
    menu.setAttribute(
      "items",
      JSON.stringify([
        {
          text: t("ribbon.opt.editing", this.host.element()),
          event: "edit-mode",
          value: "edit",
          checked: editable,
        },
        {
          text: t("ribbon.opt.viewing", this.host.element()),
          event: "edit-mode",
          value: "view",
          checked: !editable,
        },
      ]),
    );
  }

  /** The document's revision authors (w:ins/@w:author, document order,
   *  deduped) — the Specific People menu's entries. */
  revisionAuthors(): string[] {
    if (!this.host.editor()) return [];
    const seen = new Set<string>();
    for (const r of collectRevisions(this.host.editor()!.state.doc)) {
      if (r.author !== "") seen.add(r.author);
    }
    return [...seen];
  }

  /** Re-stamp the Review → Tracking display controls so label + checked match
   *  the live Display for Review state (the #syncEditModeMenu pattern — runs
   *  on every chrome re-stamp, right after the greying pass). */
  syncMarkupMenus(): void {
    const root = this.host.root();
    const display = root?.querySelector('docen-ribbon-split-button[event="display-for-review"]');
    if (!display) return;
    const view = this.host.markupView();
    const viewKey = {
      simple: "simple-marks",
      all: "all-marks",
      none: "no-marks",
      original: "original-marks",
    }[view];
    display.setAttribute("label", t(`ribbon.opt.${viewKey}`, this.host.element()));
    display.setAttribute(
      "items",
      JSON.stringify(
        (
          [
            ["simple", "simple-marks"],
            ["all", "all-marks"],
            ["none", "no-marks"],
            ["original", "original-marks"],
          ] as const
        ).map(([value, key]) => ({
          text: t(`ribbon.opt.${key}`, this.host.element()),
          event: "display-for-review",
          value,
          checked: value === view,
        })),
      ),
    );
    const authors = this.revisionAuthors();
    const filtered = this.host.markupAuthors();
    display
      .closest("docen-ribbon-group")
      ?.querySelector<HTMLElement>('docen-ribbon-menu[event="review-specific-people"]')
      ?.setAttribute(
        "items",
        JSON.stringify([
          {
            text: t("ribbon.opt.all-reviewers", this.host.element()),
            event: "review-specific-people",
            value: "all",
            checked: filtered == null,
          },
          ...authors.map((a) => ({
            text: a,
            event: "review-specific-people",
            value: a,
            checked: filtered?.includes(a) ?? false,
          })),
        ]),
      );
    // Markup Colors: the two palette entries with their live check.
    const colors = this.host.markupColors();
    display
      .closest("docen-ribbon-group")
      ?.querySelector<HTMLElement>('docen-ribbon-menu[event="markup-colors"]')
      ?.setAttribute(
        "items",
        JSON.stringify(
          (
            [
              ["author", "by-author"],
              ["changeType", "by-change-type"],
            ] as const
          ).map(([value, key]) => ({
            text: t(`ribbon.opt.${key}`, this.host.element()),
            event: "markup-colors",
            value,
            checked: value === colors,
          })),
        ),
      );
    // Show Markup → Balloons: the four scope entries with their live check.
    const balloons = this.host.balloons();
    display
      .closest("docen-ribbon-group")
      ?.querySelector<HTMLElement>('docen-ribbon-menu[event="show-markup"]')
      ?.setAttribute(
        "items",
        JSON.stringify(
          (
            [
              ["all", "balloons-all"],
              ["comments", "balloons-comments"],
              ["revisions", "balloons-revisions"],
              ["none", "balloons-none"],
            ] as const
          ).map(([value, key]) => ({
            text: t(`ribbon.opt.${key}`, this.host.element()),
            event: "show-markup",
            value,
            checked: value === balloons,
          })),
        ),
      );
  }

  /** Re-stamp the Header/Footer split drop-downs with live checked flags —
   *  the slot-visibility items read sectionProperties (titlePage /
   *  evenAndOddHeaders), which the static ribbon schema can't carry. Runs on
   *  every chrome re-stamp and every transaction (a flag toggle flips its
   *  check on the next pass). */
  syncStoryMenus(): void {
    const attrs = this.host.editor()?.state.doc.attrs as
      | {
          sectionProperties?: { titlePage?: boolean };
          documentExtras?: { settings?: { evenAndOddHeaders?: boolean } };
        }
      | undefined;
    // titlePage is a sectPr flag; evenAndOddHeaders lives in settings.xml
    // (toggled through documentExtras — see SectionsCommands.toggleSectionFlag).
    const sp = attrs?.sectionProperties;
    const oddEven = attrs?.documentExtras?.settings?.evenAndOddHeaders;
    const stamp = (kind: "header" | "footer"): void => {
      const el = this.host.root()?.querySelector(`docen-ribbon-split-button[event="${kind}"]`);
      if (!el) return;
      el.setAttribute(
        "items",
        JSON.stringify([
          {
            text: t(
              kind === "header" ? "ribbon.opt.edit-header" : "ribbon.opt.edit-footer",
              this.host.element(),
            ),
            value: "edit",
          },
          {
            text: t(
              kind === "header" ? "ribbon.opt.remove-header" : "ribbon.opt.remove-footer",
              this.host.element(),
            ),
            value: kind === "header" ? "remove-header" : "remove-footer",
          },
          {
            text: t("ribbon.opt.different-first", this.host.element()),
            value: "title-page",
            checked: !!sp?.titlePage,
          },
          {
            text: t("ribbon.opt.odd-even", this.host.element()),
            value: "odd-even",
            checked: !!oddEven,
          },
        ]),
      );
    };
    stamp("header");
    stamp("footer");
    // While a story is open the chrome re-stamp may have rebuilt the ribbon —
    // re-hang the context tab and mirror the flags into its checkboxes.
    if (this.host.stories().kind() != null) {
      this.showHeaderFooterContextTab();
      const titleCb = this.host
        .root()
        ?.querySelector('docen-ribbon-checkbox[event="header-option"][value="title-page"]');
      titleCb?.toggleAttribute("checked", !!sp?.titlePage);
      const oddEvenCb = this.host
        .root()
        ?.querySelector('docen-ribbon-checkbox[event="header-option"][value="odd-even"]');
      oddEvenCb?.toggleAttribute("checked", !!oddEven);
    }
  }

  /** Word's Header & Footer Tools — append the contextual tab while a story
   *  is open and activate it (Word drops you on the tab); idempotent across
   *  chrome re-stamps. */
  showHeaderFooterContextTab(): void {
    const root = this.host.root();
    const tablist = root?.querySelector("fluent-tablist");
    const ribbon = root?.querySelector("docen-ribbon");
    if (!root || !tablist || !ribbon) return;
    if (tablist.querySelector("#header-footer-tab")) return;
    const scope: Element =
      (root.querySelector("docen-workspace") as Element | null) ?? this.host.element();
    const built = buildContextualTab(headerFooterContextTab(), scope);
    tablist.append(built.tab);
    ribbon.append(built.panel);
    tablist.setAttribute("activeid", "header-footer-tab");
    this.applyRibbonGreying();
  }

  hideHeaderFooterContextTab(): void {
    const root = this.host.root();
    const tablist = root?.querySelector("fluent-tablist");
    const ribbon = root?.querySelector("docen-ribbon");
    if (!root || !tablist || !ribbon) return;
    if (!tablist.querySelector("#header-footer-tab")) return;
    if (tablist.getAttribute("activeid") === "header-footer-tab")
      tablist.setAttribute("activeid", DEFAULT_RIBBON_TAB);
    tablist.querySelector("#header-footer-tab")?.remove();
    ribbon.querySelector('docen-ribbon-panel[value="header-footer-tab"]')?.remove();
    this.applyRibbonGreying();
  }

  /** Mirror the caret cell's live width/height into the Cell Size combos —
   *  Word behavior: the boxes report the selection's column width and row
   *  height (in the locale's unit system), not a fixed default. Runs on every
   *  chrome re-stamp and transaction (via #setupFontSync). */
  syncCellSize(): void {
    const root = this.host.root();
    const widthEl = root?.querySelector('docen-ribbon-combobox[event="cell-width"]');
    const heightEl = root?.querySelector('docen-ribbon-combobox[event="cell-height"]');
    const editor = this.host.editor();
    if ((!widthEl && !heightEl) || !editor) return;
    const anchor = tableAncestry(editor.state);
    if (!anchor) return;
    const scope = root?.querySelector("docen-workspace") ?? this.host.element();
    const { $from } = editor.state.selection;
    if (widthEl) {
      const widths = ($from.node(anchor.tableAt).attrs as { columnWidths?: number[] | null })
        .columnWidths;
      const col = $from.index(anchor.rowAt);
      const tw = widths != null && col < widths.length ? widths[col] : undefined;
      widthEl.setAttribute("value", tw != null ? formatMeasureTwip(tw, scope) : "");
    }
    if (heightEl) {
      const h = ($from.node(anchor.rowAt).attrs as { height?: { value?: number } | null }).height;
      heightEl.setAttribute(
        "value",
        h?.value != null ? formatMeasureTwip(h.value, scope) : useCmUnits(scope) ? "自动" : "auto",
      );
    }
  }

  /** Mirror the selected drawing's live width/height into the Picture/Shape
   *  Format Size combos — Word behavior: the boxes report the selection's
   *  extent in the locale's unit system and follow every resize. Runs on
   *  every chrome re-stamp and transaction (via #setupFontSync). */
  syncDrawingSize(): void {
    const root = this.host.root();
    const widthEl = root?.querySelector('docen-ribbon-input[event="drawing-width"]');
    const heightEl = root?.querySelector('docen-ribbon-input[event="drawing-height"]');
    const editor = this.host.editor();
    if ((!widthEl && !heightEl) || !editor) return;
    const sel = editor.state.selection;
    if (!(sel instanceof NodeSelection)) return;
    const name = sel.node.type.name;
    const attrs = sel.node.attrs as Record<string, unknown>;
    // An image's extent lives in px attrs (15 tw to the px at 96 DPI); a
    // shape/group/chart's in its payload transformation EMU (635 to the tw).
    let pair: { w: unknown; h: unknown; tw: (v: number) => number } | null = null;
    if (name === "image") pair = { w: attrs.width, h: attrs.height, tw: (v) => v * 15 };
    else if (name === "wpsShape" || name === "wpgGroup" || name === "chart") {
      const t = (attrs[name] as Record<string, unknown> | undefined)?.transformation as
        | Record<string, unknown>
        | undefined;
      pair = { w: t?.width, h: t?.height, tw: (v) => v / 635 };
    }
    if (!pair) return;
    const scope = root?.querySelector("docen-workspace") ?? this.host.element();
    const show = (el: Element | null | undefined, v: unknown): void => {
      el?.setAttribute(
        "value",
        typeof v === "number" && v > 0 ? formatMeasureTwip(pair!.tw(v), scope) : "",
      );
    };
    show(widthEl, pair.w);
    show(heightEl, pair.h);
  }

  /** Word's Table Tools + Equation Tools — append/remove the contextual tabs as
   *  the selection enters/leaves their owning context (a table, or a math
   *  atom). Runs per transaction (via #setupFontSync) and after every chrome
   *  re-stamp (#renderChrome, which clears the tracking set because the ribbon
   *  DOM was rebuilt). */
  syncContextTabs(): void {
    const root = this.host.root();
    const tablist = root?.querySelector("fluent-tablist");
    const ribbon = root?.querySelector("docen-ribbon");
    if (!root || !tablist || !ribbon) return;
    const scope: Element =
      (root.querySelector("docen-workspace") as Element | null) ?? this.host.element();
    // The tab ids the current selection calls for (a picture selection and a
    // math selection never coexist; a picture inside a table keeps both —
    // Word's Table Tools stay up with Picture Tools).
    const want = new Map<string, RibbonTab>();
    if (this.host.editor()) {
      const state = this.host.editor()!.state;
      const drawing = drawingSelectionKind(state);
      if (drawing === "picture") want.set("picture-format", pictureFormatTab());
      else if (drawing === "chart") want.set("chart-design", chartDesignTab());
      else if (drawing) want.set("shape-format", shapeFormatTab());
      if (tableAncestry(state)) for (const tab of tableContextTabs(scope)) want.set(tab.id, tab);
      else if (mathAtomAt(state)) want.set("equation", equationContextTab());
    }
    const present = this.#contextTabIds;
    const changed = want.size !== present.size || [...want.keys()].some((id) => !present.has(id));
    if (!changed) return;
    // Retire the tabs whose context the selection left (activeid first, so
    // the tablist never holds an id with no matching tab).
    const active = tablist.getAttribute("activeid") ?? "";
    if (present.has(active) && !want.has(active))
      tablist.setAttribute("activeid", DEFAULT_RIBBON_TAB);
    for (const id of present) {
      if (want.has(id)) continue;
      // A retiring panel's Fluent controls may be mid-teardown in their own
      // blur handlers (clicking away from a combobox detaches its popover,
      // then the selection transaction lands here) — the removal races that
      // cleanup, so a node lost along the way is fine. The tab goes last so
      // it still retires even when the panel's teardown throws.
      try {
        ribbon.querySelector(`docen-ribbon-panel[value="${id}"]`)?.remove();
      } catch {
        /* the blur handler already tore it down */
      }
      try {
        tablist.querySelector(`#${id}`)?.remove();
      } catch {
        /* the blur handler already tore it down */
      }
    }
    // Append the fresh arrivals and activate them (Word drops you on the tab).
    let firstNew: string | null = null;
    for (const [id, tab] of want) {
      if (present.has(id)) continue;
      const built = buildContextualTab(tab, scope);
      tablist.append(built.tab);
      ribbon.append(built.panel);
      firstNew = firstNew ?? id;
    }
    if (firstNew) tablist.setAttribute("activeid", firstNew);
    present.clear();
    for (const id of want.keys()) present.add(id);
    if (want.has("table-design")) this.syncTableLookCheckboxes();
    this.applyRibbonGreying();
  }

  /** Sync Table Style Options checkboxes (w:tblLook) with the active table. */
  syncTableLookCheckboxes(): void {
    const state = this.host.editor()?.state;
    if (!state) return;
    const anchor = tableAncestry(state);
    if (!anchor) return;
    const table = state.selection.$from.node(anchor.tableAt);
    const look = resolveTableLook(table.attrs.tableLook);
    const root = this.host.root();
    if (!root) return;
    const flags = ["firstRow", "lastRow", "bandRow", "firstCol", "lastCol", "bandCol"] as const;
    for (const flag of flags) {
      const cb = root.querySelector<HTMLElement>(
        `docen-ribbon-checkbox[event="toggle-table-look"][value="${flag}"]`,
      );
      cb?.toggleAttribute("checked", !!look[flag]);
    }
  }

  /** The full set of wired command names (Tiptap dispatch + locally handled +
   *  addin commands). External add-ins register non-Tiptap actions (e.g. open a
   *  URL) via `commands`; their keys count as wired so {@link #applyRibbonGreying}
   *  doesn't disable the controls that dispatch them. */
  wiredCommands(): Set<string> {
    const wired = new Set<string>([...WIRED_DISPATCH, ...LOCAL_HANDLED]);
    for (const addin of this.host.addins()) {
      if (!addin.commands) continue;
      for (const key of Object.keys(addin.commands)) wired.add(key);
    }
    return wired;
  }

  /** Dispatch a cancelable event; returns true when a host preventDefaulted it
   *  (i.e. took over the action). Lets save/open/print/new work out-of-box yet
   *  stay overridable. */
  emitCancelable(
    name:
      | "docen:save"
      | "docen:save-as"
      | "docen:open"
      | "docen:new"
      | "docen:print"
      | "docen:close",
    detail?: { format?: SaveFormat },
  ): boolean {
    const event = new CustomEvent(name, {
      bubbles: true,
      composed: true,
      cancelable: true,
      detail,
    });
    this.host.dispatch(event);
    return event.defaultPrevented;
  }

  /** Toggle a task pane open/closed (ribbon View → toggle-navigation). */
  togglePane(id: TaskPaneId): void {
    this.host.setTaskpane(id, !this.host.getTaskpaneState(id));
  }
}
