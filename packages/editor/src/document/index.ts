/**
 * `<docen-document>` — a turnkey DOCX editor web component.
 *
 * Wires the Fluent UI host (title-bar + ribbon + document-area) to the canvas
 * route: a viewless Tiptap engine (the single source of truth for content and
 * commands) driving the layout pipeline (compile → project → layout → LeaferJS
 * paint) on every transaction. The title bar drives file I/O (open/save) and
 * language switching, ribbon commands route to the engine, and file I/O goes
 * through `parseDOCX`/`generateDOCX`. The title bar + ribbon re-render on
 * locale change.
 */

import {
  convertMillimetersToTwip,
  decodePassthroughData,
  detectUnsupportedContent,
  docxExtensions,
  effectiveRunProps,
  generateDOCX,
  normalizeDocument,
  parseDOCX,
  prepareDocument,
  selectionSlicePayload,
  type HtmlGenerateOptions,
  type JSONContent,
  type SectionPropertiesOptions,
  type StylesOptions,
  type DocxVariant,
} from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import {
  type ProjectedFlowBox,
  type ProjectedPageBackground,
  type ProjectedPageFurniture,
  type ProjectedSection,
} from "@docen/docx/layout";
import {
  browserFontMetrics,
  createMeasurer,
  loadDefaultFonts,
  registerShapingFont,
  type FlowPage,
  type FlowPageInsets,
  type RegisterDefaultFontsOptions,
} from "@docen/layout";
import { initShapingWasm } from "@docen/shaping";
import { attr, customElement } from "@microsoft/fast-element";
import type { Mark, Node as PMNode } from "@tiptap/pm/model";
import { EditorState, NodeSelection, TextSelection, type Transaction } from "@tiptap/pm/state";

import { descendGroupChild, drawingNodePos } from "../drawing";
import type { DocenAddin } from "../ui";
import {
  AddinHost,
  applyTheme,
  notifyLocaleChange,
  observeLang,
  registerComponents,
  resolveDir,
  resolveTheme,
  setUiDirection,
  t,
  type RibbonMenuItem,
} from "../ui";
import type { AutocorrectDialogValues } from "../ui/components/workspace/autocorrect-dialog";
import type { BookmarkItem } from "../ui/components/workspace/bookmark-dialog";
import type { DrawingPropertiesState } from "../ui/components/workspace/drawing-properties-dialog";
import type { FontDialogPatch } from "../ui/components/workspace/font-dialog";
import type { GoToKind, GoToPayload } from "../ui/components/workspace/go-to-dialog";
import type { LinkValues } from "../ui/components/workspace/link-dialog";
import type { DocenMiniToolbar } from "../ui/components/workspace/mini-toolbar";
import type {
  NoteKindSettings,
  NoteSettingsValues,
} from "../ui/components/workspace/note-settings-dialog";
import type { PasteSpecialFormat } from "../ui/components/workspace/paste-special-dialog";
import type {
  DocumentPropertiesCore,
  DocumentPropertiesStats,
} from "../ui/components/workspace/properties-dialog";
import type { QuickPartValues } from "../ui/components/workspace/quick-part-dialog";
import type { FormattingInfo } from "../ui/components/workspace/reveal-formatting-pane";
import type { SdtPropertiesValues } from "../ui/components/workspace/sdt-dialog";
import { createDefaultAddin, textCounter, wordCounter } from "./addin";
import { blocksOfDocAttrs, parseSlicePayload, withBlocks } from "./building-blocks";
import { A11yMirror } from "./canvas/a11y-mirror";
import { autocorrectConfigOf } from "./canvas/autocorrect";
import {
  mountEditBridge,
  type EditBridge,
  type StoryKind,
  type StorySlot,
} from "./canvas/edit-bridge";
import { CanvasStage, type CanvasStageSection, type LaidFurnitureSection } from "./canvas/stage";
import { documentStyles, documentTemplate } from "./chrome";
import { ClipboardCommands } from "./commands/clipboard";
// Side-effect: register the document-specific UI components moved out of the
// shared ui/ barrel — <docen-format-pane> (properties fallback),
// <docen-outline> (navigation Headings tab), <docen-styles-pane> (Styles).
import "./components/format-pane";
import "./components/outline";
import "./components/styles-pane";
import "../ui/components/workspace/alt-text-pane";
import { CommentsCommands } from "./commands/comments";
import { combineDocs, compareDocs } from "./commands/compare";
import { DesignCommands } from "./commands/design";
import { DialogCommands, updateDynamicFieldsBeforePrint } from "./commands/dialogs";
import { hostCommands, type HostCommandRegistry } from "./commands/host";
import {
  BuildingBlocksHostCommands,
  type BuildingBlocksHostView,
} from "./commands/host/building-blocks";
import { ReadAloudController } from "./commands/host/read-aloud";
import { SdtCommands, SdtHostCommands } from "./commands/host/sdt";
import { applyRecipientsRow, MailMergeCommands } from "./commands/mail-merge";
import { NavigationCommands } from "./commands/navigation";
import { selectSimilarFormatting } from "./commands/outline";
import { ReferencesCommands } from "./commands/references";
import { RevisionsCommands } from "./commands/revisions";
import { SectionCommands } from "./commands/sections";
import { SpellingCommands } from "./commands/spelling";
import { THEMES } from "./commands/themes";
import { contentWarningText } from "./content-warning";
import type { NewStyleDefinition } from "./extensions/commands";
import type { ModifyStylePatch, ParagraphDialogPatch } from "./extensions/commands";
import { type FieldFrame } from "./fields";
import { READONLY_LIVE, type SaveFormat } from "./file-formats";
import { ChromeDomain } from "./host/chrome";
import { InsertDomain } from "./host/insert";
import { IODomain } from "./host/io";
import { RenderDomain } from "./host/render";
import { StatusDomain } from "./host/status";
import { pageInsets, StoriesDomain } from "./host/stories";
// Side-effect import: registers the ribbon/header translation tables.
import "./i18n";
import { StylesDomain } from "./host/styles";
import { mergeSectionProperties } from "./page-setup";
import { compressPictureSrc, pickTransparentColor, type CropRect } from "./pixels";
import {
  addPermissionRange,
  applyProtectionMode,
  enforceProtection,
  findNextPermissionRange,
  hasPermissionRanges,
  isInsideEditableField,
  isInsideEditablePermission,
  stopProtection,
  withProtection,
  type ProtectionHostView,
  type ProtectionPane,
} from "./protection";
import { downloadOleObject } from "./quick-tables";
import { formattingInfoOf } from "./reveal-formatting";
import { useCmUnits } from "./ribbon";
import {
  AUTOCORRECT_TABLE_VERSION,
  getSettings,
  onSettingsChange,
  resolveIdentity,
  updateSettings,
  type DocenSettings,
  type IdentitySettings,
  type SettingsPatch,
} from "./settings";
import { getSynonyms, spellSuggestions } from "./spelling";
import { attachTemplate, type DotxTemplatePackage } from "./template-manager";
import { translateText } from "./translation";

/** Double-click window (ms) — the format painter's sticky toggle and the
 *  bare-click stroke deferral both track the system double-click time. */
const PAINTER_DOUBLE_CLICK_MS = 500;

/** The style's run definition as the Font dialog's prefill (the Format >
 *  Font open — absent values leave the combo blank, inherit-again). */
function fontPatchOfRun(run: Record<string, unknown>): FontDialogPatch {
  const underline = run.underline as { type?: unknown; color?: unknown } | undefined;
  return {
    font: typeof run.font === "string" ? run.font : null,
    size: typeof run.size === "number" ? String(run.size) : null,
    bold: run.bold === true,
    italic: run.italic === true,
    underlineStyle: typeof underline?.type === "string" ? underline.type : null,
    underlineColor: typeof underline?.color === "string" ? underline.color : null,
    strike: run.strike === true,
    doubleStrike: run.doubleStrike === true,
    superscript: run.vertAlign === "superscript",
    subscript: run.vertAlign === "subscript",
    smallCaps: run.smallCaps === true,
    allCaps: run.caps === true,
    hidden: run.vanish === true,
    shadow: run.shadow === true || Boolean(run.shadow),
    outline: run.outline === true || Boolean(run.outline),
    emboss: run.emboss === true,
    imprint: run.imprint === true,
  };
}

/** The Font dialog's commit as a style-definition run block — the Format >
 *  Font write maps dialog fields onto the w:rPr children (unchecked = the
 *  field drops out, the style inherits again). */
function fontRunPropsOf(patch: FontDialogPatch): Record<string, unknown> {
  const size = patch.size ? Number(patch.size) : undefined;
  return {
    font: patch.font ?? undefined,
    size,
    sizeComplexScript: size,
    bold: patch.bold || undefined,
    italic: patch.italic || undefined,
    underline: patch.underlineStyle
      ? {
          type: patch.underlineStyle,
          ...(patch.underlineColor ? { color: patch.underlineColor } : {}),
        }
      : undefined,
    strike: patch.strike || undefined,
    doubleStrike: patch.doubleStrike || undefined,
    smallCaps: patch.smallCaps || undefined,
    caps: patch.allCaps || undefined,
    vertAlign: patch.superscript ? "superscript" : patch.subscript ? "subscript" : undefined,
    vanish: patch.hidden || undefined,
    shadow: patch.shadow || undefined,
    outline: patch.outline || undefined,
    emboss: patch.emboss || undefined,
    imprint: patch.imprint || undefined,
  };
}

/**
 * Task pane identifiers, mirroring the Office `<TaskpaneId>` concept. The host
 * ships two built-in panes: `navigation` (start/left) and `properties` (end/right).
 */
export type TaskPaneId =
  | "navigation"
  | "properties"
  | "comments"
  | "clipboard"
  | "proofing"
  | "thesaurus"
  | "translate"
  | "revisions"
  | "styles"
  | "reveal"
  | "restrict"
  | "a11y"
  | "altText";

/**
 * Visibility mode values, matching `Office.VisibilityMode` (`taskpane` | `hidden`).
 * Carried on {@link docen:taskpane-visibility-change} event details.
 */
export type VisibilityMode = "taskpane" | "hidden";

/** The interactive horizontal ruler element (View → Ruler): the
 *  `<docen-ruler>` component surface the host drives. */
interface InteractiveRulerElement extends HTMLElement {
  bindEditor(editor: Editor): void;
  setParagraphAttrs(
    indent?: { left?: number; right?: number; firstLine?: number; hanging?: number },
    tabStops?: Array<{ position: number; type: string }>,
    geometry?: {
      pageWidthPx?: number;
      marginLeftPx?: number;
      marginRightPx?: number;
      scale?: number;
    },
  ): void;
}

@customElement({ name: "docen-document", template: documentTemplate, styles: documentStyles })
class DocenDocument extends AddinHost<Editor> {
  // ── Reactive attributes (@attr) — no `reflect` (attribute → property stays
  //  one-way). addinsAttr (attribute "addins") dodges AddinHost.addinsChanged
  //  and the `addins` getter.
  @attr editable?: string;
  @attr filename?: string;
  @attr user?: string;
  @attr avatar?: string;
  @attr({ attribute: "section-properties" }) sectionProperties?: string;
  @attr styles?: string;
  @attr({ attribute: "addins" }) addinsAttr?: string;
  @attr theme?: string;
  /** Leafer engine debug overlay — "bounds" | "hit" | "repaint" | "on". */
  @attr debug?: string;
  /** The document view (Word's View tab): "print" | "web" | "draft" | "read".
   *  Anything else falls back to "print". */
  @attr view?: string;
  @attr override dir!: string;
  @attr({ mode: "boolean", attribute: "touch-mode" }) touchMode = false;

  get isRtl(): boolean {
    const d = this.dir || this.getAttribute("dir");
    return d === "rtl" || resolveDir(this) === "rtl";
  }

  dirChanged(): void {
    const d = this.dir || this.getAttribute("dir");
    if (d === "rtl" || d === "ltr") {
      this.#syncDirTo(d);
    } else {
      this.#syncDir();
    }
  }

  touchModeChanged(): void {
    this.#syncTouchMode();
  }

  #syncTouchMode = (): void => {
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    if (ribbon) {
      if (this.touchMode) {
        ribbon.setAttribute("touch-mode", "");
        (ribbon as HTMLElement & { touchMode?: boolean }).touchMode = true;
      } else {
        ribbon.removeAttribute("touch-mode");
        (ribbon as HTMLElement & { touchMode?: boolean }).touchMode = false;
      }
    }
  };

  #syncDirTo = (nextDir: string): void => {
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    if (ribbon && ribbon.getAttribute("dir") !== nextDir) {
      ribbon.setAttribute("dir", nextDir);
    }
    const statusBar = this.shadowRoot?.querySelector("docen-status-bar");
    if (statusBar && statusBar.getAttribute("dir") !== nextDir) {
      statusBar.setAttribute("dir", nextDir);
    }
    const navPane = this.shadowRoot?.querySelector("docen-nav-pane, docen-navigation-pane");
    if (navPane && navPane.getAttribute("dir") !== nextDir) {
      navPane.setAttribute("dir", nextDir);
    }
    const workspace = this.shadowRoot?.querySelector("docen-workspace");
    if (workspace && workspace.getAttribute("dir") !== nextDir) {
      workspace.setAttribute("dir", nextDir);
    }
  };

  #syncDir = (): void => {
    const nextDir = resolveDir(this);
    if (this.getAttribute("dir") !== nextDir) {
      this.setAttribute("dir", nextDir);
    }
    this.#syncDirTo(nextDir);
    this.#syncRulerDirection();
  };

  setUiDirection(direction: "ltr" | "rtl" | "auto"): void {
    setUiDirection(direction);
    this.#syncDir();
  }

  #bridge?: EditBridge;
  /** The Markdown input mode (Options → Markdown) — session-level, like the
   *  spelling toggle: the bridge reads it per keystroke via a getter. */
  #markdown = true;
  #fieldShading: "never" | "always" | "whenSelected" = "whenSelected";
  #updateFieldsBeforePrint = false;
  #printMarkup = false;
  /** Whether the document's settings.xml carries a read-only editing
   *  restriction (Options → Document). Folds into every editable
   *  computation — never a second setEditable writer. */
  #docProtected = false;
  #protectionMode?: string;
  readonly #readAloud = new ReadAloudController();
  /** References-tab commands (citations/bibliography/index marking), split
   *  out of this class — see commands/references.ts. */
  readonly #spelling = new SpellingCommands({
    // Word checks the story being edited, not the body — route through the
    // bridge's active story so run/replace/right-click follow it.
    editor: () => this.#bridge?.activeEditor() ?? this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
  });
  readonly #navigation = new NavigationCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
    setTextSelection: (from, to) => this.#setTextSelection(from, to),
  });
  readonly #design = new DesignCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
  });
  readonly #comments = new CommentsCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
    showTaskpane: (id) => this.showTaskpane(id),
  });
  readonly #revisions = new RevisionsCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
    togglePane: (id) => this.#togglePane(id),
  });
  readonly #references = new ReferencesCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
    updateAllFields: () => this.#dialogs.updateAllFields(),
  });
  /** Mailings-tab merge commands (recipients/merge fields/preview), split out
   *  of this class — see commands/mail-merge.ts. */
  readonly #merge = new MailMergeCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
    rerender: () => this.#renderDoc(this.getJSON()),
  });
  /** Dialog-commit commands (paragraph/font/table/Chinese layout/caption/
   *  cross-reference), split out of this class — see commands/dialogs.ts. */
  readonly #dialogs = new DialogCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
    syncStatusLanguage: () => this.#syncStatusLanguage(),
    filename: () => this.filename,
    fieldFrame: (pos) => this.#fieldFrame(pos),
    positionTerms: () => ({
      above: t("crossRef.above", this),
      below: t("crossRef.below", this),
    }),
  });
  /** "This section" commands (sectPr read/write, page setup presets, the
   *  page-setup/columns/borders dialogs), split out of this class — see
   *  commands/sections.ts. */
  readonly #sections = new SectionCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
    flow: () => this.#flow,
  });
  /** Paste lanes, the paste-options bar, and the Office Clipboard pane,
   *  split out of this class — see commands/clipboard.ts. */
  readonly #clipboard = new ClipboardCommands({
    editor: () => this.editor,
    bridge: () => this.#bridge,
    element: () => this as HTMLElement,
  });
  /** Status bar / zoom / Navigation-pane thumbnails, split out of this class —
   *  see host/status.ts. */
  readonly #status = new StatusDomain({
    root: () => this.shadowRoot,
    editor: () => this.editor,
    bridge: () => this.#bridge,
    stage: () => this.#stage,
    pages: () => this.#pages,
    sectionOfPage: () => this.#sectionOfPage,
    flow: () => this.#flow,
    viewMode: () => this.#viewMode(),
    caretLanguage: () => this.#caretLanguage(),
    taskpaneOpen: (id) => this.getTaskpaneState(id),
    updateReveal: () => this.#updateRevealFormatting(),
    updateAltText: () => this.#syncAltTextPane(),
    setView: (view) => this.setAttribute("view", view),
    emitZoom: (zoom) => {
      this.dispatchEvent(
        new CustomEvent("docen:zoom-change", {
          bubbles: true,
          composed: true,
          detail: { zoom },
        }),
      );
      this.#syncRuler();
    },
  });
  /** Ribbon/title-bar chrome (QAT, auto-save, header/panes, menu syncs), split
   *  out of this class — see host/chrome.ts. */
  readonly #chrome = new ChromeDomain({
    element: () => this as HTMLElement,
    root: () => this.shadowRoot,
    editor: () => this.editor,
    bridge: () => this.#bridge,
    settings: () => this.settings,
    addins: () => this.addins as readonly DocenAddin[],
    addAddin: (addin) => this.addAddin(addin as never),
    removeAddin: (id) => this.removeAddin(id),
    getAttribute: (name) => this.getAttribute(name),
    hasAttribute: (name) => this.hasAttribute(name),
    markdown: () => this.#markdown,
    markupView: () => this.#markupView,
    markupAuthors: () => this.#markupAuthors,
    markupColors: () => this.#markupColors,
    balloons: () => this.#balloons,
    snapshots: () => this.#versionSnapshots,
    setSnapshots: (value) => {
      this.#versionSnapshots = value;
    },
    merge: () => this.#merge,
    stories: () => this.#stories,
    storyKind: () => this.#stories.kind(),
    getTaskpaneState: (id) => this.getTaskpaneState(id),
    setTaskpane: (id, open) => this.#setTaskpane(id, open),
    updateStatus: () => this.#updateStatus(),
    hostCommandEvents: () => {
      const registry = this.#hostCommandRegistry();
      return new Set<string>([...registry.chrome.keys(), ...registry.editor.keys()]);
    },
    dispatch: (event) => this.dispatchEvent(event),
  });
  /** File I/O (open/save/print/close/templates), split out of this class —
   *  see host/io.ts. */
  readonly #io = new IODomain({
    element: () => this as HTMLElement,
    root: () => this.shadowRoot,
    editor: () => this.editor,
    bridge: () => this.#bridge,
    stage: () => this.#stage,
    pages: () => this.#pages,
    sectionOfPage: () => this.#sectionOfPage,
    flow: () => this.#flow,
    lastRun: () => this.#lastRun,
    fonts: () => this.#fonts,
    viewMode: () => this.#viewMode(),
    lang: () => this.lang,
    docxVariant: () => this.#docxVariant,
    setDocxVariant: (variant) => {
      this.#docxVariant = variant;
    },
    docProtected: () => this.#docProtected,
    setDocProtected: (value) => {
      this.#docProtected = value;
    },
    protectionMode: () => this.#protectionMode,
    setProtectionMode: (value) => {
      this.#protectionMode = value;
    },
    updateFieldsOnOpen: () => this.#updateFieldsOnOpen,
    setUpdateFieldsOnOpen: (value) => {
      this.#updateFieldsOnOpen = value;
    },
    jsonDirty: () => this.#jsonDirty,
    setJsonDirty: (value) => {
      this.#jsonDirty = value;
    },
    cachedJSON: () => this.#cachedJSON,
    setCachedJSON: (value) => {
      this.#cachedJSON = value;
    },
    fileInput: () => this.#fileInput,
    getAttribute: (name) => this.getAttribute(name),
    setAttribute: (name, value) => this.setAttribute(name, value),
    emitCancelable: (name, detail) => this.#emitCancelable(name, detail),
    renderChrome: () => this.#renderChrome(),
    renderDoc: (doc) => this.#renderDoc(doc),
    applyDocumentTheme: (kind, value, persist) => this.#applyDocumentTheme(kind, value, persist),
    snapshotStyles: () => this.#snapshotStyles(),
    syncEditable: () => this.#syncEditable(),
    syncDocumentSettings: (settings) => this.#syncDocumentSettings(settings),
  });
  /** Styles pane / gallery / Modify Style dialogs, split out of this class —
   *  see host/styles.ts. */
  readonly #styles = new StylesDomain({
    element: () => this as HTMLElement,
    root: () => this.shadowRoot,
    editor: () => this.editor,
    renderChrome: () => this.#renderChrome(),
    fontPatchOfRun: (run) => fontPatchOfRun(run),
    fontRunPropsOf: (patch) => fontRunPropsOf(patch),
  });
  /** Insert/jump commands (symbols, shapes, pages, equations, tabs, bookmarks,
   *  note jumps), split out of this class — see host/insert.ts. */
  readonly #insert = new InsertDomain({
    element: () => this as HTMLElement,
    root: () => this.shadowRoot,
    editor: () => this.editor,
    bridge: () => this.#bridge,
    flow: () => this.#flow,
    pages: () => this.#pages,
    hyphenation: () => this.#hyphenation,
    setHyphenation: (value) => {
      this.#hyphenation = value;
    },
    merge: () => this.#merge,
    getJSON: () => this.getJSON(),
    setJSON: (json) => this.setJSON(json),
    renderDoc: (doc) => this.#renderDoc(doc),
    setTextSelection: (from, to) => this.#setTextSelection(from, to),
    openBookmarkDialog: () => this.#openBookmarkDialog(),
  });
  /** Canvas render pipeline (project -> layout -> paint, field feedback, page
   *  diff), split out of this class - see host/render.ts. */
  readonly #render = new RenderDomain({
    root: () => this.shadowRoot,
    editor: () => this.editor,
    bridge: () => this.#bridge,
    stage: () => this.#stage,
    stageHost: () => this.#stageHost,
    armStage: (projected) => this.#armStage(projected),
    measurer: () => this.#measurer,
    a11yMirror: () => this.#a11yMirror,
    a11yTimer: () => this.#a11yTimer,
    setA11yTimer: (value) => {
      this.#a11yTimer = value;
    },
    comments: () => this.#comments,
    revisions: () => this.#revisions,
    spelling: () => this.#spelling,
    dialogs: () => this.#dialogs,
    getJSON: () => this.getJSON(),
    getTaskpaneState: (id) => this.getTaskpaneState(id),
    hasAttribute: (name) => this.hasAttribute(name),
    isConnected: () => this.isConnected,
    filename: () => this.filename,
    debug: () => this.debug,
    hyphenation: () => this.#hyphenation,
    markupView: () => this.#markupView,
    markupAuthors: () => this.#markupAuthors,
    markupColors: () => this.#markupColors,
    balloons: () => this.#balloons,
    hiddenTextShown: () => this.#hiddenTextShown,
    setHiddenTextShown: (value) => {
      this.#hiddenTextShown = value;
    },
    fieldCodes: () => this.#fieldCodes,
    setFieldCodes: (value) => {
      this.#fieldCodes = value;
    },
    updateFieldsOnOpen: () => this.#updateFieldsOnOpen,
    setUpdateFieldsOnOpen: (value) => {
      this.#updateFieldsOnOpen = value;
    },
    documentSettings: () => this.#documentSettings(),
    pages: () => this.#pages,
    setPages: (pages) => {
      this.#pages = pages;
    },
    sectionOfPage: () => this.#sectionOfPage,
    setSectionOfPage: (sectionOfPage) => {
      this.#sectionOfPage = sectionOfPage;
    },
    flow: () => this.#flow,
    setFlow: (flow) => {
      this.#flow = flow;
      // The flow box is the ruler's geometry source — sync as soon as it lands
      // (the first render is incremental, so renderDoc's own call may run
      // before the first slice has laid out).
      this.#syncRuler();
    },
    lastRun: () => this.#lastRun,
    setLastRun: (run) => {
      this.#lastRun = run;
    },
    renderSeq: () => this.#renderSeq,
    setRenderSeq: (seq) => {
      this.#renderSeq = seq;
    },
    mergedView: (doc) => this.#mergedView(doc),
    pageOriginOf: (sections, sectionOfPage) => this.#pageOriginOf(sections, sectionOfPage),
    pageInsets: (flow, furniture, laid) => this.#pageInsets(flow, furniture, laid),
    updateStatus: () => this.#updateStatus(),
    syncStatusLanguage: () => this.#syncStatusLanguage(),
    setProgress: (label) => this.#setProgress(label),
    viewMode: () => this.#viewMode(),
  });
  #stage?: CanvasStage;
  /** Header/footer story editing, split out of this class — see
   *  host/stories.ts. */
  readonly #stories = new StoriesDomain({
    element: () => this as HTMLElement,
    editor: () => this.editor,
    bridge: () => this.#bridge,
    stage: () => this.#stage,
    sectionOfPage: () => this.#sectionOfPage,
    projectAndLayout: (doc) => this.#projectAndLayout(doc),
    setLayout: (run) => {
      this.#pages = run.pages;
      this.#sectionOfPage = run.sectionOfPage;
      this.#flow = run.sections[0]?.flow;
    },
    loadDoc: (doc) => this.#loadDoc(doc),
    sectPrPosInCurrentSection: () => this.#sections.sectionSectPrPos(),
    hideContextTab: () => this.#hideHeaderFooterContextTab(),
    spellingSchedule: () => this.#spelling.schedule(),
    spellingRun: () => this.#spelling.run(),
    isPageField: (child) => DocenDocument.isPageField(child),
  });
  #stageHost?: HTMLElement;
  /** The interactive horizontal ruler (View → Ruler). Mounted above the
   *  pages; the vertical strip stays on the stage. */
  #ruler?: InteractiveRulerElement;
  #rulerGeometry = "";
  #rulerEditor?: Editor;
  readonly #a11yMirror = new A11yMirror();
  #a11yTimer?: number;
  #versionSnapshots: Array<{
    id: string;
    timestamp: string;
    author: string;
    isAutosave: boolean;
    doc: JSONContent;
  }> = [];
  #currentLandmarkIdx = 0;
  #measurer = createMeasurer(browserFontMetrics);
  /** Host-registered font bytes by lowercased family — used for shaping
   *  (when opted in) and for PDF/DOCX font embedding. The original family
   *  spelling rides along for font-name output. */
  readonly #fonts = new Map<string, { family: string; fontData: Uint8Array }>();
  /** True while the user dismissed the unsupported-content warning for the
   *  current document (reset on every load). */
  #contentWarningDismissed = false;
  #pages: readonly FlowPage[] = [];
  /** Page index → section index (the caret's section and per-page geometry
   *  read through it). */
  #sectionOfPage: readonly number[] = [];
  /** The first section's flow box (page-width presets / TOC tab position;
   *  a multi-section refinement reads the caret's own section). */
  #flow?: ProjectedFlowBox;
  #fileInput?: HTMLInputElement;
  #imageInput?: HTMLInputElement;
  #pictureInput?: HTMLInputElement;
  /** Pending frame for the coalesced host UI sync — one chrome pass per
   *  animation frame instead of one per transaction (see #scheduleUiSync). */
  #uiSyncFrame = 0;
  /** A pending UI sync still owes the status-bar language a refresh. */
  #uiSelectionDirty = false;
  #unobserveLang?: () => void;
  /** Tears down the shared settings-store subscription (header re-stamp +
   *  `docen:settings-change` forwarding). */
  #settingsOff?: () => void;
  /** Watches the host's `lang` attribute and forwards it to the internal
   *  <docen-workspace> + notifies locale observers. MutationObserver because
   *  @attr `lang` clashes with HTMLElement.lang (TS2416); manual
   *  observedAttributes would break FASTElement's @attr dispatch. */
  #langObserver?: MutationObserver;
  /** Tears down the transaction listener mirroring caret font/size → comboboxes. */
  #fontSyncCleanup?: () => void;
  // Format Painter captured formatting + the listeners that apply it. Marks
  // and paragraph attrs capture together; sticky mode (double click) paints
  // every following selection until Esc or another painter click.
  #painterMarks: readonly Mark[] | null = null;
  #painterPara: Record<string, unknown> | null = null;
  #painterOff?: () => void;
  #painterKeyOff?: () => void;
  #painterSticky = false;
  #painterClickAt = 0;
  #painterStrokeTimer?: ReturnType<typeof setTimeout>;
  #lastMiniToolbarSelection: { from: number; to: number } | null = null;
  #previewSnapshot: {
    docContent: unknown;
    selection: { from: number; to: number };
    themeKind?: string;
    themeCssVars?: Map<string, string>;
  } | null = null;
  /** Cached unwrapped JSON (host.getJSON result). Invalidated on every user/doc
   *  change; recomputed lazily. Saves the editor.getJSON walk on every
   *  save/autosave/getJSON call. */
  #cachedJSON?: JSONContent;
  #jsonDirty = true;
  /** The docx-family variant the open document belongs to (docm/dotx/dotm
   *  opened via their extension, docx otherwise). Save keeps the document's
   *  variant — main content type, extension, picker MIME — so a macro-enabled
   *  or template document round-trips as itself instead of being mislabelled. */
  #docxVariant: DocxVariant = "docx";
  /** Word's Display for Review state (Review → Tracking): how tracked changes
   *  project onto the canvas, optionally scoped to one reviewer's revisions.
   *  Pure display state — the marks in the document are untouched. */
  #markupView: "simple" | "all" | "none" | "original" = "simple";
  #markupAuthors: string[] | null = null;
  /** Review → Markup Colors: Word's By-author palette (default) or the fixed
   *  per-change-type colors. Display-only, like the view/filter above. */
  #markupColors: "author" | "changeType" = "author";
  /** Review → Show Markup → Balloons: which annotations project into the
   *  page-margin balloon stack (comments / revisions / both / none). The
   *  default matches Word's fresh document — the margin shows what the
   *  document carries. */
  #balloons: "all" | "comments" | "revisions" | "none" = "all";
  /** Word's field-code display (Alt+F9): projects every field as its
   *  instruction text instead of the cached result. Pure display state — the
   *  document's field atoms are untouched. */
  #fieldCodes = false;
  /** The hidden-text display the last projection used (the settings store's
   *  writing.showHiddenText): a store change to it re-renders the document,
   *  Word's Options → Display toggle. */
  #hiddenTextShown = false;
  /** A loaded document asked for w:updateFields (Options → Update fields on
   *  open): consumed by the first completed render, when the bridge's page
   *  map is fresh, to run one Update All Fields. */
  #updateFieldsOnOpen = false;
  /** Table Design → Draw Border: the pen in tcBorders form (style token,
   *  size in eighth-points, color) plus the armed paint/erase mode. While
   *  armed the canvas presses sweep table edges instead of selecting. */
  #pen: { style: string; size: number; color: string } = {
    style: "single",
    size: 4,
    color: "auto",
  };
  #borderPainting = false;
  #borderErase = false;
  #borderPaintKeyOff?: () => void;

  #tableDrawing = false;
  #tableEraser = false;
  #tableDrawKeyOff?: () => void;
  #tableEraserKeyOff?: () => void;

  /** Insert → Shapes: the armed preset token (null = disarmed). While armed
   *  the canvas presses drag a ghost rectangle and insert the preset at it
   *  (Word's drag-to-draw); Esc disarms, draws keep it armed. */
  #armedShape: string | null = null;
  #armedShapeKeyOff?: () => void;
  #hyphenation: {
    auto?: boolean;
    doNotHyphenateCaps?: boolean;
    zoneTw?: number;
    limit?: number;
  } = {};

  /** The underlying Tiptap Editor (undefined before connect / after disconnect).
   *  Exposed so a host (the @docen/vue adapter, or any parent element) can drive
   *  commands programmatically — setContent / getJSON / chain / ... — without
   *  routing through the ribbon. */
  get editor(): Editor | undefined {
    return this.#bridge?.editor;
  }

  /** DocenHost surface — bridge the editor-agnostic `unknown` content contract
   *  to the typed {@link getJSON} / {@link setJSON} API. Addins (and any
   *  DocenHost consumer) read/write content through here without knowing the
   *  runtime is Tiptap JSON. */
  getContent(): unknown {
    return this.getJSON();
  }

  setContent(content: unknown): void {
    if (content && typeof content === "object") {
      this.setJSON(content as JSONContent);
    }
  }

  // ── @attr change callbacks — every handler is guarded (bridge/shadowRoot
  //  check) so an early fire during FAST's attribute hydration is a no-op.
  editableChanged(): void {
    this.#bridge?.editor.setEditable(this.editable !== "false");
    this.#syncEditModeMenu();
  }

  filenameChanged(): void {
    this.#renderChrome();
  }

  userChanged(): void {
    this.#renderChrome();
  }

  avatarChanged(): void {
    this.#renderChrome();
  }

  sectionPropertiesChanged(): void {
    this.#applySectionPropertiesAttr();
  }

  stylesChanged(): void {
    this.#applyStylesAttr();
  }

  viewChanged(): void {
    this.#applyView();
  }

  addinsAttrChanged(): void {
    this.#applyAddinsAttr();
  }

  themeChanged(): void {
    this.#applyThemeAttr(this.theme ?? "");
  }

  debugChanged(): void {
    this.#stage?.setDebug(this.debug);
  }

  /** Esc fallback: restore the ribbon to "always shown" after the browser leaves fullscreen. */
  /** ribbon-mode-change → drive browser fullscreen + status-bar hide.
   *  auto-hide = Full Screen (Office); any other mode exits it. Named so it can
   *  be removed on disconnect (an anonymous listener would leak on reconnect). */
  readonly #onRibbonModeChange = (event: Event): void => {
    const workspace = this.shadowRoot?.querySelector("docen-workspace");
    if (!workspace) return;
    const mode = (event as CustomEvent<{ mode: string }>).detail.mode;
    if (mode === "auto-hide") {
      void this.requestFullscreen?.().catch(() => {});
      workspace.setAttribute("data-fullscreen", "");
    } else {
      if (document.fullscreenElement) void document.exitFullscreen?.().catch(() => {});
      workspace.removeAttribute("data-fullscreen");
    }
  };

  readonly #onFullscreenChange = (): void => {
    if (document.fullscreenElement) return;
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    if (ribbon) ribbon.setAttribute("mode", "always-shown");
  };

  /** Ctrl+wheel zoom over the page area (Word/Office behavior). Captured ahead
   *  of the stage shell's wheel handling, which stops propagation for its own
   *  scroll — plain wheel keeps scrolling; only the Ctrl chord zooms. */
  readonly #onWheel = (event: WheelEvent): void => {
    if (!event.ctrlKey && !event.metaKey) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    this.#setZoom(this.#status.getZoom() + (event.deltaY < 0 ? 10 : -10));
  };

  /** Ctrl+= / Ctrl+- / Ctrl+0 zoom, Ctrl+F find (Word behavior). Zoom is
   *  ignored inside ribbon comboboxes and other inputs (so the keystroke reaches
   *  them); Ctrl+F is global. preventDefault blocks the browser's native zoom/find. */
  readonly #onZoomKey = (event: KeyboardEvent): void => {
    if (event.defaultPrevented) return;
    // Alt+Q focuses the command search (Office's "Tell me what you want to
    // do" shortcut). Handled before the Ctrl/Meta gate below.
    if (
      event.altKey &&
      !event.ctrlKey &&
      !event.metaKey &&
      (event.key === "q" || event.key === "Q")
    ) {
      event.preventDefault();
      const search = this.shadowRoot?.querySelector("docen-command-search") as HTMLElement | null;
      search?.focus();
      return;
    }
    // Ctrl+Shift+S focuses the Styles gallery (Word's Apply Styles shortcut).
    if (
      (event.ctrlKey || event.metaKey) &&
      event.shiftKey &&
      (event.key === "s" || event.key === "S")
    ) {
      event.preventDefault();
      (
        this.shadowRoot?.querySelector('docen-ribbon-gallery[event="style"]') as HTMLElement | null
      )?.focus();
      return;
    }
    // F12 = Save As (Word).
    if (event.key === "F12") {
      event.preventDefault();
      void this.#saveAs();
      return;
    }
    // F6 / Shift+F6 = Cyclical landmark navigation (Ribbon <-> Stage <-> Task Pane <-> Status Bar)
    if (event.key === "F6") {
      event.preventDefault();
      this.#cycleLandmarks(event.shiftKey);
      return;
    }
    // F7 = Spelling & Grammar (Word).
    if (event.key === "F7") {
      event.preventDefault();
      this.#onCommand(new CustomEvent("command", { detail: { event: "spell-check" } }));
      return;
    }
    // Alt+= inserts a blank inline equation (Word's Insert Equation shortcut).
    if (event.altKey && !event.ctrlKey && !event.metaKey && event.key === "=") {
      event.preventDefault();
      this.#insertEquation("plain");
      return;
    }
    // F9 updates the field at the caret; Alt+F9 toggles field-code display
    // (Word's Update Field / View Field Codes).
    if (event.key === "F9") {
      event.preventDefault();
      if (event.altKey) this.toggleFieldCodes();
      else this.#dialogs.fieldUpdateAtSelection();
      return;
    }
    if (!(event.ctrlKey || event.metaKey)) return;
    // Ctrl+F1 toggles ribbon minimized mode (Word behavior).
    if (event.key === "F1") {
      event.preventDefault();
      this.#toggleRibbonMinimized();
      return;
    }
    // Ctrl+Shift+8 toggles formatting marks (Word). Shift+8 turns the key
    // into "*" on US layouts, so both spellings count.
    if (event.shiftKey && (event.key === "8" || event.key === "*")) {
      event.preventDefault();
      this.setShowMarks(!this.getShowMarks());
      return;
    }
    // Ctrl+F opens Find, Ctrl+H opens Find & Replace (Word behavior).
    if (event.key === "f" || event.key === "F") {
      event.preventDefault();
      this.#navigation.openSearch();
      return;
    }
    if (event.key === "h" || event.key === "H") {
      event.preventDefault();
      this.#navigation.openFindReplace();
      return;
    }
    // Ctrl+G opens Go To dialog (Word behavior).
    if (event.key === "g" || event.key === "G") {
      event.preventDefault();
      this.#openGoToDialog();
      return;
    }
    // Ctrl+S saves (Word) — before the input gate, a save applies everywhere.
    if (event.key === "s" || event.key === "S") {
      event.preventDefault();
      if (!this.#emitCancelable("docen:save")) void this.#saveAs();
      return;
    }
    // Ctrl+P prints the canvas pages (Word) — not the browser's DOM print.
    if (event.key === "p" || event.key === "P") {
      event.preventDefault();
      if (!this.#emitCancelable("docen:print")) void this.#print();
      return;
    }
    // composedPath()[0] is the real target inside the shadow DOM (e.g. a combobox input).
    const target = event.composedPath()[0] as HTMLElement | null;
    if (
      target instanceof HTMLElement &&
      target.closest("input, textarea, docen-ribbon-combobox") &&
      // The bridge textarea IS the document input — zoom must stay live with
      // the caret in the document.
      !target.closest("[data-docen-bridge-input]")
    )
      return;
    // Ctrl+K opens the Link dialog (Word) — after the input gate, so a
    // combobox keystroke is never hijacked.
    if (event.key === "k" || event.key === "K") {
      event.preventDefault();
      this.#insertLink();
      return;
    }
    const key = event.key;
    if (key === "+" || key === "=") {
      event.preventDefault();
      this.#setZoom(this.#status.getZoom() + 10);
    } else if (key === "-" || key === "_") {
      event.preventDefault();
      this.#setZoom(this.#status.getZoom() - 10);
    } else if (key === "0") {
      event.preventDefault();
      this.#setZoom(100);
    }
  };

  #cycleLandmarks(reverse: boolean): void {
    const landmarks = [
      this.shadowRoot?.querySelector("docen-ribbon") as HTMLElement | null,
      this.shadowRoot?.querySelector(".docen-canvas, docen-document-area") as HTMLElement | null,
      this.shadowRoot?.querySelector("docen-task-pane[open]") as HTMLElement | null,
      this.shadowRoot?.querySelector("docen-status-bar") as HTMLElement | null,
    ].filter((el): el is HTMLElement => Boolean(el));

    if (landmarks.length === 0) return;
    if (reverse) {
      this.#currentLandmarkIdx =
        (this.#currentLandmarkIdx - 1 + landmarks.length) % landmarks.length;
    } else {
      this.#currentLandmarkIdx = (this.#currentLandmarkIdx + 1) % landmarks.length;
    }
    const target = landmarks[this.#currentLandmarkIdx];
    if (target) {
      if (typeof target.focus === "function") target.focus();
      else if ("tabIndex" in target) {
        target.tabIndex = -1;
        target.focus();
      }
    }
  }

  #toggleRibbonMinimized(): void {
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon") as
      | (HTMLElement & {
          currentMode?: string;
          toggleMinimized?: () => void;
        })
      | null;
    if (!ribbon) return;
    if (typeof ribbon.toggleMinimized === "function") {
      ribbon.toggleMinimized();
    } else {
      const current = ribbon.getAttribute("data-ribbon-mode") ?? "always";
      const next = current === "tabs-only" ? "always" : "tabs-only";
      if (next === "always") {
        ribbon.removeAttribute("data-ribbon-mode");
      } else {
        ribbon.setAttribute("data-ribbon-mode", next);
      }
      ribbon.removeAttribute("data-expanded");
      ribbon.dispatchEvent(new CustomEvent("ribbon-mode-change", { detail: { mode: next } }));
    }
  }

  readonly #onCompareExecute = async (event: CustomEvent<any>): Promise<void> => {
    const detail = event.detail ?? {};
    let origJson = detail.originalJson;
    let revJson = detail.revisedJson;

    if (detail.originalFile) {
      try {
        const buf = await detail.originalFile.arrayBuffer();
        origJson = await parseDOCX(buf);
      } catch (err) {
        console.error("Failed to parse original document for comparison", err);
      }
    }
    if (detail.revisedFile) {
      try {
        const buf = await detail.revisedFile.arrayBuffer();
        revJson = await parseDOCX(buf);
      } catch (err) {
        console.error("Failed to parse revised document for comparison", err);
      }
    }

    if (!origJson && this.editor) {
      origJson = this.editor.getJSON();
    }

    if (origJson && revJson) {
      const author = detail.revisedAuthor || "Comparison";
      const result =
        detail.mode === "combine"
          ? combineDocs(origJson, [{ doc: revJson, author }])
          : compareDocs(origJson, revJson, { author });
      this.setJSON(result);
    }
  };

  readonly #onSignatureLineInsert = (event: CustomEvent<any>): void => {
    const { name, title, email } = event.detail ?? {};
    this.#bridge?.focus();
    this.editor?.commands.insertContent({
      type: "paragraph",
      attrs: { alignment: "left" },
      content: [
        { type: "text", text: "_____________________________________\n" },
        { type: "text", text: `X  ${name || ""}\n${title || ""}\n${email || ""}` },
      ],
    });
  };

  readonly #onDropCapApply = (event: CustomEvent<any>): void => {
    const { position, lines, distancePt } = event.detail ?? {};
    const editor = this.editor;
    if (!editor) return;
    this.#bridge?.focus();
    const parent = editor.state.selection.$from.parent;
    if (parent.type.name === "paragraph") {
      const pos = editor.state.selection.$from.before(1);
      const distanceTwip = Math.round((distancePt ?? 0) * 20);
      const dropCap =
        position === "none"
          ? null
          : {
              val: position,
              lines: lines ?? 3,
              distance: distanceTwip,
              vDistance: 0,
            };
      const frame =
        position === "none"
          ? null
          : {
              dropCap: position,
              lines: lines ?? 3,
              // office-open's frame writer reads w:hSpace/w:vSpace from `space`.
              space: { horizontal: distanceTwip, vertical: 0 },
            };
      const attrs = {
        ...parent.attrs,
        dropCap,
        frame,
      };
      editor.view.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, attrs));
    }
  };

  readonly #onRecipientsUpdated = (event: CustomEvent<any>): void => {
    const { recipients } = event.detail ?? {};
    if (recipients) {
      this.#merge.onRecipientsOk(new CustomEvent("recipients:ok", { detail: { recipients } }));
    }
  };

  readonly #onVersionRestore = (event: CustomEvent<any>): void => {
    const { doc } = event.detail ?? {};
    if (doc) {
      this.setJSON(doc);
    }
  };

  /** The protection domain's host view — thin adapters over this element's
   *  editor + documentExtras channel so the logic in ./protection is shared
   *  with the tests instead of living as private host methods. */
  #protectionView(): ProtectionHostView {
    return {
      editor: () => this.editor,
      settings: () => this.#documentSettings(),
      commitSettings: (settings) => {
        const editor = this.editor;
        if (!editor) return;
        const attrs = (editor.state.doc.attrs ?? {}) as {
          documentExtras?: Record<string, unknown>;
        };
        const extras = attrs.documentExtras ?? {};
        editor.view.dispatch(
          editor.state.tr.setDocAttribute("documentExtras", { ...extras, settings }),
        );
      },
      setMode: (mode, docProtected) => {
        this.#protectionMode = mode;
        this.#docProtected = docProtected;
        this.#syncEditable();
        this.#syncEditModeMenu();
      },
      pane: () =>
        this.shadowRoot?.querySelector(
          "docen-restrict-editing-pane",
        ) as unknown as ProtectionPane | null,
    };
  }

  readonly #onProtectionEnforce = (event: CustomEvent<any>): void => {
    enforceProtection(this.#protectionView(), event.detail ?? {});
  };

  readonly #onProtectionStop = (): void => {
    stopProtection(this.#protectionView());
  };

  readonly #onProtectionToggleException = (event: CustomEvent<any>): void => {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return;
    const { group } = event.detail ?? {};
    addPermissionRange(editor, { editGroup: group ?? "everyone" });
    this.#syncEditable();
  };

  readonly #onProtectionFindNext = (): void => {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return;
    findNextPermissionRange(editor);
  };

  readonly #onA11ySelectIssue = (event: CustomEvent<any>): void => {
    const issue = event.detail?.issue;
    if (!issue) return;
    this.#a11yMirror.announce(`Selected issue: ${issue.message}`);
  };

  readonly #onSdtDialogOk = (event: CustomEvent<any>): void => {
    const props = event.detail?.properties;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!props || !editor) return;
    const { $from } = editor.state.selection;
    for (let d = $from.depth; d > 0; d--) {
      const node = $from.node(d);
      if (
        node.type.name === "sdtBlock" ||
        node.type.name === "sdtInline" ||
        node.attrs?.properties
      ) {
        const pos = $from.before(d);
        const existing = (node.attrs.properties ?? {}) as Record<string, unknown>;
        const updated = {
          ...existing,
          title: props.title,
          tag: props.tag,
          cannotDelete: props.cannotDelete,
          cannotEdit: props.cannotEdit,
        };
        const tr = editor.state.tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          properties: updated,
        });
        editor.view.dispatch(tr);
        this.#bridge?.replaceOverlays();
        break;
      }
    }
  };

  readonly #onNavPagesSelect = (event: CustomEvent<{ page: number }>): void => {
    const pageNum = event.detail?.page;
    if (typeof pageNum !== "number" || this.#pages.length === 0) return;
    const destPage = Math.max(0, Math.min(this.#pages.length - 1, pageNum - 1));
    const pos = this.#bridge?.firstPosOfPage(destPage);
    if (pos != null) {
      this.#setTextSelection(pos);
      this.#bridge?.scrollIntoView(pos);
    }
  };

  /** Set a text selection (or a range) on the viewless editor. Same runtime
   *  PM instance — the cast bridges the dual d.ts identity between this
   *  package's @tiptap/pm and the engine's. */
  #setTextSelection(from: number, to?: number): void {
    const editor = this.editor;
    if (!editor) return;
    const sel = TextSelection.create(editor.state.doc, from, to);
    editor.view.dispatch(editor.state.tr.setSelection(sel as never));
    this.#bridge?.focus();
  }

  /** Editing → Select menu. "all" uses the official selectAll() command;
   *  "similar" selects similar formatting. */
  #select(value?: string): void {
    // The story the caret lives in — Ctrl+A selects the story text, the menu
    // command must agree with it (a stale main-doc range would be invisible).
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return;
    if (value === "similar") {
      selectSimilarFormatting(editor);
      return;
    }
    if ((value ?? "all") !== "all") return;
    this.#bridge?.focus();
    editor.commands.selectAll();
  }

  /** Editing → Find drop-down → Go To: open the Go To dialog (Page tab). */
  #goToPage(): void {
    this.#openGoToDialog("page");
  }

  /** Format Painter: a click captures the marks + paragraph formatting at the
   *  caret/selection and arms a one-shot pointerup; the next selection (or a
   *  tapped paragraph — Word paints the whole paragraph under a bare click)
   *  receives both and disarms. A double click arms sticky mode — every
   *  following selection paints until Esc or another painter click (Word's
   *  format painter). A click while armed cancels. */
  #toggleFormatPainter(): void {
    const now = performance.now();
    const rapid = now - this.#painterClickAt < PAINTER_DOUBLE_CLICK_MS;
    this.#painterClickAt = now;
    if (this.#painterMarks) {
      if (rapid) {
        // Second click of a double click: stay armed, paint repeatedly.
        this.#painterSticky = true;
        return;
      }
      this.#stopFormatPainter();
      return;
    }
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return;
    // One canvas tool at a time — arming this brush puts the others down.
    this.#stopShapeDrawing();
    this.#stopBorderPainting();
    // Probe one character in: $from sits on the boundary, and
    // ResolvedPos.marks() reads the character BEFORE the position — the first
    // selected (or caret-following) character's marks (e.g. bold stamped on
    // [from,to)) would be lost. Word captures at the bare caret too.
    this.#painterMarks = editor.state.doc.resolve(editor.state.selection.from + 1).marks();
    const $from = editor.state.selection.$from;
    if ($from.parent.type.name === "paragraph") {
      // Paragraph formatting paints too (Word: alignment, indent, spacing,
      // lists). The section-close markers belong to the target's position,
      // not the source's — stripped here, preserved on apply.
      const para = { ...($from.parent.attrs as Record<string, unknown>) };
      delete para.sectionProperties;
      delete para.sectionHeaders;
      delete para.sectionFooters;
      this.#painterPara = para;
    }
    this.#painterSticky = rapid;
    this.toggleAttribute("format-painter", true);
    // Arming runs outside any transaction (the lit state rides the attribute,
    // and #syncFormatButtons only fires per transaction).
    this.#syncFormatButtons();
    // Word paints on a document press, not a chrome press: a pointerup over
    // the ribbon, title bar, or panes — this painter's own button included —
    // leaves the stroke armed, so the button's click stays the sole toggle
    // (its pointerup would otherwise consume the stroke before the click can
    // toggle, and a second, unhurried click could never cancel).
    const paintStroke = (ed: Editor): void => {
      this.#applyFormatPainter(ed);
      // Sticky stays armed for the next selection/paragraph click.
      if (!this.#painterSticky) this.#stopFormatPainter();
    };
    const onUp = (event: PointerEvent): void => {
      if (!event.composedPath().some((n) => (n as HTMLElement).localName === "docen-document-area"))
        return;
      const ed = this.#bridge?.activeEditor() ?? this.editor;
      if (!ed) return;
      // A stroke over a selection paints it at once. A bare click defers
      // past the double-click window: the first press of a word-select
      // double click still reads the bare caret, and painting that whole
      // paragraph would bury the word the second press selects. Any new
      // press (the double click's own) drops the pending stroke; a quiet
      // window applies Word's whole-paragraph paint.
      clearTimeout(this.#painterStrokeTimer);
      if (!ed.state.selection.empty) {
        paintStroke(ed);
        return;
      }
      this.#painterStrokeTimer = setTimeout(() => {
        this.#painterStrokeTimer = undefined;
        if (!this.#painterMarks && !this.#painterPara) return;
        const cur = this.#bridge?.activeEditor() ?? this.editor;
        if (cur) paintStroke(cur);
      }, PAINTER_DOUBLE_CLICK_MS);
    };
    const onDown = (): void => {
      clearTimeout(this.#painterStrokeTimer);
      this.#painterStrokeTimer = undefined;
    };
    const onKey = (event: Event): void => {
      if ((event as KeyboardEvent).key === "Escape") this.#stopFormatPainter();
    };
    this.addEventListener("pointerdown", onDown);
    this.addEventListener("pointerup", onUp);
    this.addEventListener("keydown", onKey);
    this.#painterOff = () => {
      this.removeEventListener("pointerdown", onDown);
      this.removeEventListener("pointerup", onUp);
    };
    this.#painterKeyOff = () => this.removeEventListener("keydown", onKey);
  }

  /** Stamp the captured marks + paragraph attrs onto the current selection —
   *  a bare click (no drag) paints the whole paragraph under the caret. */
  #applyFormatPainter(ed: Editor): void {
    const { from, to, empty } = ed.state.selection;
    if (!this.#painterMarks && !this.#painterPara) return;
    let markFrom = from;
    let markTo = to;
    if (empty) {
      const $pos = ed.state.doc.resolve(from);
      if ($pos.parent.type.name !== "paragraph") return;
      markFrom = $pos.start();
      markTo = $pos.end();
    }
    const tr = ed.state.tr;
    for (const mark of this.#painterMarks ?? []) tr.addMark(markFrom, markTo, mark);
    if (this.#painterPara) {
      ed.state.doc.nodesBetween(markFrom, markTo, (node, pos) => {
        if (node.type.name !== "paragraph") return;
        const next: Record<string, unknown> = { ...this.#painterPara! };
        // The target keeps its own section-close markers.
        const target = node.attrs as Record<string, unknown>;
        for (const key of ["sectionProperties", "sectionHeaders", "sectionFooters"]) {
          if (target[key] != null) next[key] = target[key];
        }
        tr.setNodeMarkup(pos, undefined, next);
      });
    }
    ed.view.dispatch(tr);
  }

  #stopFormatPainter(): void {
    clearTimeout(this.#painterStrokeTimer);
    this.#painterStrokeTimer = undefined;
    this.#painterMarks = null;
    this.#painterPara = null;
    this.#painterSticky = false;
    this.removeAttribute("format-painter");
    // The disarm may land outside a transaction (Esc, a bare click) — the
    // apply path's transaction fires the sync itself, this covers the rest.
    this.#syncFormatButtons();
    this.#painterOff?.();
    this.#painterOff = undefined;
    this.#painterKeyOff?.();
    this.#painterKeyOff = undefined;
  }

  /** Arm the border painter (Table Design → Draw Border): canvas presses
   *  sweep table edges until Esc or a re-click disarms (the keydown follows
   *  the format painter's arm-while-armed lifecycle). */
  #armBorderPainter(erase: boolean): void {
    // One canvas tool at a time — arming this brush puts the others down.
    this.#stopShapeDrawing();
    this.#stopFormatPainter();
    this.#borderPainting = true;
    this.#borderErase = erase;
    const onKey = (event: Event): void => {
      if ((event as KeyboardEvent).key === "Escape") this.#stopBorderPainting();
    };
    this.addEventListener("keydown", onKey);
    this.#borderPaintKeyOff = () => this.removeEventListener("keydown", onKey);
  }

  #stopBorderPainting(): void {
    this.#borderPainting = false;
    this.#borderErase = false;
    this.#borderPaintKeyOff?.();
    this.#borderPaintKeyOff = undefined;
  }

  #armTableDrawer(): void {
    this.#stopBorderPainting();
    this.#stopFormatPainter();
    this.#stopShapeDrawing();
    this.#stopTableEraser();
    this.#tableDrawing = true;
    const onKey = (event: Event): void => {
      if ((event as KeyboardEvent).key === "Escape") {
        event.stopPropagation();
        this.#stopTableDrawing();
      }
    };
    this.addEventListener("keydown", onKey, true);
    this.#tableDrawKeyOff = () => this.removeEventListener("keydown", onKey, true);
  }

  #stopTableDrawing(): void {
    this.#tableDrawing = false;
    this.#tableDrawKeyOff?.();
    this.#tableDrawKeyOff = undefined;
  }

  #armTableEraser(): void {
    this.#stopBorderPainting();
    this.#stopFormatPainter();
    this.#stopShapeDrawing();
    this.#stopTableDrawing();
    this.#tableEraser = true;
    const onKey = (event: Event): void => {
      if ((event as KeyboardEvent).key === "Escape") {
        event.stopPropagation();
        this.#stopTableEraser();
      }
    };
    this.addEventListener("keydown", onKey, true);
    this.#tableEraserKeyOff = () => this.removeEventListener("keydown", onKey, true);
  }

  #stopTableEraser(): void {
    this.#tableEraser = false;
    this.#tableEraserKeyOff?.();
    this.#tableEraserKeyOff = undefined;
  }

  /** Arm the Shapes drawer (Insert → Shapes pick): the next canvas press
   *  draws and disarms. The keydown captures — before a draw an Escape only
   *  means "put the pencil down", not the bridge's selection Escapes
   *  underneath. */
  #armShapeDrawer(preset: string): void {
    // One canvas tool at a time — arming the drawer puts the brushes down.
    this.#stopFormatPainter();
    this.#stopBorderPainting();
    this.#armedShape = preset;
    const onKey = (event: Event): void => {
      if ((event as KeyboardEvent).key === "Escape") {
        event.stopPropagation();
        this.#stopShapeDrawing();
      }
    };
    this.addEventListener("keydown", onKey, true);
    this.#armedShapeKeyOff = () => this.removeEventListener("keydown", onKey, true);
  }

  #stopShapeDrawing(): void {
    this.#armedShape = null;
    this.#armedShapeKeyOff?.();
    this.#armedShapeKeyOff = undefined;
  }

  /** Mirror the font name / size and paragraph style at the caret into the
   *  ribbon comboboxes — Word behavior: the boxes report the formatting at the
   *  cursor, not a fixed default. Transactions only mark the chrome dirty; the
   *  actual pass is coalesced to one per animation frame (caret moves, marks
   *  change, doc edits all read the latest state in the same pass). */
  #setupFontSync(): void {
    const editor = this.editor;
    if (!editor) return;
    const sync = (): void => this.#scheduleUiSync();
    editor.on("transaction", sync);
    this.#syncUiChrome();
    this.#fontSyncCleanup = (): void => {
      editor.off("transaction", sync);
      if (this.#uiSyncFrame !== 0) {
        cancelAnimationFrame(this.#uiSyncFrame);
        this.#uiSyncFrame = 0;
      }
    };
  }

  /** Mark the host chrome dirty; the work runs once on the next frame. A
   *  typing burst fires a transaction per keystroke, but the comboboxes,
   *  greying, and status bar only need the state the frame renders. */
  #scheduleUiSync(): void {
    if (this.#uiSyncFrame !== 0) return;
    this.#uiSyncFrame = requestAnimationFrame(() => {
      this.#uiSyncFrame = 0;
      this.#syncUiChrome();
    });
  }

  /** The host's per-transaction UI pass: caret formatting in the ribbon,
   *  context tabs / story menus, button greying, and the status bar. */
  #syncUiChrome(): void {
    this.#syncFontControls();
    this.#syncStyleControl();
    this.#syncStoryMenus();
    this.#syncContextTabs();
    // After #syncContextTabs: the first transaction that enters a table is
    // also the one that appends the Table Layout panel — the combos only
    // exist from that pass on. The drawing Size combos ride the same pass.
    this.#syncCellSize();
    this.#syncTableLookCheckboxes();
    this.#syncDrawingSize();
    // Selection-sensitive greying: the arrange group's liveness depends on
    // what the selection points at, which no static pass sees.
    this.#syncArrangeGreying();
    this.#syncFormatButtons();
    this.#syncDrawingMenus();
    this.#syncQuickPartsMenu();
    this.#syncMiniToolbar();
    // The ruler's direction follows the caret's paragraph (Word mirrors the
    // scale for RTL runs even in an LTR shell).
    this.#syncRulerDirection();
    if (this.#uiSelectionDirty) {
      this.#uiSelectionDirty = false;
      // The status-bar language mirrors the caret's proofing language (Word).
      this.#syncStatusLanguage();
    }
    this.#updateStatus();
  }

  #syncFontControls(): void {
    const editor = this.editor;
    if (!editor) return;
    // Resolve font + size in one pass through the style inheritance chain
    // (direct run props → paragraph style → basedOn → document defaults).
    const { font, size } = effectiveRunProps(
      this.#docStyles(editor),
      this.#currentStyleId(editor),
      editor.getAttributes("textStyle"),
    );
    const fontDisplay = font ?? "";
    const sizeDisplay = size != null ? String(size) : "";
    const fontCb = this.shadowRoot?.querySelector<HTMLElement>(
      'docen-ribbon-combobox[event="font-name"]',
    );
    const sizeCb = this.shadowRoot?.querySelector<HTMLElement>(
      'docen-ribbon-combobox[event="font-size"]',
    );
    if (fontCb && fontCb.getAttribute("value") !== fontDisplay) {
      fontCb.setAttribute("value", fontDisplay);
    }
    if (sizeCb && sizeCb.getAttribute("value") !== sizeDisplay) {
      sizeCb.setAttribute("value", sizeDisplay);
    }
  }

  #syncMiniToolbar(): void {
    const miniToolbar = this.shadowRoot?.querySelector(
      "docen-mini-toolbar",
    ) as DocenMiniToolbar | null;
    if (!miniToolbar) return;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) {
      miniToolbar.hide();
      this.#lastMiniToolbarSelection = null;
      return;
    }
    const { selection } = editor.state;
    // Read-only / protected regions have no formatting to apply — Word does
    // not float the toolbar over them.
    const editable =
      editor.isEditable &&
      (this.#protectionMode !== "forms" || isInsideEditableField(editor)) &&
      (this.#protectionMode !== "readOnly" || isInsideEditablePermission(editor));
    if (
      !editable ||
      selection.empty ||
      !(selection instanceof TextSelection) ||
      this.#drawingStateOf() != null
    ) {
      miniToolbar.hide();
      this.#lastMiniToolbarSelection = null;
      return;
    }
    const isNewSel =
      !this.#lastMiniToolbarSelection ||
      this.#lastMiniToolbarSelection.from !== selection.from ||
      this.#lastMiniToolbarSelection.to !== selection.to;
    this.#lastMiniToolbarSelection = { from: selection.from, to: selection.to };

    if (isNewSel || miniToolbar.isOpen) {
      const rect = this.#bridge?.selectionClientRect(selection.from, selection.to);
      if (rect && (rect.width > 0 || rect.height > 0)) {
        miniToolbar.showNear(rect);
      }
    }
    const { font, size } = effectiveRunProps(
      this.#docStyles(editor),
      this.#currentStyleId(editor),
      editor.getAttributes("textStyle"),
    );
    miniToolbar.updateFormatting({
      bold: editor.isActive("bold"),
      italic: editor.isActive("italic"),
      underline: editor.isActive("underline"),
      fontName: font ?? undefined,
      fontSize: size != null ? String(size) : undefined,
      fontColor: (editor.getAttributes("textStyle")?.color as string) ?? undefined,
      highlightColor: (editor.getAttributes("highlight")?.color as string) ?? undefined,
      editable,
    });
  }

  #docStyles(editor: Editor): StylesOptions | null {
    return this.#styles.docStyles(editor);
  }

  #currentStyleId(editor: Editor): string | null {
    return this.#styles.currentStyleId(editor);
  }

  #syncStyleControl(): void {
    this.#styles.syncStyleControl();
  }

  // ── Styles pane / Modify Style dialog / Style Inspector ──────────────────

  #renderStylesPane(): void {
    this.#styles.renderStylesPane();
  }

  #renderStylesInspector(): void {
    this.#styles.renderStylesInspector();
  }

  #openModifyStyle(id: string): void {
    this.#styles.openModifyStyle(id);
  }

  #openNewStyle(): void {
    this.#styles.openNewStyle();
  }

  #openStyleFormat(id: string, target: "font" | "paragraph"): void {
    this.#styles.openStyleFormat(id, target);
  }

  /** The Paragraph dialog's OK — a style-target open (Format > Paragraph)
   *  restamps the style's definition; the regular open stamps the selection
   *  (DialogCommands' path). */
  readonly #onParagraphDialogOk = (event: Event): void => {
    const patch = (event as CustomEvent<ParagraphDialogPatch>).detail;
    const dialog = this.shadowRoot?.querySelector("docen-paragraph-dialog");
    const styleId = dialog?.getAttribute("data-for-style") ?? null;
    if (styleId && patch) {
      dialog?.removeAttribute("data-for-style");
      this.editor?.commands["style-paragraph-patch"]({ id: styleId, patch });
      this.#renderStylesPane();
      this.#bridge?.focus();
      return;
    }
    this.#dialogs.onParagraphOk(event as CustomEvent<ParagraphDialogPatch>);
  };

  /** The Font dialog's OK — the style-target twin of #onParagraphDialogOk. */
  readonly #onFontDialogOk = (event: Event): void => {
    const patch = (event as CustomEvent<FontDialogPatch>).detail;
    const dialog = this.shadowRoot?.querySelector("docen-font-dialog");
    const styleId = dialog?.getAttribute("data-for-style") ?? null;
    if (styleId && patch) {
      dialog?.removeAttribute("data-for-style");
      this.editor?.commands["style-run-patch"]({ id: styleId, props: fontRunPropsOf(patch) });
      this.#renderStylesPane();
      this.#bridge?.focus();
      return;
    }
    this.#dialogs.onFontOk(event as CustomEvent<FontDialogPatch>);
  };

  /** A ribbon gallery's right-click — only the Styles gallery routes it
   *  today: Modify the right-clicked style (Word's gallery context entry). */
  readonly #onItemContext = (event: Event): void => {
    const detail = (event as CustomEvent<{ event?: string; value?: string }>).detail;
    if (detail?.event === "style" && detail.value) this.#openModifyStyle(detail.value);
  };

  #snapshotStyles(): void {
    this.#styles.snapshotStyles();
  }

  #restoreStylesSnapshot(): void {
    this.#styles.restoreStylesSnapshot();
  }

  async connectedCallback(): Promise<void> {
    super.connectedCallback();
    // Forward this host's `lang` attribute to the internal <docen-workspace>
    // so resolveLang (scoped to docen-workspace) honors <docen-document lang>,
    // not just <html lang>. See #syncLang for why MutationObserver, not @attr.
    this.#langObserver = new MutationObserver(() => this.#syncLang());
    this.#langObserver.observe(this, { attributes: true, attributeFilter: ["lang"] });
    this.#syncLang();
    await registerComponents();
    applyTheme(resolveTheme(this.getAttribute("theme")));

    this.#fileInput = this.shadowRoot!.querySelector<HTMLInputElement>("#file-input")!;
    this.#imageInput = this.shadowRoot!.querySelector<HTMLInputElement>("#image-input")!;
    this.#pictureInput = this.shadowRoot!.querySelector<HTMLInputElement>("#picture-input")!;
    this.#renderChrome();
    this.#loadPersistedHistory();
    // Once attributes: initial task-pane visibility (Office `setStartupBehavior`
    // equivalent). Absent → closed; present → open. Read once on connect —
    // runtime toggles go through showTaskpane/hideTaskpane.
    this.#setTaskpane("navigation", this.hasAttribute("navigation-pane"));
    this.#setTaskpane("properties", this.hasAttribute("properties-pane"));
    // Once attribute: initial zoom level (percent). Runtime zoom goes through
    // setZoom / the status-bar slider / Ctrl+ -/=/0.
    const initialZoom = this.getAttribute("zoom");
    if (initialZoom) this.#setZoom(Number(initialZoom) || 100);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-status-bar")
      ?.addEventListener("zoom:change", this.#status.onZoomChange as EventListener);
    // The proofing surfaces: the status-bar book opens the pane; the pane's
    // actions (replace/ignore/add/step) come back as events.
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-status-bar")
      ?.addEventListener("spellcheck:open", () =>
        this.#onCommand(new CustomEvent("command", { detail: { event: "spell-check" } })),
      );
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:replace", ((event: CustomEvent<string>) =>
        this.#spelling.replace(event.detail)) as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:ignore-once", () => this.#spelling.ignore("once"));
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:ignore-all", () => this.#spelling.ignoreAll());
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:add", () => this.#spelling.ignore("add"));
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-spelling-pane")
      ?.addEventListener("spelling:nav", ((event: CustomEvent<number>) =>
        this.#spelling.goto(this.#spelling.activeIndex() + event.detail)) as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-thesaurus-pane")
      ?.addEventListener("thesaurus:insert", ((event: CustomEvent<string>) => {
        if (event.detail && this.editor) {
          this.editor.commands.insertContent(event.detail);
        }
      }) as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-translate-pane")
      ?.addEventListener("translate:insert", ((event: CustomEvent<string>) => {
        if (event.detail && this.editor) {
          this.editor.commands.insertContent(event.detail);
        }
      }) as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-translate-pane")
      ?.addEventListener("translate:document", ((
        event: CustomEvent<{ from?: string; to?: string }>,
      ) => {
        this.#translateDocument(event.detail?.from, event.detail?.to);
      }) as EventListener);
    this.addEventListener("docen:translate", ((event: CustomEvent<{ text?: string }>) => {
      this.#openTranslate(event.detail?.text);
    }) as EventListener);

    this.#stageHost = this.shadowRoot!.querySelector<HTMLElement>(".docen-canvas") ?? undefined;
    this.#mountRuler();
    this.shadowRoot
      ?.querySelector<HTMLElement>(".content-warning-close")
      ?.addEventListener("click", () => this.#dismissContentWarning());
    this.#stageHost?.addEventListener("wheel", this.#onWheel as EventListener, {
      capture: true,
      passive: false,
    });
    if (!this.#stageHost) return;

    // Fonts must be loaded before the pipeline measures, else the layout
    // drifts from the browser's actual font metrics.
    await document.fonts?.ready;
    // The production shaping set: register the bundled metric-compatible
    // faces before the first layout so the default Word families shape
    // deterministically instead of silently falling back to canvas. A failure
    // (e.g. a bundler that did not emit the package's asset URLs) is loud but
    // non-fatal — the canvas measurer keeps working.
    try {
      await this.registerDefaultFonts();
    } catch (err) {
      console.warn(
        "[docen-document] bundled shaping fonts unavailable — using canvas metrics. " +
          "Call registerDefaultFonts({ baseUrl }) with the emitted assets/fonts directory to opt in.",
        err,
      );
    }

    const contentAttr = this.getAttribute("content");
    // Declarative section-properties / styles (JSON) seed doc-level attrs so a
    // host can bootstrap page setup + named styles without openDOCX/setJSON.
    const initAttrs = this.#readInitAttrs();
    // The content attribute accepts Tiptap JSON only; a malformed string
    // mounts an empty document rather than throwing mid-connection.
    let baseDoc = {} as JSONContent;
    if (contentAttr) {
      try {
        baseDoc = JSON.parse(contentAttr) as JSONContent;
      } catch {
        console.warn("[docen-document] content attribute is not valid JSON — ignored");
      }
    }
    const seeded =
      Object.keys(initAttrs).length > 0
        ? { ...baseDoc, attrs: { ...baseDoc.attrs, ...initAttrs } }
        : baseDoc;
    // Fill office-open's document-level defaults (the built-in style library +
    // page geometry + docGrid linePitch) so a freshly mounted document matches
    // what export produces. Host-declared initAttrs win — normalizeDocument
    // shallow-merges user attrs over the defaults.
    const initialDoc =
      (seeded.attrs as { sectionProperties?: unknown } | undefined)?.sectionProperties == null
        ? normalizeDocument(seeded)
        : seeded;
    // The default document add-in contributes the engine extensions + every
    // wired ribbon command. Registered before the editor mounts so its
    // extensions seed the schema. Ribbon events route straight to the engine
    // (editor.commands.<event>), not addin.commands.
    const defaultAddin = createDefaultAddin({
      onOutlineUpdate: (anchors) => this.#navigation.renderOutline(anchors),
    });
    this.addAddin(defaultAddin);
    // Declarative external add-ins (JSON `addins` attribute) register after the
    // default so their ribbon tabs append to the built-ins via mergeRibbonSchema.
    this.#applyAddinsAttr();

    this.#bridge = mountEditBridge({
      host: this.#stageHost,
      canEdit: (ed) => {
        if (this.#protectionMode === "forms") return isInsideEditableField(ed);
        if (this.#protectionMode === "readOnly") return isInsideEditablePermission(ed);
        return true;
      },
      // The textarea must live outside docen-context-menu (fluent-menu eats
      // Space/Enter) — the input layer at the shadow root is menu-free.
      inputHost: this.shadowRoot!.querySelector<HTMLElement>(".input-layer")!,
      content: initialDoc,
      onDoc: (json) => this.#renderDoc(json),
      markdown: () => this.#markdown,
      // The AutoCorrect rules + user table are read per keystroke, so edits in
      // the AutoCorrect Options dialog land mid-session.
      autocorrect: () => autocorrectConfigOf(getSettings().writing.autocorrect),
      pageHost: (page) => this.#stage?.slotAt(page)?.parentElement ?? null,
      extensions: [...docxExtensions, ...(defaultAddin.extensions ?? [])],
      scale: () => this.#stage?.scale() ?? 1,
      onZoomChange: (scale) => this.#setZoom(Math.round(scale * 100)),
      setScale: (scale) => this.#setZoom(Math.round(scale * 100)),
      onPointerTypeChange: (type) => {
        if (type === "touch" && !this.touchMode) {
          this.touchMode = true;
        }
      },
      contentWidthPx: () => this.#flow?.contentWidthPx,
      frontFloats: (page) => this.#stage?.frontFloatBoxes(page) ?? [],
      // Word's paste-options bar hangs after every rich paste; the clipboard
      // pane collects each in-editor copy/cut.
      onRichPaste: (source) => this.#clipboard.showPasteOptions(source),
      onClipboardCollect: (item) => this.#clipboard.collect(item),
      // Header/footer edit stories — the bridge routes band double-clicks
      // here; the host owns the slots' persistence (attrs on the doc node or
      // the section's last paragraph).
      story: {
        geometry: (kind, page) => {
          const stage = this.#stage;
          const band = stage?.furnitureBand(kind, page);
          if (!stage || !band) return null;
          return {
            stack: stage.furnitureStack(kind, page),
            band,
            slot: stage.slotOfPage(page),
          };
        },
        read: (kind, slot, page) => this.#readStorySource(kind, slot, page),
        entered: (kind, slot, page) => this.#stories.enter(kind, slot, page),
        onDoc: (kind, slot, json) => this.#renderStoryFurniture(kind, slot, json),
        exit: ({ kind, slot, json, dirty }) => this.#exitStory(kind, slot, json, dirty),
      },
      // Drawing selection — the stage's painted-box hit table resolves the
      // click; the caret map pairs the hit's host paragraph to the PM
      // position, and the index-th drawing node inside it becomes the
      // NodeSelection (projectDrawings collects drawings in run order, the
      // same order the paragraph's content carries the nodes).
      drawingAt: (page, lx, ly) => this.#stage?.drawingAt(page, lx, ly) ?? null,
      // Margin balloons: the stage's painted card table routes clicks to the
      // comment/revision commands and the hover tone back to the stage.
      balloonAt: (page, lx, ly) => this.#stage?.balloonAt(page, lx, ly) ?? null,
      onBalloonSelect: (hit) => this.#revisions.onBalloonSelect(hit),
      onBalloonHover: (hit) => this.#stage?.hoverBalloon(hit),
      drawingSelection: (hit, enter) =>
        this.#drawingNodePos(hit.para, hit.index, hit.kind, hit.childPath, enter),
      pageFlow: (page) => this.#stage?.flowOf(page) ?? null,
      siblingBoxes: (page, hit) =>
        this.#stage?.pageDrawingBoxes(
          page,
          hit
            ? ({ para: hit.para, index: hit.index, childPath: hit.childPath } as never)
            : undefined,
        ) ?? [],
      chartPartBoxes: (para, index, kind) => this.#stage?.chartPartBoxesOf(para, index, kind) ?? [],
      shapeTextStacks: () => this.#stage?.allShapeTextStacks() ?? [],
      shapeResolve: (host) => {
        const pos = this.#drawingNodePos(host.para, host.index, "drawing");
        const doc = this.#bridge?.editor.state.doc;
        // A group-interior stack pairs with the member wpsShape the childPath
        // addresses — no selection gate (registration is a render-time pass).
        const target =
          pos != null && doc && host.childPath?.length
            ? descendGroupChild(doc, pos, host.childPath)
            : pos;
        const node = target != null && doc ? doc.nodeAt(target) : null;
        return target != null && node ? { pos: target, node } : null;
      },
      drawingBoxOf: (para, index, kind, childPath) => {
        const stage = this.#stage;
        const hit = stage?.drawingBoxOf(para, index, kind, childPath) ?? null;
        if (hit || !stage) return hit;
        // A write-back (a resize) re-laid the doc and re-objected every
        // paragraph — the caller's hit reference is stale. Re-resolve the
        // PM selection (still a NodeSelection on the same drawing) against
        // the fresh boxes so the frame follows the resized picture.
        const sel = this.editor?.state.selection;
        if (!(sel instanceof NodeSelection)) return null;
        return (
          stage
            .drawingBoxes()
            .find(
              (box) =>
                !box.chartPart &&
                this.#drawingNodePos(box.para, box.index, box.kind, box.childPath) === sel.from,
            ) ?? null
        );
      },
      // `#name` links are bookmark anchors — the host owns the in-page jump.
      onInternalAnchor: (name) => this.#jumpToBookmark(name),
      // The border painter's armed state and its commit (Table Design →
      // Draw Border): a sweep's crossed edges ride one paint/erase command,
      // the pen merged from the host's pen state.
      borderPaint: () => ({ active: this.#borderPainting, eraser: this.#borderErase }),
      formatPaint: () => this.hasAttribute("format-painter"),
      // The Shapes drawer (Insert → Shapes): the armed preset rides the ghost
      // drag; the commit lands the shape run at the rect (page-local px).
      shapeDraw: () => this.#armedShape,
      applyShapeDraw: (rect) => {
        const preset = this.#armedShape ?? "rect";
        // Word's default: one draw returns the pointer to the selection tool
        // (consecutive draws need another Shapes pick, or Esc never fires).
        this.#stopShapeDrawing();
        this.#insertShapeAt(preset, rect);
      },
      applyBorderPaint: (sides) => {
        const editor = this.#bridge?.activeEditor() ?? this.editor;
        if (!editor || !this.#borderPainting || !sides.length) return;
        if (this.#borderErase) {
          editor.commands["erase-cell-border"](JSON.stringify({ sides }));
        } else {
          editor.commands["paint-cell-border"](JSON.stringify({ sides, pen: this.#pen }));
        }
      },
      tableDraw: () => this.#tableDrawing,
      tableEraser: () => this.#tableEraser,
      stopTableDraw: () => this.#stopTableDrawing(),
      stopTableEraser: () => this.#stopTableEraser(),
      applyTableDraw: (rect) => {
        const editor = this.#bridge?.activeEditor() ?? this.editor;
        if (!editor || !this.#tableDrawing) return;
        const inTable =
          this.#bridge?.cellAtPoint(rect.page, rect.stroke.x1, rect.stroke.y1) != null;
        (editor.commands as any)["draw-table-stroke"]?.({
          page: rect.page,
          widthPx: rect.width,
          heightPx: rect.height,
          dx: rect.stroke.x2 - rect.stroke.x1,
          dy: rect.stroke.y2 - rect.stroke.y1,
          inTable,
        });
      },
      applyTableEraser: (sides) => {
        const editor = this.#bridge?.activeEditor() ?? this.editor;
        if (!editor || !this.#tableEraser || !sides.length) return;
        (editor.commands as any)["table-eraser-click"]?.({ sides });
      },
    });
    if (this.getAttribute("editable") === "false") this.#bridge.editor.setEditable(false);
    // First paint + caret map feed (transactions re-render via the bridge's
    // raf-merged onDoc from here on).
    this.#snapshotStyles();
    this.#renderDoc(initialDoc);

    // Mirror the caret's font/size into the ribbon comboboxes (Word behavior).
    this.#setupFontSync();

    // command = ribbon buttons; change = menu items + auto-save switch. Listen
    // on the shadow root so non-composed Fluent events (menu-item "change")
    // reach us, not just composed ones (ribbon "command").
    this.shadowRoot!.addEventListener("command", this.#onCommand as EventListener);
    // Ribbon gallery right-click (Word's gallery context entry) — only the
    // Styles gallery routes it today: Modify the right-clicked style.
    this.shadowRoot!.addEventListener("item-context", this.#onItemContext as EventListener);
    this.shadowRoot!.addEventListener("item-preview", this.#onItemPreview as EventListener);
    this.shadowRoot!.addEventListener("item-preview-end", this.#onItemPreviewEnd as EventListener);
    this.shadowRoot!.addEventListener("change", this.#onChange as EventListener);
    // Right-click → Word's context menu. Captured on the shadow root so the
    // items are built before <docen-context-menu>'s own capture handler opens
    // the Fluent menu (capture runs outermost-first).
    this.shadowRoot!.addEventListener("contextmenu", this.#onContextMenu as EventListener, true);
    // QAT undo/redo history flyouts — click delegation on the shadow root (the
    // title bar is re-stamped per #renderChrome): a caret trigger fills its
    // list with the live depths, an entry rewinds/advances that many steps.
    this.shadowRoot!.addEventListener("click", this.#onHistoryClick as EventListener);
    this.#fileInput.addEventListener("change", this.#io.onFileChange);
    this.#imageInput.addEventListener("change", this.#io.onImageChange);
    this.#pictureInput.addEventListener("change", this.#io.onPictureChange);
    // Outline (Headings tab) → jump to the clicked heading.
    this.shadowRoot!.querySelector("docen-outline")?.addEventListener(
      "outline:select",
      this.#navigation.onOutlineSelect as EventListener,
    );
    // Nav-pane search → Find (live highlight, next/prev, results count).
    this.addEventListener("navigation:search", this.#navigation.onSearch as EventListener);
    this.addEventListener("navigation:find", this.#navigation.onFind as EventListener);
    // Comments pane → create/cancel (compose box), select (scroll to range),
    // update/delete (inline card actions), reply/resolve (thread actions).
    this.addEventListener("comment:create", this.#comments.onCommentCreate as EventListener);
    this.addEventListener("comment:cancel", this.#comments.onCommentCancel as EventListener);
    this.addEventListener("comment:select", this.#comments.onCommentSelect as EventListener);
    this.addEventListener("comment:update", this.#comments.onCommentUpdate as EventListener);
    this.addEventListener("comment:delete", this.#comments.onCommentDelete as EventListener);
    this.addEventListener("comment:reply", this.#comments.onCommentReply as EventListener);
    this.addEventListener("comment:resolve", this.#comments.onCommentResolve as EventListener);
    // Reviewing pane → select (scroll to the revision), accept/reject (the
    // by-id command form — one transaction per card action).
    this.addEventListener("revision:select", this.#revisions.onRevisionSelect as EventListener);
    this.addEventListener("revision:accept", this.#revisions.onRevisionAccept as EventListener);
    this.addEventListener("revision:reject", this.#revisions.onRevisionReject as EventListener);
    // Office Clipboard pane → paste one entry / paste all / clear.
    this.addEventListener("clipboard:paste", this.#clipboard.onPanePaste as EventListener);
    this.addEventListener("clipboard:paste-all", this.#clipboard.onPanePasteAll as EventListener);
    this.addEventListener("clipboard:clear", this.#clipboard.onPaneClear as EventListener);
    // Styles pane → apply a style / open Modify Style / switch to the
    // inspector view (which pushes fresh data back).
    this.shadowRoot!.querySelector("docen-styles-pane")?.addEventListener("style-apply", ((
      event: CustomEvent<string>,
    ) => {
      this.editor?.commands.style(event.detail);
      this.#bridge?.focus();
    }) as EventListener);
    this.shadowRoot!.querySelector("docen-styles-pane")?.addEventListener("new-style", (() =>
      this.#openNewStyle()) as EventListener);
    this.shadowRoot!.querySelector("docen-styles-pane")?.addEventListener("modify-style", ((
      event: CustomEvent<string>,
    ) => this.#openModifyStyle(event.detail)) as EventListener);
    this.shadowRoot!.querySelector("docen-styles-pane")?.addEventListener("view", ((event) => {
      if ((event as CustomEvent<string>).detail === "inspector") this.#renderStylesInspector();
      else this.#renderStylesPane();
    }) as EventListener);
    // Modify Style dialog — stamp the patched definition (the modify-style
    // command rides the undo history) and re-render the pane so the row
    // previews follow the new formatting.
    this.shadowRoot!.querySelector("docen-modify-style-dialog")?.addEventListener(
      "modify-style:ok",
      ((event: CustomEvent<ModifyStylePatch>) => {
        this.editor?.commands["modify-style"](event.detail);
        this.#renderStylesPane();
        this.#bridge?.focus();
      }) as EventListener,
    );
    // New Style dialog — create a new paragraph or character style.
    this.shadowRoot!.querySelector("docen-new-style-dialog")?.addEventListener("new-style:ok", ((
      event: CustomEvent<NewStyleDefinition>,
    ) => {
      this.editor?.commands["new-style"](event.detail);
      this.#renderStylesPane();
      this.#bridge?.focus();
    }) as EventListener);
    // Modify Style dialog's Format > Font/Paragraph — open that dialog
    // against the style being modified (the detail carries its id).
    this.shadowRoot!.querySelector("docen-modify-style-dialog")?.addEventListener(
      "modify-style:format",
      ((event: CustomEvent<{ id: string; target: "font" | "paragraph" }>) =>
        this.#openStyleFormat(event.detail.id, event.detail.target)) as EventListener,
    );
    // Click a Results entry → jump to that match (delegated on the container).
    this.shadowRoot!.querySelector(".search-results")?.addEventListener(
      "click",
      this.#navigation.onSearchResultClick as EventListener,
    );
    // Find & Replace dialog → Replace / Replace All (prosemirror-search).
    this.shadowRoot!.querySelector("docen-find-replace-dialog")?.addEventListener(
      "find-replace:action",
      this.#navigation.onFindReplace as EventListener,
    );
    // Options dialog — ok (UI language + theme).
    this.shadowRoot!.querySelector("docen-options-dialog")?.addEventListener(
      "options:ok",
      this.#onOptionsOk as EventListener,
    );
    // Options → Proofing → AutoCorrect Options: open the dedicated dialog
    // (over the open Options modal) and commit its OK to the settings store.
    this.shadowRoot!.querySelector("docen-options-dialog")?.addEventListener(
      "options:autocorrect",
      this.#openAutocorrectDialog as EventListener,
    );
    this.shadowRoot!.querySelector("docen-autocorrect-dialog")?.addEventListener(
      "autocorrect:ok",
      this.#onAutocorrectOk as EventListener,
    );
    // Quick Parts dialogs — the save commit + the organizer's per-row actions
    // route into the building-blocks domain.
    this.shadowRoot!.querySelector("docen-quick-part-dialog")?.addEventListener(
      "quick-part:save",
      this.#onQuickPartSave as EventListener,
    );
    this.shadowRoot!.querySelector("docen-building-blocks-dialog")?.addEventListener(
      "building-blocks:insert",
      this.#onBuildingBlockEvent as EventListener,
    );
    this.shadowRoot!.querySelector("docen-building-blocks-dialog")?.addEventListener(
      "building-blocks:rename",
      this.#onBuildingBlockEvent as EventListener,
    );
    this.shadowRoot!.querySelector("docen-building-blocks-dialog")?.addEventListener(
      "building-blocks:delete",
      this.#onBuildingBlockEvent as EventListener,
    );
    // Footnote/endnote settings dialog — ok (document-level numbering).
    this.shadowRoot!.querySelector("docen-note-settings-dialog")?.addEventListener(
      "note-settings:ok",
      this.onNoteSettingsOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-note-settings-dialog")?.addEventListener(
      "note-settings:convert",
      this.onNoteSettingsConvert as EventListener,
    );
    // Line Numbering Options dialog — ok (current section's w:lnNumType).
    this.shadowRoot!.querySelector("docen-line-numbers-dialog")?.addEventListener(
      "line-numbers:ok",
      this.#sections.onLineNumbersOk as EventListener,
    );
    // Page Number Format dialog — ok (current section's w:pgNumType).
    this.shadowRoot!.querySelector("docen-page-number-format-dialog")?.addEventListener(
      "page-number-format:ok",
      this.#sections.onPageNumberFormatOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-hyphenation-dialog")?.addEventListener(
      "hyphenation:ok",
      this.#insert.onHyphenationOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-tabs-dialog")?.addEventListener(
      "tabs:ok",
      this.#insert.onTabsOk as EventListener,
    );
    // Bookmark dialog — add, delete, goto.
    const bmEl = this.shadowRoot!.querySelector("docen-bookmark-dialog");
    bmEl?.addEventListener("bookmark:add", this.onBookmarkAdd as EventListener);
    bmEl?.addEventListener("bookmark:delete", this.onBookmarkDelete as EventListener);
    bmEl?.addEventListener("bookmark:goto", this.onBookmarkGoTo as EventListener);
    // Go To dialog — navigate.
    this.shadowRoot!.querySelector("docen-go-to-dialog")?.addEventListener(
      "goto:navigate",
      this.onGoToNavigate as EventListener,
    );
    // Document Properties dialog — ok.
    this.shadowRoot!.querySelector("docen-properties-dialog")?.addEventListener(
      "properties:ok",
      this.onPropertiesOk as EventListener,
    );
    // Document Inspector (检查问题) — the findings dialog's removal buttons.
    this.shadowRoot!.querySelector("docen-inspect-dialog")?.addEventListener(
      "inspect:clear-comments",
      this.#onInspect as EventListener,
    );
    this.shadowRoot!.querySelector("docen-inspect-dialog")?.addEventListener(
      "inspect:accept-revisions",
      this.#onInspect as EventListener,
    );
    // New from Template dialog — instantiate the picked built-in template as a
    // new document (model JSON through the normal load path; no canvas DOM).
    this.shadowRoot!.querySelector("docen-template-dialog")?.addEventListener("template:create", ((
      event: CustomEvent<{ id: string }>,
    ) => this.#newFromTemplate(event.detail.id)) as EventListener);
    this.shadowRoot!.querySelector("docen-template-dialog")?.addEventListener("template:attach", ((
      event: CustomEvent<{ data: Uint8Array; autoUpdateStyles: boolean; filename: string }>,
    ) =>
      this.attachTemplate(event.detail.data, {
        autoUpdateStyles: event.detail.autoUpdateStyles,
      })) as EventListener);
    this.shadowRoot!.querySelector("docen-print-preview")?.addEventListener(
      "print-preview:markup-change",
      ((event: Event) => {
        const customEvent = event as CustomEvent<{ markup: boolean }>;
        if (!this.#stage) return;
        void (async () => {
          const shots = await this.#stage!.printSnapshots({ markup: customEvent.detail.markup });
          const previewEl = this.shadowRoot?.querySelector("docen-print-preview") as
            | (HTMLElement & {
                seed(options: unknown): void;
                updatePreview(): void;
              })
            | null;
          const cursor = this.editor?.state.selection.from ?? 0;
          const page = (this.#bridge?.pageOf(cursor) ?? 0) + 1;
          previewEl?.seed({
            snapshots: shots,
            filename: this.getAttribute("filename") ?? "Document",
            currentPage: page,
            printMarkup: customEvent.detail.markup,
          });
          previewEl?.updatePreview();
        })();
      }) as EventListener,
    );
    // Language dialog — commit the selection's proofing language (w:lang).
    this.shadowRoot!.querySelector("docen-language-dialog")?.addEventListener(
      "language:ok",
      this.#dialogs.onLanguageOk as EventListener,
    );
    // Date and Time dialog — insert the picked format (static text or a DATE
    // field when "update automatically" is checked).
    this.shadowRoot!.querySelector("docen-date-time-dialog")?.addEventListener(
      "date-time:insert",
      ((event: CustomEvent<{ text: string; instruction?: string }>) =>
        this.#insertDateTime(event.detail)) as EventListener,
    );
    // Custom Table of Contents dialog — insert the TOC with the picked level
    // window / leader / page-number shape, then repaginate-and-update (the
    // same post-insert pass as the plain toc command).
    this.shadowRoot!.querySelector("docen-toc-dialog")?.addEventListener("toc:ok", ((
      event: CustomEvent<{
        headingRange: string;
        leader: string;
        showPageNumbers: boolean;
        alignPageNumbers: boolean;
        hyperlink?: boolean;
        styles?: string;
      }>,
    ) => this.#insertCustomToc(event.detail)) as EventListener);
    // Phonetic guide dialog — split the selection into per-character ruby
    // runs, or strip the guides off it.
    this.shadowRoot!.querySelector("docen-phonetic-dialog")?.addEventListener(
      "phonetic:ok",
      this.#dialogs.onPhoneticOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-phonetic-dialog")?.addEventListener(
      "phonetic:clear",
      this.#dialogs.onPhoneticClear as EventListener,
    );
    // Two Lines in One dialog — pack the selection's text into two half-size
    // lines (双行合一 / 合并字符).
    this.shadowRoot!.querySelector("docen-two-in-one-dialog")?.addEventListener(
      "two-in-one:ok",
      this.#dialogs.onTwoInOneOk as EventListener,
    );
    // Define New Multilevel List dialog — register the levels as a document
    // numbering definition and stamp the selection with it.
    this.shadowRoot!.querySelector("docen-define-list-dialog")?.addEventListener(
      "define-list:ok",
      this.#dialogs.onDefineListOk as EventListener,
    );
    // Caption dialog — seed a Caption-styled paragraph with a SEQ field next
    // to the caret's paragraph.
    this.shadowRoot!.querySelector("docen-caption-dialog")?.addEventListener(
      "caption:ok",
      this.#dialogs.onCaptionOk as EventListener,
    );
    // Note dialog — insert a footnote/endnote reference at the caret (or
    // rewrite the referenced note's body, when opened from the context menu).
    this.shadowRoot!.querySelector("docen-note-dialog")?.addEventListener(
      "note:ok",
      this.#dialogs.onNoteOk as EventListener,
    );
    // Field dialog — insert a field atom at the caret (or rewrite the
    // referenced field's instruction + cache, when opened from the context
    // menu).
    this.shadowRoot!.querySelector("docen-field-dialog")?.addEventListener(
      "field:ok",
      this.#dialogs.onFieldOk as EventListener,
    );
    // Chart data dialog — the Edit Data grid's commit (Chart Design tab).
    this.shadowRoot!.querySelector("docen-chart-data-dialog")?.addEventListener(
      "chart:ok",
      this.#dialogs.onChartOk as EventListener,
    );
    // Chart type dialog — the Change Chart Type picker commit (Chart Design tab).
    this.shadowRoot!.querySelector("docen-chart-type-dialog")?.addEventListener(
      "chart-type:ok",
      this.#dialogs.onChartTypeOk as EventListener,
    );
    // Compress Pictures — the OK path re-encodes the selected picture's
    // pixels (the picture-pixels command swaps the source in).
    this.shadowRoot!.querySelector("docen-compress-pictures-dialog")?.addEventListener(
      "compress:ok",
      this.#onCompressOk as EventListener,
    );
    // Cross-reference dialog — seed a cached REF/PAGEREF field at the caret.
    this.shadowRoot!.querySelector("docen-cross-reference-dialog")?.addEventListener(
      "cross-ref:ok",
      this.#dialogs.onCrossRefOk as EventListener,
    );
    // Sources dialog — write the bibliography source list (attrs.bibliography)
    // and seed a cached CITATION field at the caret.
    this.shadowRoot!.querySelector("docen-sources-dialog")?.addEventListener(
      "sources:ok",
      this.#references.onSourcesOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-sources-dialog")?.addEventListener(
      "citation:ok",
      this.#references.onCitationOk as EventListener,
    );
    // Recipients dialog — write the merge data source (attrs.recipients).
    this.shadowRoot!.querySelector("docen-recipients-dialog")?.addEventListener(
      "recipients:ok",
      this.#merge.onRecipientsOk as EventListener,
    );
    // Merge-field dialog — seed a MERGEFIELD simple field at the caret.
    this.shadowRoot!.querySelector("docen-merge-field-dialog")?.addEventListener(
      "merge-field:ok",
      this.#onMergeFieldOk as EventListener,
    );
    // Status-bar language item — open the language dialog (Word semantics).
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "language:open",
      this.#onLanguageOpen as EventListener,
    );
    // Symbol dialog — insert the picked character at the caret.
    this.shadowRoot!.querySelector("docen-symbol-dialog")?.addEventListener(
      "symbol:insert",
      this.#onSymbolInsert as EventListener,
    );
    this.shadowRoot!.appendChild(this.#a11yMirror.root);
    this.shadowRoot!.querySelector("docen-compare-dialog")?.addEventListener(
      "compare:execute",
      this.#onCompareExecute as unknown as EventListener,
    );
    this.shadowRoot!.querySelector("docen-signature-line-dialog")?.addEventListener(
      "signature-line:insert",
      this.#onSignatureLineInsert as EventListener,
    );
    this.shadowRoot!.querySelector("docen-dropcap-dialog")?.addEventListener(
      "dropcap:apply",
      this.#onDropCapApply as EventListener,
    );
    this.shadowRoot!.querySelector("docen-merge-recipients-dialog")?.addEventListener(
      "recipients:updated",
      this.#onRecipientsUpdated as EventListener,
    );
    this.shadowRoot!.querySelector("docen-version-history-dialog")?.addEventListener(
      "version:restore",
      this.#onVersionRestore as EventListener,
    );
    this.shadowRoot!.querySelector("docen-restrict-editing-pane")?.addEventListener(
      "protection:enforce",
      this.#onProtectionEnforce as EventListener,
    );
    this.shadowRoot!.querySelector("docen-restrict-editing-pane")?.addEventListener(
      "protection:stop",
      this.#onProtectionStop as EventListener,
    );
    this.shadowRoot!.querySelector("docen-restrict-editing-pane")?.addEventListener(
      "protection:toggle-exception",
      this.#onProtectionToggleException as EventListener,
    );
    this.shadowRoot!.querySelector("docen-restrict-editing-pane")?.addEventListener(
      "protection:find-next",
      this.#onProtectionFindNext as EventListener,
    );
    this.shadowRoot!.querySelector("docen-reveal-formatting-pane")?.addEventListener(
      "reveal:compare-toggle",
      this.#onRevealCompareToggle as EventListener,
    );
    this.shadowRoot!.querySelector("docen-a11y-checker-pane")?.addEventListener(
      "a11y:select-issue",
      this.#onA11ySelectIssue as EventListener,
    );
    this.shadowRoot!.querySelector("docen-a11y-checker-pane")?.addEventListener(
      "a11y:refresh",
      () => {
        (this.shadowRoot?.querySelector("docen-a11y-checker-pane") as any)?.check(this.getJSON());
      },
    );
    const altTextPane = this.shadowRoot!.querySelector("docen-alt-text-pane");
    altTextPane?.addEventListener("alt-text:change", ((e: CustomEvent) => {
      const detail = e.detail;
      if (detail && this.editor) {
        (this.editor.commands as any)["drawing-alt-text"]?.(detail);
      }
    }) as EventListener);
    altTextPane?.addEventListener("alt-text:apply", ((e: CustomEvent) => {
      const detail = e.detail;
      if (detail && this.editor) {
        (this.editor.commands as any)["drawing-alt-text"]?.(detail);
      }
    }) as EventListener);
    this.shadowRoot!.querySelector("docen-sdt-dialog")?.addEventListener(
      "sdt-dialog:ok",
      this.#onSdtDialogOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-nav-pages")?.addEventListener(
      "nav-pages:select",
      this.#onNavPagesSelect as EventListener,
    );
    // Paragraph dialog — stamp the committed patch onto the selection (or the
    // targeted style when opened through the Modify Style dialog's Format).
    this.shadowRoot!.querySelector("docen-paragraph-dialog")?.addEventListener(
      "paragraph:ok",
      this.#onParagraphDialogOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-paragraph-dialog")?.addEventListener(
      "paragraph:open-tabs",
      this.#insert.openTabsDialog as EventListener,
    );
    // Paragraph dialog's Set As Default — the patch lands on the Normal style.
    this.shadowRoot!.querySelector("docen-paragraph-dialog")?.addEventListener(
      "paragraph:default",
      this.#dialogs.onParagraphDefault as EventListener,
    );
    // Page Setup dialog — write the committed geometry into the current
    // section (the Custom Margins / More Paper Sizes entries open it).
    this.shadowRoot!.querySelector("docen-page-setup-dialog")?.addEventListener(
      "page-setup:ok",
      this.#sections.onPageSetupOk as EventListener,
    );
    // Table grid — insert the picked shape through insert-table.
    this.shadowRoot!.querySelector("docen-table-dialog")?.addEventListener(
      "table-grid:insert",
      this.#onTableInsert as EventListener,
    );
    // Columns dialog — write the committed layout into the current section.
    this.shadowRoot!.querySelector("docen-columns-dialog")?.addEventListener(
      "columns:ok",
      this.#sections.onColumnsOk as EventListener,
    );
    // Link dialog — commit the hyperlink (mark / replace / insert / remove).
    this.shadowRoot!.querySelector("docen-link-dialog")?.addEventListener(
      "link:ok",
      this.#onLinkOk as EventListener,
    );
    // Zoom dialog — apply the preset or free percent; the status-bar percent
    // click opens it.
    this.shadowRoot!.querySelector("docen-zoom-dialog")?.addEventListener(
      "zoom:ok",
      this.#status.onZoomOk as EventListener,
    );
    // Paste Special dialog — the format pick re-runs #paste in that mode.
    this.shadowRoot!.querySelector("docen-paste-special-dialog")?.addEventListener(
      "paste-special:ok",
      this.#onPasteSpecialOk as EventListener,
    );
    // Font dialog — stamp the committed run state onto the selection (or the
    // targeted style when opened through the Modify Style dialog's Format).
    this.shadowRoot!.querySelector("docen-font-dialog")?.addEventListener(
      "font:ok",
      this.#onFontDialogOk as EventListener,
    );
    // Table Properties dialog — rewrite the caret table's alignment/indent.
    this.shadowRoot!.querySelector("docen-table-properties-dialog")?.addEventListener(
      "table-properties:ok",
      this.#dialogs.onTablePropertiesOk as EventListener,
    );
    // Size-and-Position dialog — restamp the selected drawing's geometry.
    this.shadowRoot!.querySelector("docen-drawing-properties-dialog")?.addEventListener(
      "drawing-properties:ok",
      this.#dialogs.onDrawingPropertiesOk as EventListener,
    );
    // Borders and Shading dialog — stamp the border/page/shading tab.
    this.shadowRoot!.querySelector("docen-borders-shading-dialog")?.addEventListener(
      "borders-shading:ok",
      this.#sections.onBordersShadingOk as EventListener,
    );
    // Custom watermark dialog — clear/stamp the header watermark.
    this.shadowRoot!.querySelector("docen-watermark-dialog")?.addEventListener(
      "watermark:ok",
      this.#design.onWatermarkOk as EventListener,
    );
    // Fill Effects dialog — set/clear the page's picture fill.
    this.shadowRoot!.querySelector("docen-fill-effects-dialog")?.addEventListener(
      "fill-effects:ok",
      this.#design.onFillEffectsOk as EventListener,
    );
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "zoom:open",
      this.#status.onZoomOpen as EventListener,
    );
    // Status-bar word count → the statistics dialog (Word).
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "wordcount:open",
      this.#status.onWordCountOpen as EventListener,
    );
    // Status-bar view shortcuts (Word's Reading / Print Layout / Web Layout
    // buttons) — the same `view` attribute the ribbon's View tab writes.
    this.shadowRoot!.querySelector("docen-status-bar")?.addEventListener(
      "view:select",
      this.#status.onViewSelect as EventListener,
    );

    this.#syncDir();
    this.#syncTouchMode();
    // Re-render header + ribbon when the page locale (<html lang>) changes.
    this.#unobserveLang = observeLang(() => {
      this.#syncDir();
      this.#renderChrome();
    });

    // Persisted settings (identity + writing toggles) — any store change
    // (this element's Options commit / setSettings, another <docen-document>,
    // an add-in) re-stamps the chrome and bubbles out as `docen:settings-change`.
    this.#settingsOff = onSettingsChange((settings) => {
      this.#renderChrome();
      // The hidden-text display is a projection input: a store change that
      // flips it re-renders the document (Options → Display, or another host
      // sharing the store).
      if (settings.writing.showHiddenText !== this.#hiddenTextShown && this.#stageHost) {
        this.#renderDoc(this.getJSON());
      }
      this.dispatchEvent(
        new CustomEvent("docen:settings-change", {
          bubbles: true,
          composed: true,
          detail: { settings },
        }),
      );
    });

    // Ribbon Display Options → drive browser fullscreen + status-bar hide.
    // auto-hide = Full Screen (Office); any other mode exits it.
    const ribbon = this.shadowRoot!.querySelector("docen-ribbon");
    ribbon?.addEventListener("ribbon-mode-change", this.#onRibbonModeChange);
    // Emit docen:change on every content change (autosave driver) and docen:ready
    // once the editor is live — both bubble out so a host can react.
    this.editor?.on("transaction", this.#onTransaction);
    // Selection moves repaint the anchored comment card (Word highlights the
    // card whose range the caret sits in).
    this.editor?.on("selectionUpdate", this.#comments.syncActiveCommentCard);
    this.editor?.on("selectionUpdate", this.#revisions.syncActiveRevision);
    this.editor?.on("selectionUpdate", this.#onSelectionUpdateForTranslate);
    document.addEventListener("fullscreenchange", this.#onFullscreenChange);
    this.addEventListener("keydown", this.#onZoomKey);
    this.dispatchEvent(new CustomEvent("docen:ready", { bubbles: true, composed: true }));
  }

  #readStorySource(kind: StoryKind, slot: StorySlot, page: number): JSONContent[] {
    return this.#stories.readSource(kind, slot, page);
  }

  #renderStoryFurniture(kind: StoryKind, slot: StorySlot, json: JSONContent[]): void {
    this.#stories.renderFurniture(kind, slot, json);
  }

  #exitStory(kind: StoryKind, slot: StorySlot, json: JSONContent[], dirty: boolean): void {
    this.#stories.exit(kind, slot, json, dirty);
  }

  #removeStory(kind: StoryKind): void {
    this.#stories.remove(kind);
  }

  /** Is this inline passthrough atom a PAGE (not NUMPAGES/PAGEREF) field? */
  static isPageField(child: JSONContent): boolean {
    if (child.type !== "inlinePassthrough") return false;
    try {
      const data = JSON.parse(String((child.attrs as { data?: string } | undefined)?.data)) as {
        simpleField?: { instruction?: string };
      };
      const instr = data.simpleField?.instruction?.trim().toUpperCase() ?? "";
      return instr.startsWith("PAGE") && !instr.startsWith("PAGES") && !instr.startsWith("PAGEREF");
    } catch {
      return false;
    }
  }

  #removePageNumbers(): void {
    this.#stories.removePageNumbers();
  }

  #pageInsets(
    flow: ProjectedFlowBox,
    furniture: ProjectedPageFurniture | undefined,
    laid: LaidFurnitureSection | undefined,
  ): FlowPageInsets | undefined {
    return pageInsets(flow, furniture, laid);
  }

  /** The PM node position of a drawing hit's target — the host paragraph's
   *  inner position via the caret map, then the index-th node of the hit's
   *  kind ("drawing" counts floating pictures + wps shapes + wpg groups,
   *  "inline" the non-floating images; a childPath targets a group member
   *  only while the group is entered). The resolution itself lives in the
   *  drawing domain (drawing/docx.ts). */
  #drawingNodePos(
    para: unknown,
    index: number,
    kind: "drawing" | "inline",
    childPath?: readonly number[],
    enterGroup?: boolean,
  ): number | null {
    const bridge = this.#bridge;
    if (!bridge) return null;
    return drawingNodePos(
      {
        doc: bridge.editor.state.doc,
        selection: bridge.editor.state.selection,
        posOfPara: (p) => bridge.posOfPara(p),
      },
      para,
      index,
      kind,
      childPath,
      enterGroup,
    );
  }

  /** The selected drawing's geometry in the dialog's display unit (cm for
   *  size and offsets, degrees for rotation) — an image sizes in px attrs
   *  while a shape/group payload sizes in EMU; null on any other selection. */
  #drawingStateOf(): DrawingPropertiesState | null {
    const sel = this.editor?.state.selection;
    if (!(sel instanceof NodeSelection)) return null;
    const attrs = sel.node.attrs as Record<string, unknown>;
    // A shape/group carries its payload (floating + transformation) under
    // wpsShape/wpgGroup; an image carries floating flat and sizes in px.
    const payload = (attrs.wpsShape ?? attrs.wpgGroup) as
      | Record<string, unknown>
      | null
      | undefined;
    const floating = (payload ? payload.floating : attrs.floating) as Record<
      string,
      unknown
    > | null;
    if (!floating) return null;
    const EMU_PER_CM = 360000;
    const PX_PER_CM = 96 / 2.54;
    const sizeDiv = payload ? EMU_PER_CM : PX_PER_CM;
    const t = (payload?.transformation ?? {}) as Record<string, unknown>;
    const num = (v: unknown): number => (typeof v === "number" && Number.isFinite(v) ? v : 0);
    const cm = (v: number): number => Math.round((v / sizeDiv) * 100) / 100;
    const offsetCm = (v: unknown): number => Math.round((num(v) / EMU_PER_CM) * 100) / 100;
    const hPos = floating.horizontalPosition as Record<string, unknown> | undefined;
    const vPos = floating.verticalPosition as Record<string, unknown> | undefined;
    const margins = floating.margins as Record<string, unknown> | undefined;
    return {
      widthCm: cm(payload ? num(t.width) : num(attrs.width)),
      heightCm: cm(payload ? num(t.height) : num(attrs.height)),
      rotationDeg: num(payload ? t.rotation : attrs.rotation),
      offsetHCm: offsetCm(hPos?.offset),
      offsetVCm: offsetCm(vPos?.offset),
      relativeH: typeof hPos?.relative === "string" ? hPos.relative : "column",
      relativeV: typeof vPos?.relative === "string" ? vPos.relative : "paragraph",
      allowOverlap: floating.allowOverlap === true,
      layoutInCell: floating.layoutInCell !== false,
      lockAnchor: floating.lockAnchor === true,
      distanceCm: {
        top: offsetCm(margins?.top),
        bottom: offsetCm(margins?.bottom),
        left: offsetCm(margins?.left),
        right: offsetCm(margins?.right),
      },
      altText: typeof attrs.title === "string" ? attrs.title : "",
    };
  }

  /** The selected node when it's a source-carrying image (crop's target) —
   *  null for shapes and every other selection. */
  #selectedImage(): { src: string } | null {
    const sel = this.editor?.state.selection;
    if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return null;
    const src = sel.node.attrs.src;
    return typeof src === "string" && src ? { src } : null;
  }

  /** Whether the selected image actually carries an a:srcRect crop (Reset
   *  Crop's enable — resetting an uncropped picture is a no-op). */
  #selectedImageHasCrop(): boolean {
    const sel = this.editor?.state.selection;
    if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return false;
    return sel.node.attrs.crop != null;
  }

  /** The selected or caret-adjacent OLE embedded object, if any. */
  #selectedOleObject(): { data: Uint8Array; fileName: string; progId: string } | null {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return null;
    const { state } = editor;
    const { selection } = state;
    let found: { data: Uint8Array; fileName: string; progId: string } | null = null;
    const checkNode = (node: PMNode): void => {
      if (node.type.name === "inlinePassthrough" && node.attrs?.data) {
        try {
          const parsed = decodePassthroughData<{
            object?: { embed?: { data?: Uint8Array; fileName?: string; progId?: string } };
          }>(node.attrs.data);
          if (parsed?.object?.embed) {
            const embed = parsed.object.embed;
            found = {
              data:
                embed.data instanceof Uint8Array
                  ? embed.data
                  : new Uint8Array((embed.data as any) ?? []),
              fileName: embed.fileName ?? "Microsoft_Excel_Worksheet.xlsx",
              progId: embed.progId ?? "Excel.Sheet.12",
            };
          }
        } catch {
          // ignore malformed data
        }
      }
    };

    if (selection instanceof NodeSelection) {
      checkNode(selection.node);
    } else if (selection.empty) {
      const $pos = selection.$from;
      if ($pos.nodeBefore) checkNode($pos.nodeBefore);
      if (!found && $pos.nodeAfter) checkNode($pos.nodeAfter);
    } else {
      state.doc.nodesBetween(selection.from, selection.to, (node) => {
        if (!found) checkNode(node);
      });
    }
    return found;
  }

  /** The active view, normalized (an unknown attr value reads as print). */
  #viewMode(): "print" | "web" | "draft" | "read" | "outline" {
    return this.view === "web" ||
      this.view === "draft" ||
      this.view === "read" ||
      this.view === "outline"
      ? this.view
      : "print";
  }

  /** Apply the `view` attribute: restage the render mode, re-project (the
   *  continuous views re-layout at the viewport width; Draft drops the
   *  furniture insets), trim the chrome in Read Mode, and sync the status
   *  bar. Word's Read Mode is read-only; the other views keep editability. */
  #applyView(): void {
    const mode = this.#viewMode();
    this.#stage?.setViewMode(mode);
    this.#syncReadChrome(mode === "read");
    this.#syncEditable();
    const outlineView = this.shadowRoot?.querySelector("docen-outline-view") as
      | (HTMLElement & { setEditor: (e: any) => void; syncFromEditor: () => void })
      | null;
    if (outlineView && this.editor) {
      outlineView.setEditor(this.editor);
    }
    this.#renderDoc(this.getJSON());
    this.#updateStatus();
  }

  /** The one editable computation: host attr ∧ view mode ∧ document
   *  protection. Every path that derives editability from state calls this —
   *  protection changes just flip #docProtected and re-run it. */
  #syncEditable(): void {
    if (!this.editor) return;
    const hasPerms = hasPermissionRanges(this.editor.state.doc);
    const effectiveProtected =
      this.#docProtected && !(this.#protectionMode === "readOnly" && hasPerms);
    const editable =
      this.editable !== "false" && this.#viewMode() !== "read" && !effectiveProtected;
    if (this.editor.isEditable !== editable) {
      this.editor.setEditable(editable);
      this.#syncEditModeMenu();
    }
  }

  /** Read Mode trims the chrome to the document (Word hides the ribbon and
   *  most of the tab row). */
  #syncReadChrome(read: boolean): void {
    const ribbon = this.shadowRoot?.querySelector("docen-ribbon");
    if (ribbon) ribbon.toggleAttribute("hidden", read);
  }

  #projectAndLayout(doc: JSONContent): {
    pages: FlowPage[];
    sectionOfPage: number[];
    sections: (ProjectedSection & CanvasStageSection)[];
    background?: ProjectedPageBackground;
  } {
    return this.#render.projectAndLayout(doc);
  }

  #fieldFrame(pos: number): FieldFrame | undefined {
    return this.#render.fieldFrame(pos);
  }

  /** Alt+F9 — Word's field-code display: every field projects its instruction
   *  text instead of its cached result until toggled back. */
  toggleFieldCodes(): void {
    this.#fieldCodes = !this.#fieldCodes;
    this.#renderDoc(this.getJSON());
  }

  #pageOriginOf(
    sections: readonly (ProjectedSection & CanvasStageSection)[],
    sectionOfPage: readonly number[],
  ): (page: number) => { contentLeftPx: number; contentTopPx: number } {
    return this.#stories.pageOriginOf(sections, sectionOfPage);
  }

  #renderDoc(doc: JSONContent): void {
    this.#render.renderDoc(doc);
    // Page geometry may have changed (page setup, sections) — keep the
    // interactive ruler's width/margins in step.
    this.#syncRuler();
    this.#updateContentWarning(doc);
  }

  /** Show/refresh the unsupported-content warning bar from the document JSON
   *  (item 14): content the editor can preserve but not edit — altChunk,
   *  subDoc, SmartArt, OLE, raw/custom XML, content parts — must be visible to
   *  the user, never silently carried. Hidden while the user has dismissed it
   *  for this document. */
  #updateContentWarning(doc: JSONContent): void {
    const bar = this.shadowRoot?.querySelector<HTMLElement>(".content-warning");
    if (!bar) return;
    const translate = (key: string): string => t(key, this);
    const close = bar.querySelector<HTMLElement>(".content-warning-close");
    close?.setAttribute("aria-label", translate("contentWarning.dismiss"));
    close?.setAttribute("title", translate("contentWarning.dismiss"));
    const report = detectUnsupportedContent(doc);
    if (report.total === 0 || this.#contentWarningDismissed) {
      bar.hidden = true;
      return;
    }
    const text = bar.querySelector<HTMLElement>(".content-warning-text");
    if (text) text.textContent = contentWarningText(report, translate);
    bar.hidden = false;
  }

  #dismissContentWarning(): void {
    this.#contentWarningDismissed = true;
    const bar = this.shadowRoot?.querySelector<HTMLElement>(".content-warning");
    if (bar) bar.hidden = true;
  }

  /** Mount the interactive horizontal ruler above the pages (idempotent).
   *  The stage keeps only the vertical strip; this strip owns the draggable
   *  indent markers, tab stops and the unit toggle. */
  #mountRuler(): void {
    if (this.#ruler || !this.#stageHost) return;
    const area = this.#stageHost.closest("docen-document-area");
    if (!area) return;
    const ruler = document.createElement("docen-ruler") as InteractiveRulerElement;
    Object.assign(ruler.style, {
      position: "sticky",
      top: "0",
      zIndex: "6",
      margin: "0 auto",
      display: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    ruler.addEventListener("ruler:open-tabs", () => this.#insert.openTabsDialog());
    // Direct child of the scroll container (NOT the flex canvas wrapper, where
    // a second flex item would break the page centering).
    area.insertBefore(ruler, area.firstChild);
    this.#ruler = ruler;
  }

  /** Mirror the interactive ruler for RTL paragraphs (Word flips the scale at
   *  the caret) and for an RTL shell direction; LTR paragraphs in an LTR shell
   *  keep Word's left-to-right scale. */
  #syncRulerDirection(): void {
    const ruler = this.#ruler;
    if (!ruler) return;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    let bidi = false;
    if (editor) {
      const { $from } = editor.state.selection;
      for (let d = $from.depth; d > 0; d--) {
        const node = $from.node(d);
        if (node.type.name === "paragraph" || node.type.name === "heading") {
          bidi = node.attrs.bidirectional === true;
          break;
        }
      }
    }
    const dir = bidi || this.getAttribute("dir") === "rtl" ? "rtl" : "ltr";
    if (ruler.getAttribute("dir") !== dir) ruler.setAttribute("dir", dir);
  }

  /** Show/hide and re-geometry the interactive ruler from the current flow
   *  box + zoom. Cheap when nothing changed (signature-guarded). */
  #syncRuler(): void {
    const ruler = this.#ruler;
    if (!ruler) return;
    this.#syncRulerDirection();
    const on = this.getShowRuler();
    ruler.style.display = on ? "block" : "none";
    if (!on) return;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (editor && editor !== this.#rulerEditor) {
      this.#rulerEditor = editor;
      ruler.bindEditor(editor);
    }
    const flow = this.#flow;
    if (!flow) return;
    const zoom = this.#stage ? this.#stage.zoom / 100 : 1;
    const geometry = {
      pageWidthPx: flow.pageWidthPx,
      marginLeftPx: flow.contentLeftPx,
      marginRightPx: flow.pageWidthPx - flow.contentLeftPx - flow.contentWidthPx,
      scale: zoom,
    };
    const key = `${geometry.pageWidthPx}|${geometry.marginLeftPx}|${geometry.marginRightPx}|${zoom}`;
    if (key === this.#rulerGeometry) return;
    this.#rulerGeometry = key;
    ruler.style.width = `${geometry.pageWidthPx * zoom}px`;
    ruler.setParagraphAttrs(undefined, undefined, geometry);
  }

  /** Create the stage on first use and refresh its per-render context —
   *  idempotent setters both render paths call before their first sync. */
  #armStage(p: {
    sections: (ProjectedSection & CanvasStageSection)[];
    background?: ProjectedPageBackground;
    viewMode: "print" | "web" | "draft" | "read" | "outline";
  }): CanvasStage {
    this.#stage ??= new CanvasStage(this.#stageHost!, {
      metrics: browserFontMetrics,
      sections: p.sections,
      sectionOfPage: [],
      background: p.background,
    });
    // Viewport virtualization → overlay culling (the bridge paints squiggles,
    // selection and search only on pages the stage keeps painted).
    this.#stage.onLiveChange = (page, live) => this.#bridge?.setPageLive(page, live);
    // A debug attribute stamped before the first render lands here.
    if (this.debug) this.#stage.setDebug(this.debug);
    this.#stage.setMarksLabels({
      pageBreak: t("marks.pageBreak", this),
      sectionBreak: t("marks.sectionBreak", this),
      sectionBreakContinuous: t("marks.sectionBreakContinuous", this),
      sectionBreakEvenPage: t("marks.sectionBreakEvenPage", this),
      sectionBreakOddPage: t("marks.sectionBreakOddPage", this),
    });
    this.#stage.setFieldShading(this.#fieldShading);
    // A `zoom` attribute parsed before the stage existed only recorded the
    // level here — push it in before the first sync sizes the slots. The
    // `show-marks` and `view` attributes get the same once-over (idempotent
    // setters; the read-only + chrome trimming rides #applyView's gate).
    if (this.#stage.zoom !== this.#status.getZoom()) this.#stage.setZoom(this.#status.getZoom());
    if (this.hasAttribute("show-marks")) this.#stage.setShowMarks(true);
    if (this.hasAttribute("show-ruler") || this.hasAttribute("ruler"))
      this.#stage.setShowRuler(true);
    this.#syncRuler();
    if (this.#stage.viewMode !== p.viewMode) {
      this.#stage.setViewMode(p.viewMode);
      this.#syncReadChrome(p.viewMode === "read");
    }
    this.#syncEditable();
    return this.#stage;
  }

  /** The previous render's flow result — the diff base for the next one. */
  #lastRun?: {
    pages: FlowPage[];
    sectionOfPage: number[];
    sections: (ProjectedSection & CanvasStageSection)[];
    background?: ProjectedPageBackground;
    viewMode?: "print" | "web" | "draft" | "read" | "outline";
  };

  /** Bumped by every render — an incremental layout walk compares its capture
   *  against this each slice and drops when another render has started. */
  #renderSeq = 0;

  disconnectedCallback(): void {
    this.#clipboard.hidePasteOptions();
    clearTimeout(this.#a11yTimer);
    this.#langObserver?.disconnect();
    this.#unobserveLang?.();
    this.#unobserveLang = undefined;
    this.shadowRoot?.removeEventListener("command", this.#onCommand as EventListener);
    this.shadowRoot?.removeEventListener("item-context", this.#onItemContext as EventListener);
    this.shadowRoot?.removeEventListener("item-preview", this.#onItemPreview as EventListener);
    this.shadowRoot?.removeEventListener(
      "item-preview-end",
      this.#onItemPreviewEnd as EventListener,
    );
    this.shadowRoot?.removeEventListener("change", this.#onChange as EventListener);
    this.#fileInput?.removeEventListener("change", this.#io.onFileChange);
    this.#imageInput?.removeEventListener("change", this.#io.onImageChange);
    this.#pictureInput?.removeEventListener("change", this.#io.onPictureChange);
    this.shadowRoot
      ?.querySelector("docen-outline")
      ?.removeEventListener("outline:select", this.#navigation.onOutlineSelect as EventListener);
    this.removeEventListener("navigation:search", this.#navigation.onSearch as EventListener);
    this.removeEventListener("navigation:find", this.#navigation.onFind as EventListener);
    this.shadowRoot
      ?.querySelector(".search-results")
      ?.removeEventListener("click", this.#navigation.onSearchResultClick as EventListener);
    this.shadowRoot
      ?.querySelector("docen-find-replace-dialog")
      ?.removeEventListener("find-replace:action", this.#navigation.onFindReplace as EventListener);
    this.shadowRoot
      ?.querySelector("docen-options-dialog")
      ?.removeEventListener("options:ok", this.#onOptionsOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-options-dialog")
      ?.removeEventListener("options:autocorrect", this.#openAutocorrectDialog as EventListener);
    this.shadowRoot
      ?.querySelector("docen-autocorrect-dialog")
      ?.removeEventListener("autocorrect:ok", this.#onAutocorrectOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-quick-part-dialog")
      ?.removeEventListener("quick-part:save", this.#onQuickPartSave as EventListener);
    for (const buildingBlocksEvent of [
      "building-blocks:insert",
      "building-blocks:rename",
      "building-blocks:delete",
    ]) {
      this.shadowRoot
        ?.querySelector("docen-building-blocks-dialog")
        ?.removeEventListener(buildingBlocksEvent, this.#onBuildingBlockEvent as EventListener);
    }
    this.shadowRoot
      ?.querySelector("docen-note-settings-dialog")
      ?.removeEventListener("note-settings:ok", this.onNoteSettingsOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-note-settings-dialog")
      ?.removeEventListener("note-settings:convert", this.onNoteSettingsConvert as EventListener);
    this.shadowRoot
      ?.querySelector("docen-line-numbers-dialog")
      ?.removeEventListener("line-numbers:ok", this.#sections.onLineNumbersOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-hyphenation-dialog")
      ?.removeEventListener("hyphenation:ok", this.#insert.onHyphenationOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-tabs-dialog")
      ?.removeEventListener("tabs:ok", this.#insert.onTabsOk as EventListener);
    const bmEl = this.shadowRoot?.querySelector("docen-bookmark-dialog");
    bmEl?.removeEventListener("bookmark:add", this.onBookmarkAdd as EventListener);
    bmEl?.removeEventListener("bookmark:delete", this.onBookmarkDelete as EventListener);
    bmEl?.removeEventListener("bookmark:goto", this.onBookmarkGoTo as EventListener);
    this.shadowRoot
      ?.querySelector("docen-go-to-dialog")
      ?.removeEventListener("goto:navigate", this.onGoToNavigate as EventListener);
    this.shadowRoot
      ?.querySelector("docen-properties-dialog")
      ?.removeEventListener("properties:ok", this.onPropertiesOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("language:open", this.#onLanguageOpen as EventListener);
    this.shadowRoot
      ?.querySelector("docen-language-dialog")
      ?.removeEventListener("language:ok", this.#dialogs.onLanguageOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-phonetic-dialog")
      ?.removeEventListener("phonetic:ok", this.#dialogs.onPhoneticOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-phonetic-dialog")
      ?.removeEventListener("phonetic:clear", this.#dialogs.onPhoneticClear as EventListener);
    this.shadowRoot
      ?.querySelector("docen-two-in-one-dialog")
      ?.removeEventListener("two-in-one:ok", this.#dialogs.onTwoInOneOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-define-list-dialog")
      ?.removeEventListener("define-list:ok", this.#dialogs.onDefineListOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-caption-dialog")
      ?.removeEventListener("caption:ok", this.#dialogs.onCaptionOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-note-dialog")
      ?.removeEventListener("note:ok", this.#dialogs.onNoteOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-field-dialog")
      ?.removeEventListener("field:ok", this.#dialogs.onFieldOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-chart-data-dialog")
      ?.removeEventListener("chart:ok", this.#dialogs.onChartOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-chart-type-dialog")
      ?.removeEventListener("chart-type:ok", this.#dialogs.onChartTypeOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-cross-reference-dialog")
      ?.removeEventListener("cross-ref:ok", this.#dialogs.onCrossRefOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-sources-dialog")
      ?.removeEventListener("sources:ok", this.#references.onSourcesOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-sources-dialog")
      ?.removeEventListener("citation:ok", this.#references.onCitationOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-recipients-dialog")
      ?.removeEventListener("recipients:ok", this.#merge.onRecipientsOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-merge-field-dialog")
      ?.removeEventListener("merge-field:ok", this.#onMergeFieldOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-symbol-dialog")
      ?.removeEventListener("symbol:insert", this.#onSymbolInsert as EventListener);
    this.shadowRoot
      ?.querySelector("docen-paragraph-dialog")
      ?.removeEventListener("paragraph:ok", this.#onParagraphDialogOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-paragraph-dialog")
      ?.removeEventListener("paragraph:default", this.#dialogs.onParagraphDefault as EventListener);
    this.shadowRoot
      ?.querySelector("docen-paragraph-dialog")
      ?.removeEventListener("paragraph:open-tabs", this.#insert.openTabsDialog as EventListener);
    this.shadowRoot
      ?.querySelector("docen-paste-special-dialog")
      ?.removeEventListener("paste-special:ok", this.#onPasteSpecialOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-font-dialog")
      ?.removeEventListener("font:ok", this.#onFontDialogOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-table-properties-dialog")
      ?.removeEventListener(
        "table-properties:ok",
        this.#dialogs.onTablePropertiesOk as EventListener,
      );
    this.shadowRoot
      ?.querySelector("docen-borders-shading-dialog")
      ?.removeEventListener(
        "borders-shading:ok",
        this.#sections.onBordersShadingOk as EventListener,
      );
    this.shadowRoot
      ?.querySelector("docen-watermark-dialog")
      ?.removeEventListener("watermark:ok", this.#design.onWatermarkOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-fill-effects-dialog")
      ?.removeEventListener("fill-effects:ok", this.#design.onFillEffectsOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-page-setup-dialog")
      ?.removeEventListener("page-setup:ok", this.#sections.onPageSetupOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-table-dialog")
      ?.removeEventListener("table-grid:insert", this.#onTableInsert as EventListener);
    this.shadowRoot
      ?.querySelector("docen-columns-dialog")
      ?.removeEventListener("columns:ok", this.#sections.onColumnsOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-link-dialog")
      ?.removeEventListener("link:ok", this.#onLinkOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-zoom-dialog")
      ?.removeEventListener("zoom:ok", this.#status.onZoomOk as EventListener);
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("zoom:open", this.#status.onZoomOpen as EventListener);
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("wordcount:open", this.#status.onWordCountOpen as EventListener);
    this.shadowRoot
      ?.querySelector<HTMLElement>("docen-status-bar")
      ?.removeEventListener("zoom:change", this.#status.onZoomChange as EventListener);
    this.shadowRoot
      ?.querySelector("docen-status-bar")
      ?.removeEventListener("view:select", this.#status.onViewSelect as EventListener);
    this.#stageHost?.removeEventListener("wheel", this.#onWheel as EventListener, {
      capture: true,
    });
    this.editor?.off("transaction", this.#onTransaction);
    this.editor?.off("selectionUpdate", this.#comments.syncActiveCommentCard);
    this.editor?.off("selectionUpdate", this.#onSelectionUpdateForTranslate);
    this.removeEventListener("docen:translate", this.#onDocenTranslate as EventListener);
    document.removeEventListener("fullscreenchange", this.#onFullscreenChange);
    this.removeEventListener("keydown", this.#onZoomKey);
    this.shadowRoot
      ?.querySelector("docen-ribbon")
      ?.removeEventListener("ribbon-mode-change", this.#onRibbonModeChange);
    this.#fontSyncCleanup?.();
    this.#fontSyncCleanup = undefined;
    this.#settingsOff?.();
    this.#settingsOff = undefined;
    this.#chrome.dispose();
    this.#status.dispose();
    this.#stopFormatPainter();
    this.#stopBorderPainting();
    this.#stopShapeDrawing();
    this.#navigation.dispose();
    this.#spelling.dispose();
    this.#bridge?.destroy();
    this.#bridge = undefined;
    this.#stage?.destroy();
    this.#stage = undefined;
    super.disconnectedCallback();
  }

  // ── Quick Access Toolbar (title bar) ──────────────────────────────────────

  #toggleQat(id: string): void {
    this.#chrome.toggleQat(id);
  }

  // ── Auto-save (title bar switch) ──────────────────────────────────────────

  #autosaveEnabled(): boolean {
    return this.#chrome.autosaveEnabled();
  }

  #setAutosave(on: boolean): void {
    this.#chrome.setAutosave(on);
  }

  #loadPersistedHistory(): void {
    this.#chrome.loadPersistedHistory();
  }

  #scheduleAutosave(): void {
    this.#chrome.scheduleAutosave();
  }

  #renderChrome(): void {
    this.#chrome.renderChrome();
  }

  #fillHistory(kind: "undo" | "redo"): void {
    this.#chrome.fillHistory(kind);
  }

  /** Shadow-root click delegation for the history flyouts: the split's caret
   *  fills its item list as Fluent opens the drop-down (the primary's click is
   *  stopped inside the split, so only caret/menu clicks reach here). */
  readonly #onHistoryClick = (event: Event): void => {
    const split = (event.target as HTMLElement | null)?.closest?.(
      "docen-ribbon-split-button[data-history]",
    );
    if (split) {
      this.#fillHistory(split.getAttribute("data-history") as "undo" | "redo");
    }
  };

  /** Addin registry changed (add-in registered/removed) — re-stamp the ribbon
   *  so an external add-in's ribbon contribution appears. The default add-in
   *  contributes no ribbon, so this is a no-op for it; only extra add-ins add
   *  tabs. */
  protected addinsChanged(): void {
    this.#renderChrome();
  }

  #applyAddinsAttr(): void {
    this.#chrome.applyAddinsAttr();
  }

  #applyThemeAttr(value: string): void {
    this.#chrome.applyThemeAttr(value);
  }

  #syncArrangeGreying(): void {
    this.#chrome.syncArrangeGreying();
  }

  #syncFormatButtons(): void {
    this.#chrome.syncFormatButtons();
  }

  #syncDrawingMenus(): void {
    this.#chrome.syncDrawingMenus();
  }

  #syncQuickPartsMenu(): void {
    this.#chrome.syncQuickPartsMenu();
  }

  #syncEditModeMenu(): void {
    this.#chrome.syncEditModeMenu();
  }

  #syncMarkupMenus(): void {
    this.#chrome.syncMarkupMenus();
  }

  #syncStoryMenus(): void {
    this.#chrome.syncStoryMenus();
  }

  #hideHeaderFooterContextTab(): void {
    this.#chrome.hideHeaderFooterContextTab();
  }

  #syncCellSize(): void {
    this.#chrome.syncCellSize();
  }

  #syncDrawingSize(): void {
    this.#chrome.syncDrawingSize();
  }

  #syncTableLookCheckboxes(): void {
    this.#chrome.syncTableLookCheckboxes();
  }

  #syncContextTabs(): void {
    this.#chrome.syncContextTabs();
  }

  #emitCancelable(
    name:
      | "docen:save"
      | "docen:save-as"
      | "docen:open"
      | "docen:new"
      | "docen:print"
      | "docen:close",
    detail?: { format?: SaveFormat },
  ): boolean {
    return this.#chrome.emitCancelable(name, detail);
  }

  /** docen:change — fired on every doc-changing transaction (autosave driver,
   *  mirroring OnlyOffice's onDocumentStateChange). Selection-only transactions
   *  are skipped. */
  readonly #onTransaction = (props: { transaction: Transaction }): void => {
    // The status-bar language mirrors the caret's proofing language (Word).
    // Marked dirty here and flushed with the frame's UI sync (the listener
    // #setupFontSync registers already schedules the frame).
    if (props.transaction.selectionSet) {
      this.#uiSelectionDirty = true;
      this.#scheduleUiSync();
    }
    if (props.transaction.docChanged) {
      this.#jsonDirty = true;
      this.#spelling.mapThrough(props.transaction);
      this.dispatchEvent(
        new CustomEvent("docen:change", { bubbles: true, composed: true, detail: { dirty: true } }),
      );
      if (this.#autosaveEnabled()) this.#scheduleAutosave();
    }
  };

  #togglePane(id: TaskPaneId): void {
    this.#chrome.togglePane(id);
  }

  /** Parse the declarative `section-properties` / `styles` attributes (JSON).
   *  Lets a host bootstrap page setup + named styles without openDOCX/setJSON.
   *  Malformed JSON is ignored (warned) so a typo never breaks the editor. */
  #readInitAttrs(): {
    sectionProperties?: SectionPropertiesOptions;
    styles?: StylesOptions;
  } {
    const out: { sectionProperties?: SectionPropertiesOptions; styles?: StylesOptions } = {};
    const sp = this.getAttribute("section-properties");
    if (sp) {
      try {
        out.sectionProperties = JSON.parse(sp) as SectionPropertiesOptions;
      } catch {
        console.warn("[docen-document] invalid section-properties JSON — ignored");
      }
    }
    const st = this.getAttribute("styles");
    if (st) {
      try {
        out.styles = JSON.parse(st) as StylesOptions;
      } catch {
        console.warn("[docen-document] invalid styles JSON — ignored");
      }
    }
    return out;
  }

  /** Runtime `section-properties` change: deep-merge into the body section's
   *  sectPr (a default doc is single-section); the dispatched transaction
   *  re-renders the canvas with the new geometry. */
  #applySectionPropertiesAttr(): void {
    const editor = this.editor;
    if (!editor) return;
    const parsed = this.#readInitAttrs().sectionProperties;
    if (!parsed) return;
    const cur = (editor.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions })
      .sectionProperties;
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("sectionProperties", mergeSectionProperties(cur, parsed)),
    );
  }

  /** Runtime `styles` change: replace doc.attrs.styles (the style library
   *  re-renders through the layout pipeline and the Styles gallery). */
  #applyStylesAttr(): void {
    const editor = this.editor;
    if (!editor) return;
    const parsed = this.#readInitAttrs().styles;
    if (parsed === undefined) return;
    editor.view.dispatch(editor.state.tr.setDocAttribute("styles", parsed));
    this.#renderChrome();
  }

  #setZoom(pct: number): void {
    this.#status.setZoom(pct);
  }

  /** Paste Special's pick — re-run the paste in the picked format. */
  readonly #onPasteSpecialOk = (event: CustomEvent<PasteSpecialFormat>): void => {
    void this.#clipboard.pasteSpecial(event.detail);
  };

  #updateStatus(): void {
    this.#status.updateStatus();
  }

  #refreshNavThumbnails(): Promise<void> {
    return this.#status.refreshNavThumbnails();
  }

  /** Word Count (Review tab) — compute the document statistics twice (Word's
   *  dialog shape): the body alone, and with textboxes + footnotes/endnotes
   *  folded back in — the dialog's "include" toggle (default ON) switches
   *  between the two readouts. Textboxes are wpsShape subtrees and textbox
   *  nodes in the body; the notes live in the documentExtras channels. */
  // ── Picture pixel tools (Picture Format → Adjust) ──

  /** The selected picture's attrs, or null when the selection isn't one. */
  #selectedPictureAttrs(): Record<string, unknown> | null {
    const editor = this.editor;
    if (!editor) return null;
    const sel = editor.state.selection;
    if (!(sel instanceof NodeSelection) || sel.node.type.name !== "image") return null;
    return sel.node.attrs as Record<string, unknown>;
  }

  #showCompressPictures(): void {
    const dialog = this.shadowRoot?.querySelector("docen-compress-pictures-dialog") as
      | (HTMLElement & { show(): void })
      | undefined;
    if (!this.#selectedPictureAttrs() || !dialog) return;
    dialog.show();
  }

  readonly #onCompressOk = (
    event: CustomEvent<{ ppi: number | null; dropCrop: boolean }>,
  ): void => {
    const editor = this.editor;
    const attrs = this.#selectedPictureAttrs();
    if (!editor || !attrs || typeof attrs.src !== "string") return;
    const { ppi, dropCrop } = event.detail ?? { ppi: 220, dropCrop: false };
    // The resample size is the display size at the target ppi (the frame
    // keeps its CSS size — only the pixel density changes); "keep current
    // resolution" re-encodes without resampling. Never upscales.
    const width = typeof attrs.width === "number" ? attrs.width : 400;
    const height = typeof attrs.height === "number" ? attrs.height : 300;
    const target =
      ppi == null ? { width, height } : { width: (width / 96) * ppi, height: (height / 96) * ppi };
    void compressPictureSrc(attrs.src, target, attrs.crop as CropRect | undefined, dropCrop)
      .then((src) => {
        editor.commands["picture-pixels"](JSON.stringify({ src, dropCrop }));
      })
      .catch(() => {
        /* a source that no longer decodes keeps the current one */
      });
  };

  /** Arm the canvas eyedropper: the next press on a picture re-encodes it
   *  with the sampled color cleared (Word's Set Transparent Color). */
  #armTransparentPick(): void {
    const bridge = this.#bridge;
    const editor = this.editor;
    if (!bridge || !editor) return;
    bridge.setTransparentPick((hit, nx, ny) => {
      const pos = this.#drawingNodePos(hit.para, hit.index, hit.kind, hit.childPath);
      if (pos == null) return;
      const node = editor.state.doc.nodeAt(pos);
      if (!node || node.type.name !== "image" || typeof node.attrs.src !== "string") return;
      void pickTransparentColor(
        node.attrs.src as string,
        nx,
        ny,
        node.attrs.crop as CropRect | undefined,
      )
        .then((src) => {
          // patchPicture re-checks the selection, so a press elsewhere in
          // between simply declines the swap.
          editor.commands["picture-pixels"](JSON.stringify({ src }));
        })
        .catch(() => {
          /* an undecodable source keeps the current pixels */
        });
    });
  }

  /** Symbol dialog Insert → drop the picked character at the caret (the
   *  dialog stays open, Word-style, so several symbols can go in a row). */
  readonly #onSymbolInsert = (event: CustomEvent<{ char?: string }>): void => {
    const char = event.detail?.char;
    if (!char) return;
    this.#bridge?.focus();
    this.editor?.commands.insertContent(char);
  };

  #setHyphenation(mode: "none" | "auto" | "manual"): void {
    this.#insert.setHyphenation(mode);
  }

  #openHyphenationOptions(): void {
    this.#insert.openHyphenationOptions();
  }

  #insertSoftHyphen(): void {
    this.#insert.insertSoftHyphen();
  }

  // The Paragraph dialog's OK — stamp its patch onto every selected paragraph
  // in the editor input currently routes into (a furniture story's editor
  // while a story is open, else the main document).
  // The Font dialog's prefill — the selection's first text run decides every
  // field (Word reads the same way; a mixed-format selection shows the first
  // run's values). Underline falls back to the textStyle attr channel when no
  // underline mark is present (both carry the same w:u shape).
  #runStateOf(state: EditorState): FontDialogPatch {
    return this.#insert.runStateOf(state);
  }

  // The table grid's pick (hover grid or the classic dialog shape) — the
  // engine's insert-table takes rows/cols (Word's 3×3 preset is the default).
  readonly #onTableInsert = (event: CustomEvent<{ rows?: number; cols?: number }>): void => {
    const { rows, cols } = event.detail ?? {};
    const target = this.#bridge?.activeEditor() ?? this.editor;
    target?.commands["insert-table"]?.({ rows, cols });
  };

  #insertBookmark(): void {
    this.#insert.insertBookmark();
  }

  #insertEquation(template: string): void {
    this.#insert.insertEquation(template);
  }

  #insertEquationSymbol(char: string): void {
    this.#insert.insertEquationSymbol(char);
  }

  /** Merge-field dialog commit — seed the picked field at the caret. */
  readonly #onMergeFieldOk = (event: Event): void => {
    const { name } = (event as CustomEvent<{ name?: string }>).detail ?? {};
    if (name) this.#merge.insertMergeField(name);
  };

  #mergedView(doc: JSONContent): JSONContent {
    return this.#insert.mergedView(doc);
  }

  /** Finish & Merge — assemble one document from the recipient rows (Word's
   *  Edit Individual Documents): "letters" clones the body once per row as
   *  its own section (each clone's last paragraph carries the body sectPr,
   *  closing the section the way OOXML wants); "directory" streams every
   *  record into one continuous body. The result downloads as a .docx. */
  async #finishMerge(mode: "edit" | "print" | "email"): Promise<void> {
    if (mode !== "edit" || !this.editor) return;
    const recipients = this.#merge.recipients();
    if (!recipients) return;
    // Prepare a copy once — the per-row merges below then carry whatever the
    // local (network-free) default preparation embedded.
    const json = await prepareDocument(this.getJSON());
    const body = json.content ?? [];
    const bodyAttrs = (json.attrs ?? {}) as {
      sectionProperties?: SectionPropertiesOptions | null;
    };
    const rows =
      this.#merge.mergeType() === "directory"
        ? [recipients.rows.map((_, i) => i)]
        : recipients.rows.map((_, i) => [i]);
    const content: JSONContent[] = [];
    rows.forEach((rowIndexes, copy) => {
      const last = copy < rows.length - 1;
      for (const row of rowIndexes) {
        const view = applyRecipientsRow({ ...json, content: body }, recipients, row);
        for (const [index, block] of (view.content ?? []).entries()) {
          if (last && block.type === "paragraph" && index === (view.content?.length ?? 0) - 1) {
            content.push({
              ...block,
              attrs: {
                ...block.attrs,
                sectionProperties: structuredClone(bodyAttrs.sectionProperties ?? null),
              },
            });
          } else {
            content.push(block);
          }
        }
      }
    });
    const buffer = await generateDOCX(
      { ...json, content },
      { prepare: false, packer: { type: "uint8array" } },
    );
    const blob = new Blob([buffer as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "mail-merge.docx";
    anchor.click();
    URL.revokeObjectURL(url);
  }

  #insertLink(): void {
    this.#insert.insertLink();
  }

  // The Link dialog's OK — Word's Insert Link semantics: an empty address
  // removes an existing link; a selection gets marked (its text replaced when
  // the dialog's display text was edited); an empty selection inserts fresh
  // display text carrying the mark.
  readonly #onLinkOk = (event: CustomEvent<LinkValues | undefined>): void => {
    const values = event.detail;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!values || !editor) return;
    this.#bridge?.focus();
    const raw = values.href.trim();
    // The link mark riding the selection, if any.
    const existing = editor.getAttributes("link").href as string | undefined;
    if (raw === "") {
      if (existing) editor.chain().focus().extendMarkRange("link").unsetLink().run();
      return;
    }
    // `#name` stays a bookmark anchor; bare hosts gain the https scheme.
    const href = raw.startsWith("#") || /^[a-z][a-z0-9+.-]*:/i.test(raw) ? raw : `https://${raw}`;
    const mark = [
      { type: "link", attrs: { href, target: href.startsWith("#") ? null : "_blank" } },
      { type: "textStyle", attrs: { style: "Hyperlink" } },
    ] as const;
    const { empty, from, to } = editor.state.selection;
    if (!empty) {
      const text = values.text.trim();
      const selected = editor.state.doc.textBetween(from, to, " ");
      if (text && text !== selected) {
        // The display text was edited — replace the selection with the fresh
        // marked run (one undo step).
        editor.commands.insertContentAt({ from, to }, { type: "text", text, marks: [...mark] });
        return;
      }
      // Word stamps hyperlink runs with the "Hyperlink" character style —
      // that style (not the w:hyperlink element) paints links blue.
      editor.chain().focus().extendMarkRange("link").setLink({ href }).run();
      editor.commands.setMark("textStyle", { style: "Hyperlink" });
      return;
    }
    // No selection: the display text inserts marked.
    const text = values.text.trim();
    if (!text) return;
    editor.commands.insertContent({ type: "text", text, marks: [...mark] });
  };

  /** The href of the link mark at the caret (the context menu's open/copy
   *  source; the right-click collapsed the caret onto the link first), or
   *  null. */
  #hrefAtCaret(): string | null {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return null;
    const mark = editor.state.doc
      .resolve(Math.min(editor.state.selection.from, editor.state.doc.content.size))
      .marks()
      .find((m) => m.type.name === "link");
    const href = mark?.attrs.href;
    return typeof href === "string" && href ? href : null;
  }

  #jumpToBookmark(name: string): void {
    this.#insert.jumpToBookmark(name);
  }

  #jumpNextNote(): void {
    this.#insert.jumpNextNote();
  }

  #jumpPreviousNote(): void {
    this.#insert.jumpPreviousNote();
  }

  /** Right-click on the canvas — Word's context menu, rebuilt per click.
   *  Clicking outside the selection first moves the caret there (Word's
   *  behavior), the clipboard section appears only with a selection, and a
   *  click on a hyperlink swaps the Link item for Edit/Remove. Menu items
   *  dispatch the same command ids as the ribbon, so #onCommand handles them.
   *  While a furniture story (header/footer) is being edited the positions
   *  belong to the story's editor, which these main-story commands cannot
   *  target — suppress the menu there. */
  readonly #onContextMenu = (event: MouseEvent): void => {
    const menu = this.shadowRoot?.querySelector("docen-context-menu") ?? null;
    const editor = this.editor;
    if (!menu || !editor || !event.composedPath().includes(menu)) return;
    if (this.#bridge?.storyKind() != null) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    // Word: a right-click on a picture selects it and shows the picture menu
    // instead of the text menu — the caret-based branches below never run.
    if (this.#bridge?.selectDrawingAtClient(event.clientX, event.clientY)) {
      const items: RibbonMenuItem[] = [];
      // The clipboard entries work on a NodeSelection: the slice payload
      // carries the whole drawing, deleteSelection cuts it.
      items.push({ text: t("context.cut", this), event: "cut" });
      items.push({ text: t("context.copy", this), event: "copy" });
      items.push({ text: t("context.paste", this), event: "paste" });
      items.push({ text: "-" });
      items.push({
        text: t("ribbon.cmd.wrap", this) || "Wrap Text",
        items: [
          {
            text: t("ribbon.opt.wrap-inline", this) || "In Line with Text",
            event: "wrap",
            value: "inline",
          },
          { text: t("ribbon.opt.wrap-square", this) || "Square", event: "wrap", value: "square" },
          { text: t("ribbon.opt.wrap-tight", this) || "Tight", event: "wrap", value: "tight" },
          {
            text: t("ribbon.opt.wrap-behind", this) || "Behind Text",
            event: "wrap",
            value: "behind",
          },
          {
            text: t("ribbon.opt.wrap-front", this) || "In Front of Text",
            event: "wrap",
            value: "front",
          },
          {
            text: t("ribbon.opt.wrap-top-bottom", this) || "Top and Bottom",
            event: "wrap",
            value: "top-bottom",
          },
        ],
      });
      items.push({ text: "-" });
      items.push({ text: t("context.bring-forward", this), event: "bring-forward" });
      items.push({ text: t("context.send-backward", this), event: "send-backward" });
      items.push({ text: t("context.bring-to-front", this), event: "bring-to-front" });
      items.push({ text: t("context.send-to-back", this), event: "send-to-back" });
      // The numeric layout dialog needs an offset-anchored floating drawing —
      // an inline picture has no offset to edit.
      items.push({
        text: t("context.size-position", this),
        event: "drawing-properties",
        ...(this.#drawingStateOf() ? {} : { disabled: true }),
      });
      // Crop edits the picture's source — a shape has no source of its own.
      items.push({
        text: t("context.crop", this),
        event: "drawing-crop",
        ...(this.#selectedImage() ? {} : { disabled: true }),
      });
      // Reset Crop — only meaningful when the picture actually carries one.
      items.push({
        text: t("context.crop-reset", this),
        event: "drawing-crop-reset",
        ...(this.#selectedImageHasCrop() ? {} : { disabled: true }),
      });
      items.push({ text: "-" });
      items.push({ text: t("context.delete-picture", this), event: "delete-picture" });
      menu.setAttribute("items", JSON.stringify(items));
      return;
    }
    const { selection } = editor.state;
    const pos = this.#bridge?.posAtClient(event.clientX, event.clientY) ?? null;
    const inSelection =
      !selection.empty && pos != null && pos >= selection.from && pos <= selection.to;
    const onLink =
      pos != null &&
      editor.state.doc
        .resolve(Math.min(Math.max(pos, 0), editor.state.doc.content.size))
        .marks()
        .some((m) => m.type.name === "link");
    // The click sits inside a table when any ancestor (nearest wins for the
    // command) is a table node — Word then carries table entries on the menu.
    const inTable =
      pos != null &&
      (() => {
        const $p = editor.state.doc.resolve(
          Math.min(Math.max(pos, 0), editor.state.doc.content.size),
        );
        for (let d = $p.depth; d > 0; d -= 1) {
          if ($p.node(d).type === editor.state.schema.nodes.table) return true;
        }
        return false;
      })();
    // Word: a right-click outside the selection collapses the caret there.
    if (pos != null && !inSelection) editor.commands.setTextSelection(pos);
    // Word's spelling menu: a right-click on a flagged word leads with its
    // suggestions and the ignore levels. activateAt marks it active so the
    // shared replace/ignore commands below act on this occurrence.
    const spellingHit = pos != null ? this.#spelling.activateAt(pos) : null;
    const grammarHit = !spellingHit && pos != null ? this.#spelling.activateGrammarAt(pos) : null;
    const items: RibbonMenuItem[] = [];
    if (spellingHit) {
      const suggestions = spellSuggestions(spellingHit.word, 5, spellingHit.lang);
      if (suggestions.length) {
        for (const suggestion of suggestions) {
          items.push({ text: suggestion, event: "spell-pick", value: suggestion });
        }
      } else {
        items.push({ text: t("spelling.no-suggestions", this), disabled: true });
      }
      items.push({ text: "-" });
      items.push({ text: t("spelling.ignore-once", this), event: "spell-ignore-once" });
      items.push({ text: t("spelling.ignore-all", this), event: "spell-ignore-all" });
      items.push({ text: t("spelling.add", this), event: "spell-add" });
      items.push({ text: t("pane.spelling", this) || "Spelling...", event: "spell-check" });
      items.push({ text: "-" });
    } else if (grammarHit) {
      if (grammarHit.replacements.length) {
        for (const rep of grammarHit.replacements) {
          items.push({ text: rep, event: "spell-pick", value: rep });
        }
      }
      items.push({ text: "-" });
      items.push({ text: t("spelling.ignore-once", this), event: "grammar-ignore-once" });
      items.push({ text: "-" });
    }

    let wordUnderCursor = "";
    if (inSelection) {
      const selText = editor.state.doc.textBetween(selection.from, selection.to, " ").trim();
      if (/^[A-Za-z\u00C0-\u024F\u1E00-\u1EFF]+$/.test(selText)) {
        wordUnderCursor = selText;
      }
    } else if (pos != null) {
      const $pos = editor.state.doc.resolve(Math.min(pos, editor.state.doc.content.size));
      const text = $pos.parent.textBetween(0, $pos.parent.content.size, " ");
      const offset = $pos.parentOffset;
      const re = /\b([A-Za-z\u00C0-\u024F\u1E00-\u1EFF]+)\b/g;
      let m: RegExpExecArray | null;
      while ((m = re.exec(text)) !== null) {
        if (m.index <= offset && offset <= m.index + m[0].length) {
          wordUnderCursor = m[1];
          break;
        }
      }
    }
    if (wordUnderCursor) {
      const synonyms = getSynonyms(wordUnderCursor, this.#caretLanguage().value, 4);
      if (synonyms.length > 0) {
        items.push({ text: `${t("context.synonyms", this)}:`, disabled: true });
        for (const syn of synonyms) {
          items.push({ text: `  ${syn}`, event: "thesaurus-replace", value: syn });
        }
      }
      items.push({
        text: `${t("pane.thesaurus", this)}...`,
        event: "thesaurus",
        value: wordUnderCursor,
      });
      items.push({ text: "-" });
    }
    if (inSelection) {
      items.push({ text: t("context.cut", this), event: "cut" });
      items.push({ text: t("context.copy", this), event: "copy" });
    }
    items.push({ text: t("context.paste", this), event: "paste" });
    items.push({
      text: t("context.keep-text-only", this),
      event: "paste",
      value: "keep-text-only",
    });
    items.push({ text: "-" });
    items.push({ text: t("context.font", this) || "Font...", event: "font-dialog" });
    items.push({ text: t("context.paragraph", this) || "Paragraph...", event: "paragraph-dialog" });
    items.push({
      text: t("context.styles", this) || "Styles",
      items: [
        { text: t("styles.pane.title", this) || "Styles Pane...", event: "styles-pane" },
        { text: "-" },
        { text: "Normal", event: "style", value: "Normal" },
        { text: "Heading 1", event: "style", value: "Heading 1" },
        { text: "Heading 2", event: "style", value: "Heading 2" },
        { text: "Heading 3", event: "style", value: "Heading 3" },
      ],
    });
    items.push({
      text: t("context.bullets-numbering", this) || "Bullets & Numbering",
      items: [
        { text: t("ribbon.cmd.bullet-list", this) || "Bullet List", event: "bullet-list" },
        { text: t("ribbon.cmd.ordered-list", this) || "Numbered List", event: "ordered-list" },
        {
          text: t("ribbon.cmd.multilevel-list", this) || "Multilevel List",
          event: "multilevel-list",
        },
      ],
    });
    items.push({ text: "-" });
    if (onLink) {
      items.push({ text: t("context.open-link", this), event: "open-link" });
      items.push({ text: t("context.copy-link", this), event: "copy-link" });
      items.push({ text: t("context.edit-link", this), event: "link" });
      items.push({ text: t("context.unlink", this), event: "unset-link" });
      items.push({ text: "-" });
    } else {
      items.push({ text: t("context.link", this) || "Link…", event: "link" });
    }
    items.push({ text: t("context.translate", this) || "Translate", event: "translate" });
    items.push({ text: t("context.comment", this), event: "new-comment" });
    items.push({ text: "-" });
    // Word: a right-click on a note reference offers edit/delete for the
    // referenced note's body (the caret was collapsed onto the atom above).
    const note = this.#dialogs.noteTarget();
    if (note) {
      const noun = note.kind === "endnote" ? "endnote" : "footnote";
      items.push({ text: t(`context.edit-${noun}`, this), event: "edit-note" });
      items.push({ text: t(`context.delete-${noun}`, this), event: "delete-note" });
      items.push({ text: "-" });
    }
    // A right-click on a field atom offers update (Word's F9) / edit; a form
    // checkbox offers the flip instead (nothing to update).
    const field = this.#dialogs.fieldTarget();
    if (field) {
      if (field.ref.kind === "formField") {
        items.push({
          text: t(field.ref.checked ? "context.uncheck" : "context.check", this),
          event: "toggle-field-checkbox",
        });
      } else {
        items.push({ text: t("context.update-field", this), event: "update-field" });
        items.push({ text: t("context.edit-field", this), event: "edit-field" });
        items.push({ text: t("context.update-all-fields", this), event: "update-all-fields" });
      }
      items.push({ text: "-" });
    }
    const ole = this.#selectedOleObject();
    if (ole) {
      const isExcel =
        ole.fileName.endsWith(".xlsx") ||
        ole.fileName.endsWith(".xls") ||
        ole.progId.toLowerCase().includes("excel");
      items.push({
        text: isExcel ? t("context.open-worksheet", this) : t("context.open-object", this),
        event: "open-embedded-object",
      });
      items.push({
        text: isExcel ? t("context.download-worksheet", this) : t("context.download-object", this),
        event: "download-embedded-object",
      });
      items.push({ text: "-" });
    }
    items.push({ text: t("context.select-all", this), event: "select" });
    if (inTable) {
      items.push({ text: "-" });
      items.push({
        text: t("context.insert", this),
        items: [
          { text: t("ribbon.cmd.insert-column-left", this), event: "insert-column-left" },
          { text: t("ribbon.cmd.insert-column-right", this), event: "insert-column-right" },
          { text: "-" },
          { text: t("ribbon.cmd.insert-row-above", this), event: "insert-row-above" },
          { text: t("ribbon.cmd.insert-row-below", this), event: "insert-row-below" },
          { text: "-" },
          { text: t("context.insert-cells", this), event: "insert-cells" },
        ],
      });
      items.push({
        text: t("context.delete", this),
        items: [
          { text: t("context.delete-cells", this), event: "delete-cells" },
          { text: t("ribbon.cmd.delete-column", this), event: "delete-column" },
          { text: t("ribbon.cmd.delete-row", this), event: "delete-row" },
          { text: "-" },
          { text: t("context.delete-table", this), event: "delete-table" },
        ],
      });
      items.push({
        text: t("context.select", this),
        items: [
          { text: t("ribbon.cmd.select-table-cell", this), event: "select-table-cell" },
          { text: t("ribbon.cmd.select-table-column", this), event: "select-table-column" },
          { text: t("ribbon.cmd.select-table-row", this), event: "select-table-row" },
          { text: t("ribbon.cmd.select-table", this), event: "select-table" },
        ],
      });
      items.push({ text: "-" });
      items.push({ text: t("ribbon.cmd.merge-cells", this), event: "merge-cells" });
      items.push({ text: t("ribbon.cmd.split-cell", this), event: "split-cell" });
      items.push({ text: t("ribbon.cmd.split-table", this), event: "split-table" });
      items.push({ text: "-" });
      items.push({
        text: t("ribbon.cmd.autofit", this),
        items: [
          { text: t("ribbon.opt.autofit-contents", this), event: "autofit-contents" },
          { text: t("ribbon.opt.autofit-window", this), event: "autofit-window" },
          { text: t("ribbon.opt.fixed-column-width", this), event: "fixed-column-width" },
        ],
      });
      items.push({
        text: t("ribbon.cmd.align-cell", this),
        items: [
          { text: t("ribbon.opt.cell-align-tl", this), event: "align-cell", value: "tl" },
          { text: t("ribbon.opt.cell-align-tc", this), event: "align-cell", value: "tc" },
          { text: t("ribbon.opt.cell-align-tr", this), event: "align-cell", value: "tr" },
          { text: t("ribbon.opt.cell-align-ml", this), event: "align-cell", value: "ml" },
          { text: t("ribbon.opt.cell-align-mc", this), event: "align-cell", value: "mc" },
          { text: t("ribbon.opt.cell-align-mr", this), event: "align-cell", value: "mr" },
          { text: t("ribbon.opt.cell-align-bl", this), event: "align-cell", value: "bl" },
          { text: t("ribbon.opt.cell-align-bc", this), event: "align-cell", value: "bc" },
          { text: t("ribbon.opt.cell-align-br", this), event: "align-cell", value: "br" },
        ],
      });
      items.push({
        text: t("context.distribute", this),
        items: [
          { text: t("ribbon.cmd.distribute-rows", this), event: "distribute-rows" },
          { text: t("ribbon.cmd.distribute-columns", this), event: "distribute-columns" },
        ],
      });
      items.push({ text: t("ribbon.cmd.text-direction", this), event: "text-direction" });
      items.push({
        text: t("ribbon.cmd.cell-margins", this),
        items: [
          {
            text: t("ribbon.opt.cell-margin-normal", this),
            event: "cell-margins",
            value: "normal",
          },
          { text: t("ribbon.opt.cell-margin-none", this), event: "cell-margins", value: "none" },
        ],
      });
      items.push({
        text: t("ribbon.opt.borders-shading", this),
        event: "border",
        value: "borders-shading",
      });
      items.push({ text: t("ribbon.cmd.repeat-header-rows", this), event: "repeat-header-rows" });
      items.push({ text: t("ribbon.group.sort", this) || "Sort...", event: "sort" });
      items.push({ text: t("context.formula", this), event: "table-formula" });
      items.push({ text: "-" });
      items.push({ text: t("context.table-properties", this), event: "table-properties" });
    }
    menu.setAttribute("items", JSON.stringify(items));
  };

  #insertShapeAt(
    preset: string | undefined,
    rect?: {
      page: number;
      x: number;
      y: number;
      w: number;
      h: number;
      flipH?: boolean;
      flipV?: boolean;
    },
  ): void {
    this.#insert.insertShapeAt(preset, rect);
  }

  #insertWordArt(): void {
    this.#insert.insertWordArt();
  }

  #insertBlankPage(): void {
    this.#insert.insertBlankPage();
  }

  #insertCoverPage(): void {
    this.#insert.insertCoverPage();
  }

  #insertDateTime(detail: { text: string; instruction?: string }): void {
    this.#insert.insertDateTime(detail);
  }

  #insertCustomToc(detail: {
    headingRange: string;
    leader: string;
    showPageNumbers: boolean;
    alignPageNumbers: boolean;
    hyperlink?: boolean;
    styles?: string;
  }): void {
    this.#insert.insertCustomToc(detail);
  }

  #insertFileText(): void {
    this.#insert.insertFileText();
  }

  /** Event → handler tables for the extracted host-command domains. Built on
   *  first dispatch (the adapter closures read live element state), then
   *  cached. Each domain receives only the narrow view its bodies call. */
  #hostRegistry?: HostCommandRegistry;
  #sdt?: SdtCommands;

  #sdtCommand(): SdtCommands {
    return (this.#sdt ??= new SdtCommands({
      editor: () => this.editor,
      bridge: () =>
        this.#bridge
          ? {
              activeEditor: () => this.#bridge!.activeEditor(),
              focus: () => this.#bridge!.focus(),
            }
          : undefined,
      element: () => this as HTMLElement,
      rerender: () => {
        this.#bridge?.replaceOverlays();
      },
    }));
  }

  /** Quick Parts domain — its dialog commits (save/rename/delete) arrive as
   *  element events and route back into this instance. */
  #buildingBlocks?: BuildingBlocksHostCommands;

  #hostCommandRegistry(): HostCommandRegistry {
    this.#buildingBlocks ??= new BuildingBlocksHostCommands(this.#buildingBlocksView());
    return (this.#hostRegistry ??= hostCommands(
      {
        navigation: {
          editor: () => this.editor,
          togglePane: (id) => this.#togglePane(id),
          goToPage: () => this.#goToPage(),
          openGoToDialog: (kind) => this.#openGoToDialog(kind as GoToKind | undefined),
          openPropertiesDialog: () => this.#openPropertiesDialog(),
          openSearch: () => this.#navigation.openSearch(),
          openFindReplace: () => this.#navigation.openFindReplace(),
          zoom: () => this.#status.getZoom(),
          setZoom: (pct) => this.#setZoom(pct),
          showZoomDialog: () => this.#status.showZoomDialog(),
          zoomPreset: (preset) => this.#status.zoomPreset(preset),
          docProtected: () => this.#docProtected,
          syncEditModeMenu: () => this.#syncEditModeMenu(),
          setShowMarks: (on) => this.setShowMarks(on),
          getShowMarks: () => this.getShowMarks(),
          showRuler: () => this.getShowRuler(),
          setShowRuler: (on) => this.setShowRuler(on),
          showGridlines: () => this.#stage?.showGridlines ?? false,
          setShowGridlines: (on) => this.#stage?.setShowGridlines(on),
          setView: (view) => this.setAttribute("view", view),
          setUiDirection: (dir) => this.setUiDirection(dir),
          toggleSplitWindow: () => this.#toggleSplitWindow(),
          toggleFocusMode: () => this.#toggleFocusMode(),
        },
        sections: {
          openPageSetup: () => this.#sections.openPageSetup(),
          setPageSize: (value) => this.#sections.setPageSize(value),
          setOrientation: (value) => this.#sections.setOrientation(value),
          setMargins: (value) => this.#sections.setMargins(value),
          openColumnsDialog: () => this.#sections.openColumnsDialog(),
          setColumnCount: (count) => this.#sections.setColumnCount(count),
          openLineNumbersOptions: () => this.#sections.openLineNumbersOptions(),
          setLineNumbers: (mode) => this.#sections.setLineNumbers(mode),
          openBordersDialog: (tab) => this.#sections.openBordersDialog(tab),
          setPageBorders: (preset) => this.#sections.setPageBorders(preset),
          insertCoverPage: () => this.#insertCoverPage(),
          insertBlankPage: () => this.#insertBlankPage(),
          setHyphenation: (mode) => this.#setHyphenation(mode),
          openHyphenationOptions: () => this.#openHyphenationOptions(),
          insertSoftHyphen: () => this.#insertSoftHyphen(),
        },
        references: {
          editor: () => this.editor,
          bridge: () => this.#bridge,
          flow: () => this.#flow,
          element: () => this as HTMLElement,
          openNoteSettings: () => this.#openNoteSettings(),
          markIndexEntry: (target) => this.#references.markIndexEntry(target),
          markCitation: (target) => this.#references.markCitation(target),
          setCitationStyle: (style) => this.#references.setCitationStyle(style),
          insertBibliography: () => this.#references.insertBibliography(),
          bibliographySources: () => this.#references.bibliographySources(),
          crossReferenceTargets: () => this.#dialogs.crossReferenceTargets(),
          noteInsert: (kind) => this.#dialogs.noteInsert(kind),
          noteEditAtSelection: () => this.#dialogs.noteEditAtSelection(),
          noteDeleteAtSelection: () => this.#dialogs.noteDeleteAtSelection(),
          jumpNextNote: () => this.#jumpNextNote(),
          jumpPreviousNote: () => this.#jumpPreviousNote(),
          openBookmarkDialog: () => this.#openBookmarkDialog(),
          insertBookmark: () => this.#insertBookmark(),
        },
        mailMerge: {
          element: () => this as HTMLElement,
          recipients: () => this.#merge.recipients(),
          insertAddressBlock: () => this.#merge.insertAddressBlock(),
          insertGreetingLine: () => this.#merge.insertGreetingLine(),
          togglePreview: () => this.#merge.togglePreview(),
          firstRecord: () => this.#merge.firstRecord(),
          lastRecord: () => this.#merge.lastRecord(),
          setMergeType: (type) => this.#merge.setMergeType(type),
          finishMerge: (mode) => void this.#finishMerge(mode),
        },
        comments: {
          insertComment: () => this.#comments.insertComment(),
          editComment: () => this.#comments.editComment(),
          deleteComment: () => this.#comments.deleteComment(),
          jumpComment: (direction) => this.#comments.jumpComment(direction),
          togglePane: (id) => this.#togglePane(id),
          setTaskpane: (id, open) => this.#setTaskpane(id, open),
          getTaskpaneState: (id) => this.getTaskpaneState(id),
        },
        revisions: {
          editor: () => this.editor,
          togglePane: (id) => this.#togglePane(id),
          setMarkupView: (view) => {
            this.#markupView = view;
          },
          getMarkupAuthors: () => this.#markupAuthors,
          setMarkupAuthors: (authors) => {
            this.#markupAuthors = authors;
          },
          setMarkupColors: (colors) => {
            this.#markupColors = colors;
          },
          setBalloons: (mode) => {
            this.#balloons = mode;
          },
          renderDoc: (doc) => this.#renderDoc(doc),
          syncMarkupMenus: () => this.#syncMarkupMenus(),
          getJSON: () => this.getJSON(),
        },
        proofing: {
          editor: () => this.editor,
          showWordCount: () => this.#status.showWordCount(),
          spellingRun: () => this.#spelling.run(),
          setTaskpane: (id, open) => this.#setTaskpane(id, open),
          spellingIssues: () => this.#spelling.issues(),
          spellingGoto: (index) => this.#spelling.goto(index),
          spellingReplace: (replacement) => this.#spelling.replace(replacement),
          spellingIgnore: (mode) => this.#spelling.ignore(mode),
          openLanguageDialog: () => this.#onLanguageOpen(),
          openThesaurus: (word?: string) => this.#openThesaurus(word),
          readAloud: () => this.#toggleReadAloud(),
        },
        fields: {
          fieldInsert: () => this.#dialogs.fieldInsert(),
          fieldUpdateAtSelection: () => this.#dialogs.fieldUpdateAtSelection(),
          updateAllFields: () => this.#dialogs.updateAllFields(),
          fieldEditAtSelection: () => this.#dialogs.fieldEditAtSelection(),
          fieldToggleCheckboxAtSelection: () => this.#dialogs.fieldToggleCheckboxAtSelection(),
          toggleFieldCodes: () => this.toggleFieldCodes(),
          insertEquation: (template) => this.#insertEquation(template),
          insertEquationSymbol: (char) => this.#insertEquationSymbol(char),
        },
        clipboard: {
          editor: () => this.editor,
          activeEditor: () => this.#bridge?.activeEditor() ?? this.editor,
          element: () => this as HTMLElement,
          copySelection: (cut) => this.#bridge?.copySelection(cut),
          paste: (textOnly) => this.#clipboard.paste(textOnly),
          togglePane: (id) => this.#togglePane(id),
          showTaskpane: (id) => this.showTaskpane(id),
          renderStylesPane: () => this.#renderStylesPane(),
          toggleMarkdownInput: () => {
            this.#markdown = !this.#markdown;
          },
          syncFormatButtons: () => this.#syncFormatButtons(),
          insertLink: () => this.#insertLink(),
          hrefAtCaret: () => this.#hrefAtCaret(),
          jumpToBookmark: (name) => this.#jumpToBookmark(name),
          select: (value) => this.#select(value),
          toggleFormatPainter: () => this.#toggleFormatPainter(),
        },
        drawing: {
          editor: () => this.editor,
          activeEditor: () => this.#bridge?.activeEditor() ?? this.editor,
          element: () => this as HTMLElement,
          showCompressPictures: () => this.#showCompressPictures(),
          armTransparentPick: () => this.#armTransparentPick(),
          drawingMulti: () => this.#bridge?.drawingMulti(),
          pickImage: () => this.#imageInput?.click(),
          pickPicture: () => this.#pictureInput?.click(),
          focusBridge: () => this.#bridge?.focus(),
          drawingState: () => this.#drawingStateOf(),
          enterCropMode: () => {
            this.#bridge?.enterCropMode();
          },
          enterEditPointsMode: () => {
            this.#bridge?.enterEditPointsMode();
          },
          insertShapeAt: (preset) => this.#insertShapeAt(preset),
          armShapeDrawer: (preset) => this.#armShapeDrawer(preset),
          insertWordArt: () => this.#insertWordArt(),
        },
        tables: {
          element: () => this as HTMLElement,
          editor: () => this.editor,
          activeEditor: () => this.#bridge?.activeEditor() ?? this.editor,
          contentWidthPx: () => this.#flow?.contentWidthPx,
          setPenStyle: (style) => {
            this.#pen = { ...this.#pen, style };
          },
          setPenSize: (size) => {
            this.#pen = { ...this.#pen, size };
          },
          setPenColor: (color) => {
            this.#pen = { ...this.#pen, color };
          },
          borderPainting: () => this.#borderPainting,
          borderErase: () => this.#borderErase,
          stopBorderPainting: () => this.#stopBorderPainting(),
          armBorderPainter: (erase) => this.#armBorderPainter(erase),
          tableDrawing: () => this.#tableDrawing,
          tableEraser: () => this.#tableEraser,
          toggleDrawTable: () => {
            if (this.#tableDrawing) this.#stopTableDrawing();
            else this.#armTableDrawer();
          },
          toggleTableEraser: () => {
            if (this.#tableEraser) this.#stopTableEraser();
            else this.#armTableEraser();
          },
        },
        dialogs: {
          element: () => this as HTMLElement,
          activeEditor: () => this.#bridge?.activeEditor() ?? this.editor,
          docStyles: (editor) => this.#docStyles(editor),
          runState: (state) => this.#runStateOf(state),
          chartEditAtSelection: () => this.#dialogs.chartEditAtSelection(),
          chartTypeAtSelection: () => this.#dialogs.chartTypeAtSelection(),
          phoneticOpen: () => this.#dialogs.phoneticOpen(),
          twoInOneOpen: () => this.#dialogs.twoInOneOpen(),
          defineListOpen: () => this.#dialogs.defineListOpen(),
        },
        fileIo: {
          emitCancelable: (name) => this.#emitCancelable(name),
          saveAs: () => this.#saveAs(),
          pickFile: () => this.#pickFile(),
          print: () => this.#print(),
          insertFileText: () => this.#insertFileText(),
        },
        headerFooter: {
          editor: () => this.editor,
          bridge: () => this.#bridge,
          activeEditor: () => this.#bridge?.activeEditor() ?? this.editor,
          storyPage: () => this.#stories.page(),
          toggleSectionFlag: (flag) => this.#sections.toggleSectionFlag(flag),
          removeStory: (kind) => this.#removeStory(kind),
          removePageNumbers: () => this.#removePageNumbers(),
          openPageNumberFormat: () => this.#sections.openPageNumberFormat(),
        },
        design: {
          setPageColor: (value) => this.#design.setPageColor(value),
          setParagraphSpacing: (preset) => this.#design.setParagraphSpacing(preset),
          openWatermarkDialog: () => this.#design.openWatermarkDialog(),
          setWatermark: (preset) => this.#design.setWatermark(preset),
          openFillEffectsDialog: () => this.#design.openFillEffectsDialog(),
          restoreStylesSnapshot: () => this.#restoreStylesSnapshot(),
        },
      },
      [
        this.#buildingBlocks,
        new SdtHostCommands(this.#sdtCommand(), () => this.#openSdtPropertiesDialog()),
      ],
    ));
  }

  /** The Quick Parts domain's narrow view — the block list lives in the MAIN
   *  document's attrs (documentExtras.docenBlocks/glossary) even while a
   *  furniture story is being edited; insertion targets the active editor so a
   *  story caret receives the content. Mutations are single doc-attr
   *  transactions; insertion routes the Tiptap command. */
  #buildingBlocksView(): BuildingBlocksHostView {
    const active = (): Editor | null | undefined => this.#bridge?.activeEditor() ?? this.editor;
    return {
      element: () => this as HTMLElement,
      editable: () => active()?.isEditable === true,
      blocks: () => (this.editor ? blocksOfDocAttrs(this.editor.state.doc.attrs) : []),
      setBlocks: (blocks) => {
        const editor = this.editor;
        if (!editor?.isEditable) return;
        editor.commands.command(({ state, dispatch }) => {
          dispatch?.(
            state.tr.setDocAttribute(
              "documentExtras",
              withBlocks(state.doc.attrs.documentExtras, blocks),
            ),
          );
          return true;
        });
      },
      insertBlock: (id) => {
        const editor = active();
        if (!editor?.isEditable) return;
        editor.commands["insert-building-block"](id);
        this.#bridge?.focus();
      },
      selectionSlice: () => {
        const editor = active();
        if (!editor) return null;
        const slice = parseSlicePayload(selectionSlicePayload(editor.state));
        if (!slice) return null;
        const { from, to } = editor.state.selection;
        const text = editor.state.doc.textBetween(from, to, "\n", "\n").trim();
        const firstLine = text.split("\n")[0] ?? "";
        return {
          slice,
          preview: text.slice(0, 200),
          suggestedName: firstLine.slice(0, 64).trim(),
        };
      },
      focusBridge: () => this.#bridge?.focus(),
    };
  }

  #splitWindow = false;

  #toggleSplitWindow(): void {
    this.#splitWindow = !this.#splitWindow;
    const docArea = this.shadowRoot?.querySelector("docen-document-area");
    if (docArea) {
      if (this.#splitWindow) docArea.setAttribute("split", "");
      else docArea.removeAttribute("split");
    }
  }

  #toggleFocusMode(): void {
    const currentView = this.getAttribute("view");
    this.setAttribute("view", currentView === "read" ? "print" : "read");
  }

  #applyDocumentTheme(kind: string, value?: string, persist = true): void {
    const themeId = value?.toLowerCase() || "office";
    const themeDef = THEMES[themeId] ?? THEMES.office;
    if (!themeDef) return;

    const targets = [this, this.shadowRoot?.querySelector("docen-workspace")].filter(
      Boolean,
    ) as HTMLElement[];
    for (const target of targets) {
      if (kind === "theme" || kind === "theme-color") {
        target.style.setProperty("--docen-theme-accent1", `#${themeDef.colors.accent1}`);
        target.style.setProperty("--docen-theme-accent2", `#${themeDef.colors.accent2}`);
        target.style.setProperty("--docen-theme-accent3", `#${themeDef.colors.accent3}`);
        target.style.setProperty("--docen-theme-accent4", `#${themeDef.colors.accent4}`);
        target.style.setProperty("--docen-theme-accent5", `#${themeDef.colors.accent5}`);
        target.style.setProperty("--docen-theme-accent6", `#${themeDef.colors.accent6}`);
      }
      if (kind === "theme" || kind === "theme-font") {
        target.style.setProperty("--docen-theme-font-major", themeDef.fonts.majorFont);
        target.style.setProperty("--docen-theme-font-minor", themeDef.fonts.minorFont);
      }
    }
    if (persist && this.editor) {
      const attrs = (this.editor.state.doc.attrs ?? {}) as {
        documentExtras?: Record<string, unknown>;
      };
      const extras = attrs.documentExtras ?? {};
      const settings = (extras.settings as Record<string, unknown>) ?? {};
      this.editor.view.dispatch(
        this.editor.state.tr.setDocAttribute("documentExtras", {
          ...extras,
          settings: { ...settings, theme: { id: themeId, kind } },
        }),
      );
    }
    this.#bridge?.replaceOverlays();
  }

  #openVersionHistory(): void {
    const dialog = this.shadowRoot?.querySelector("docen-version-history-dialog") as {
      show(versions: any[]): void;
    } | null;
    dialog?.show(this.#versionSnapshots);
  }

  #toggleReadAloud(): void {
    if (this.#readAloud.isPlaying()) {
      this.#readAloud.stop();
      return;
    }
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return;
    const { $from, empty } = editor.state.selection;
    const text = empty
      ? $from.parent.textBetween(0, $from.parent.content.size, " ")
      : editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, " ");
    if (text.trim()) {
      this.#readAloud.speak(text.trim());
    }
  }

  #openSdtPropertiesDialog(): void {
    const dialog = this.shadowRoot?.querySelector("docen-sdt-dialog") as {
      show(initial?: SdtPropertiesValues): void;
    } | null;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    let initialProps: SdtPropertiesValues | undefined;
    if (editor) {
      const { $from } = editor.state.selection;
      for (let d = $from.depth; d > 0; d--) {
        const node = $from.node(d);
        if (node.attrs?.properties) {
          const p = node.attrs.properties as Record<string, unknown>;
          initialProps = {
            title: (p.title as string) ?? "",
            tag: (p.tag as string) ?? "",
            cannotDelete: Boolean(p.cannotDelete),
            cannotEdit: Boolean(p.cannotEdit),
          };
          break;
        }
      }
    }
    dialog?.show(initialProps);
  }

  #updateRevealFormatting(): void {
    const pane = this.shadowRoot?.querySelector("docen-reveal-formatting-pane") as {
      setFormatting?(info: FormattingInfo): void;
    } | null;
    if (!pane || !this.getTaskpaneState("reveal")) return;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return;
    pane.setFormatting?.(formattingInfoOf(editor));
  }

  #syncAltTextPane(): void {
    const pane = this.shadowRoot?.querySelector("docen-alt-text-pane") as {
      setTarget?(target: { kind?: string; title?: string; descr?: string } | null): void;
    } | null;
    if (!pane || !this.getTaskpaneState("altText")) return;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) {
      pane.setTarget?.(null);
      return;
    }
    const sel = editor.state.selection;
    if (sel instanceof NodeSelection) {
      const node = sel.node;
      const attrs = node.attrs as Record<string, unknown>;
      const kind = node.type.name;
      if (kind === "model3d" || kind === "ink") {
        pane.setTarget?.({
          kind,
          title: (attrs.title as string) ?? "",
          descr: (attrs.descr as string) ?? "",
        });
        return;
      }
      if (kind === "image") {
        pane.setTarget?.({
          kind,
          title: (attrs.name as string) ?? "",
          descr: (attrs.title as string) ?? "",
        });
        return;
      }
      if (kind === "wpsShape" || kind === "wpgGroup" || kind === "chart") {
        const payload = (attrs[kind] ?? {}) as Record<string, unknown>;
        pane.setTarget?.({
          kind,
          title: (payload.title as string) ?? "",
          descr: (payload.descr as string) ?? "",
        });
        return;
      }
    }
    pane.setTarget?.(null);
  }

  readonly #onRevealCompareToggle = (event: CustomEvent<{ enabled?: boolean }>): void => {
    const pane = this.shadowRoot?.querySelector("docen-reveal-formatting-pane") as {
      setComparison?(reference: FormattingInfo | null): void;
    } | null;
    if (!pane?.setComparison) return;
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    pane.setComparison(event.detail?.enabled === true && editor ? formattingInfoOf(editor) : null);
  };

  readonly #onItemPreview = (event: CustomEvent<{ event?: string; value?: string }>): void => {
    const { event: name, value } = event.detail ?? {};
    if (!name || !value) return;

    const editor = this.editor;
    if (!editor) return;

    if (name === "theme" || name === "theme-color" || name === "theme-font") {
      if (!this.#previewSnapshot) {
        const vars = new Map<string, string>();
        const varNames = [
          "--docen-theme-accent1",
          "--docen-theme-accent2",
          "--docen-theme-accent3",
          "--docen-theme-accent4",
          "--docen-theme-accent5",
          "--docen-theme-accent6",
          "--docen-theme-font-major",
          "--docen-theme-font-minor",
        ];
        for (const v of varNames) {
          vars.set(v, this.style.getPropertyValue(v));
        }
        this.#previewSnapshot = {
          docContent: null,
          selection: { from: editor.state.selection.from, to: editor.state.selection.to },
          themeKind: name,
          themeCssVars: vars,
        };
      }
      this.#applyDocumentTheme(name, value, false);
      return;
    }

    if (name === "style" || name === "table-style" || name === "style-set") {
      if (!this.#previewSnapshot) {
        this.#previewSnapshot = {
          docContent: editor.state.doc.toJSON(),
          selection: { from: editor.state.selection.from, to: editor.state.selection.to },
        };
      } else {
        this.#revertPreviewDoc();
      }

      const target = this.#bridge?.activeEditor() ?? editor;
      const commands = target.commands as unknown as Record<string, (value?: string) => unknown>;
      const cmd = commands[name];
      if (typeof cmd === "function") {
        const origDispatch = target.view.dispatch.bind(target.view);
        target.view.dispatch = (tr) => {
          tr.setMeta("addToHistory", false);
          origDispatch(tr);
        };
        try {
          cmd(value);
        } finally {
          target.view.dispatch = origDispatch;
        }
      }
    }
  };

  #revertPreviewDoc(): void {
    if (!this.#previewSnapshot?.docContent || !this.editor) return;
    const editor = this.editor;
    try {
      const restoredNode = editor.schema.nodeFromJSON(
        this.#previewSnapshot.docContent as JSONContent,
      );
      const tr = editor.state.tr.replaceWith(
        0,
        editor.state.doc.content.size,
        restoredNode.content,
      );
      const targetPos = Math.min(this.#previewSnapshot.selection.from, tr.doc.content.size);
      tr.setSelection(TextSelection.near(tr.doc.resolve(targetPos)));
      tr.setMeta("addToHistory", false);
      editor.view.dispatch(tr);
    } catch {
      // safe fallback
    }
  }

  readonly #onItemPreviewEnd = (): void => {
    if (!this.#previewSnapshot) return;

    if (this.#previewSnapshot.themeCssVars) {
      const targets = [this, this.shadowRoot?.querySelector("docen-workspace")].filter(
        Boolean,
      ) as HTMLElement[];
      for (const target of targets) {
        for (const [key, val] of this.#previewSnapshot.themeCssVars) {
          if (val) target.style.setProperty(key, val);
          else target.style.removeProperty(key);
        }
      }
      this.#bridge?.replaceOverlays();
    }

    if (this.#previewSnapshot.docContent) {
      this.#revertPreviewDoc();
    }

    this.#previewSnapshot = null;
  };

  readonly #onCommand = (event: CustomEvent<{ event?: string; value?: string }>): void => {
    const { event: name, value } = event.detail ?? {};
    if (typeof name !== "string") return;
    if (this.#previewSnapshot) {
      this.#onItemPreviewEnd();
    }
    if (name === "toggle-ribbon-minimized") {
      this.#toggleRibbonMinimized();
      return;
    }
    if (name === "translate") {
      const editor = this.#bridge?.activeEditor() ?? this.editor;
      const selected =
        editor && !editor.state.selection.empty
          ? editor.state.doc.textBetween(
              editor.state.selection.from,
              editor.state.selection.to,
              " ",
            )
          : "";
      this.dispatchEvent(
        new CustomEvent("docen:translate", {
          bubbles: true,
          composed: true,
          detail: { text: selected },
        }),
      );
      this.#openTranslate(selected);
      return;
    }
    if (name === "theme" || name === "theme-color" || name === "theme-font") {
      this.#applyDocumentTheme(name, value);
      return;
    }
    if (name === "version-history") {
      this.#openVersionHistory();
      return;
    }
    if (name === "toggle-checkbox") {
      this.#sdtCommand().toggleCheckboxAtCaret();
      return;
    }
    if (name === "open-embedded-object" || name === "download-embedded-object") {
      const ole = this.#selectedOleObject();
      if (ole) {
        downloadOleObject(ole.data, ole.fileName);
        this.dispatchEvent(
          new CustomEvent("ole:open", { bubbles: true, composed: true, detail: ole }),
        );
      }
      return;
    }
    // Read-only documents (Viewing mode) reject document-changing commands —
    // the viewless editor has no DOM surface to refuse them, so the gate
    // lives here (Word's read-only ribbon). Chrome actions and clipboard
    // reads stay live.
    if (this.editor && !this.editor.isEditable && !READONLY_LIVE.has(name)) {
      if (
        this.#protectionMode === "comments" &&
        (name === "new-comment" ||
          name === "delete-comment" ||
          name === "resolve-comment" ||
          name === "reopen-comment" ||
          name === "show-comments" ||
          name.includes("comment"))
      ) {
        // allow comment actions in comments protection mode
      } else {
        return;
      }
    }
    // Forms or ReadOnly protection mode: allow only field / permission region interaction outside readonly live
    if (
      (this.#protectionMode === "forms" || this.#protectionMode === "readOnly") &&
      !READONLY_LIVE.has(name) &&
      !name.startsWith("sdt-") &&
      name !== "toggle-checkbox"
    ) {
      const active = this.#bridge?.activeEditor() ?? this.editor;
      if (active) {
        if (this.#protectionMode === "forms" && !isInsideEditableField(active)) {
          return;
        }
        if (this.#protectionMode === "readOnly" && !isInsideEditablePermission(active)) {
          return;
        }
      }
    }
    // Local host commands (chrome actions plus document actions the engine
    // can't express) route through the per-domain registry — chrome handlers
    // need no editor, the rest run once a document has opened. The wired
    // Tiptap commands / add-in commands below stay in this element.
    const host = this.#hostCommandRegistry();
    const chrome = host.chrome.get(name);
    if (chrome?.(value)) return;
    const editor = this.editor;
    if (!editor) return;
    const local = host.editor.get(name);
    if (local?.(value)) return;
    // Built-in commands route to editor.commands.<event>(value) —
    // DocumentCommands registers every ribbon event as a native Tiptap command.
    // A user add-in overrides one by contributing a Tiptap extension whose
    // addCommands redefines the same name (Tiptap's native override mechanism).
    // They target the editor input currently routes into — while a furniture
    // story is open that's the story's editor, not the main document (whose
    // selection is stale and would be stamped instead).
    const target = this.#bridge?.activeEditor() ?? editor;
    const commands = target.commands as unknown as Record<string, (value?: string) => unknown>;
    const cmd = commands[name];
    if (typeof cmd === "function") {
      // Word's measure boxes read a bare number in the locale's unit system —
      // "5" beside "4.39 厘米" means 5cm. The commands parse bare numbers as
      // twips, so qualify before dispatch; "0" (cell-height's auto) and
      // unit-suffixed text pass through untouched.
      let arg = value;
      if (
        typeof value === "string" &&
        (name === "drawing-width" ||
          name === "drawing-height" ||
          name === "cell-width" ||
          name === "cell-height")
      ) {
        const n = Number(value.trim());
        if (value.trim() !== "" && Number.isFinite(n) && n > 0) {
          const scope = this.shadowRoot?.querySelector("docen-workspace") ?? this;
          arg = `${value.trim()}${useCmUnits(scope) ? "cm" : "in"}`;
        }
      }
      cmd(arg);
      if (name === "next-change" || name === "previous-change") {
        this.#bridge?.scrollIntoView(target.state.selection.from);
      }
      // The comboboxes keep focus to filter their lists — after a pick, hand
      // the keyboard back to the document (buttons never take it: their
      // mousedown preventDefaults).
      if (name === "font-name" || name === "font-size" || name === "style") {
        this.#bridge?.focus();
      }
      return;
    }
    // Not a Tiptap command — route to the first add-in that declares it. This
    // covers non-Tiptap actions contributed by external add-ins (e.g. a Help
    // button that opens a URL) that Tiptap can't express.
    this.dispatchCommand(name, value);
  };

  /** Menu items and the auto-save switch carry their action in `data-event`. */
  readonly #onChange = (event: Event): void => {
    // The QAT customize menu toggles bar membership (data-qat), not commands.
    const qat = (event.target as HTMLElement)?.dataset?.qat;
    if (qat) {
      this.#toggleQat(qat);
      return;
    }
    const name = (event.target as HTMLElement)?.dataset?.event;
    if (!name) return;
    switch (name) {
      case "autosave": {
        // The switch itself carries the new state; persists + (re)schedules or
        // clears the content backup.
        this.#setAutosave((event.target as HTMLInputElement).checked === true);
        break;
      }
      case "version-history":
        this.#openVersionHistory();
        break;
      case "open":
        // Host can take over via docen:open (preventDefault); else open the
        // picker — #onFileChange auto-detects docx/md from the extension.
        if (!this.#emitCancelable("docen:open")) this.#pickFile();
        break;
      case "save-as": {
        const format = this.#docxVariant;
        if (!this.#emitCancelable("docen:save-as", { format })) void this.#saveAs(format);
        break;
      }
      case "save-as-template":
        if (!this.#emitCancelable("docen:save-as", { format: "dotx" })) void this.#saveAsTemplate();
        break;
      case "save-as-markdown":
        if (!this.#emitCancelable("docen:save-as", { format: "markdown" }))
          void this.#saveAs("markdown");
        break;
      case "save-as-rtf":
        if (!this.#emitCancelable("docen:save-as", { format: "rtf" })) void this.#saveAs("rtf");
        break;
      case "save-as-html":
        if (!this.#emitCancelable("docen:save-as", { format: "html" })) void this.#saveAs("html");
        break;
      case "save-as-txt":
        if (!this.#emitCancelable("docen:save-as", { format: "txt" })) void this.#saveAs("txt");
        break;
      case "save-as-odt":
        if (!this.#emitCancelable("docen:save-as", { format: "odt" })) void this.#saveAs("odt");
        break;
      case "save-as-pdf":
        if (!this.#emitCancelable("docen:save-as", { format: "pdf" })) void this.#saveAsPdf();
        break;
      case "print":
        if (!this.#emitCancelable("docen:print")) void this.#print();
        break;
      case "print-preview":
        void this.openPrintPreview();
        break;
      case "properties":
        // Word's File → Info: the document properties pane.
        this.showTaskpane("properties");
        break;
      case "inspect-document":
        this.#inspectDocument();
        break;
      case "share":
        void this.#share();
        break;
      case "close":
        this.#closeDocument();
        break;
      case "new":
        // No built-in "new" — always hand to the host (docen:new).
        this.#emitCancelable("docen:new");
        break;
      case "new-from-template":
        this.#openTemplateDialog();
        break;
      case "options": {
        // Filename menu → open the Options dialog (UI language + theme +
        // spell-as-you-type + Markdown input + the document settings).
        const optionsEl = this.shadowRoot?.querySelector("docen-options-dialog");
        if (optionsEl) {
          optionsEl.setAttribute("locale", this.lang || document.documentElement.lang || "zh-CN");
          optionsEl.setAttribute("theme", this.theme ?? "light");
          optionsEl.setAttribute("proofing", String(this.#spelling.enabled()));
          optionsEl.setAttribute("markdown", String(this.#markdown));
          // The Document section seeds from the settings.xml slice; the tab
          // stop converts twips → cm (round-trip of convertMillimetersToTwip).
          const s = this.#documentSettings();
          (optionsEl as unknown as { document?: unknown }).document = {
            defaultTabStop:
              typeof s.defaultTabStop === "number"
                ? Math.round((s.defaultTabStop / (1440 / 2.54)) * 100) / 100
                : undefined,
            updateFields: s.updateFields === true,
            updateFieldsBeforePrint: this.#updateFieldsBeforePrint,
            fieldShading: this.#fieldShading,
            protection: (s.documentProtection as { edit?: string } | undefined)?.edit ?? "none",
            compatVersion: (s.compatibility as { version?: number } | undefined)?.version ?? 15,
          };
          // General/User section — the store is the source of truth (the
          // `user` attribute only overrides the rendered header).
          (optionsEl as unknown as { identity?: IdentitySettings }).identity =
            getSettings().identity;
          (optionsEl as unknown as { show?: () => void }).show?.();
        }
        break;
      }
    }
  };

  /** Forward this host's `lang` attribute to the internal <docen-workspace>
   *  and notify locale observers. Called on connect and whenever `lang`
   *  mutates (via #langObserver). The workspace is the resolveLang scope, so
   *  forwarding is what makes <docen-document lang> reach child components
   *  across the shadow boundary. */
  #syncLang(): void {
    const workspace = this.shadowRoot?.querySelector("docen-workspace");
    const lang = this.lang;
    if (lang) workspace?.setAttribute("lang", lang);
    else workspace?.removeAttribute("lang");
    notifyLocaleChange();
  }

  /** The caret's proofing language (the textStyle mark's w:lang fields).
   *  Runs without an explicit mark show the default proofing language —
   *  Word mirrors the editing language implied by the UI locale here. */
  #caretLanguage(): { value: string; noProof: boolean } {
    const editor = this.editor;
    const mark = editor?.state.selection.$from.marks().find((m) => m.type.name === "textStyle");
    const language = mark?.attrs.language as { value?: string } | undefined;
    const fallback = (document.documentElement.lang || "en").startsWith("zh") ? "zh-CN" : "en-US";
    return { value: language?.value || fallback, noProof: mark?.attrs.noProof === true };
  }

  /** Status-bar language item / Review → Language — open the dialog prefilled
   *  from the caret's current proofing language. */
  readonly #onLanguageOpen = (): void => {
    const dialog = this.shadowRoot?.querySelector("docen-language-dialog") as unknown as {
      show(tag: string | null, noProof?: boolean): void;
    } | null;
    const { value, noProof } = this.#caretLanguage();
    dialog?.show(value || null, noProof);
  };

  /** Open the Thesaurus task pane with the given or selected/current word. */
  #openThesaurus(word?: string): void {
    const editor = this.editor;
    let target = word;
    if (!target && editor) {
      const sel = editor.state.selection;
      if (!sel.empty) {
        target = editor.state.doc.textBetween(sel.from, sel.to, " ").trim();
      } else {
        const $pos = sel.$from;
        const text = $pos.parent.textBetween(0, $pos.parent.content.size, " ");
        const offset = $pos.parentOffset;
        const re = /\b([A-Za-z\u00C0-\u024F\u1E00-\u1EFF]+)\b/g;
        let m: RegExpExecArray | null;
        while ((m = re.exec(text)) !== null) {
          if (m.index <= offset && offset <= m.index + m[0].length) {
            target = m[1];
            break;
          }
        }
      }
    }
    this.#setTaskpane("thesaurus", true);
    const pane = this.shadowRoot?.querySelector("docen-thesaurus-pane") as
      | (HTMLElement & { lookup: (w: string) => void; lang: string })
      | null;
    if (pane) {
      pane.lang = this.#caretLanguage().value;
      if (target) {
        pane.lookup(target);
      }
    }
  }

  /** Open the Translate task pane with optional initial selection text. */
  #openTranslate(text?: string): void {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    let target = text;
    if (target === undefined && editor) {
      const sel = editor.state.selection;
      if (!sel.empty) {
        target = editor.state.doc.textBetween(sel.from, sel.to, " ").trim();
      }
    }
    this.#setTaskpane("translate", true);
    const pane = this.shadowRoot?.querySelector("docen-translate-pane") as
      | (HTMLElement & { setSelectionText: (t: string) => void })
      | null;
    if (pane && typeof pane.setSelectionText === "function" && target !== undefined) {
      pane.setSelectionText(target);
    }
  }

  /** Translates paragraph blocks in the document. */
  #translateDocument(from = "auto", to = "id"): void {
    const editor = this.#bridge?.activeEditor() ?? this.editor;
    if (!editor) return;

    const blocks: Array<{ from: number; to: number; text: string }> = [];
    editor.state.doc.descendants((node, pos) => {
      if (node.isTextblock && node.textContent.trim().length > 0) {
        blocks.push({
          from: pos + 1,
          to: pos + node.nodeSize - 1,
          text: node.textContent,
        });
        return false;
      }
      return true;
    });

    if (blocks.length === 0) return;

    const tr = editor.state.tr;
    for (let i = blocks.length - 1; i >= 0; i--) {
      const b = blocks[i];
      const translated = translateText(b.text, from, to);
      tr.insertText(translated, b.from, b.to);
    }
    editor.view.dispatch(tr);

    const pane = this.shadowRoot?.querySelector("docen-translate-pane") as
      | (HTMLElement & { onDocumentTranslated: (count: number) => void })
      | null;
    pane?.onDocumentTranslated?.(blocks.length);
  }

  readonly #onSelectionUpdateForTranslate = (): void => {
    if (this.#paneEl("translate")?.open && this.editor) {
      const sel = this.editor.state.selection;
      if (!sel.empty) {
        const text = this.editor.state.doc.textBetween(sel.from, sel.to, " ").trim();
        if (text) {
          const pane = this.shadowRoot?.querySelector("docen-translate-pane") as
            | (HTMLElement & { setSelectionText: (t: string) => void })
            | null;
          pane?.setSelectionText?.(text);
        }
      }
    }
  };

  readonly #onDocenTranslate = (event: Event): void => {
    const detail = (event as CustomEvent<{ text?: string }>).detail;
    this.#openTranslate(detail?.text);
  };

  #syncStatusLanguage(): void {
    this.#status.syncStatusLanguage();
  }

  /** Options dialog 确定 — commit the UI language + theme + the user identity
   *  + the document settings. */
  readonly #onOptionsOk = (event: Event): void => {
    const {
      lang,
      theme,
      spellcheck,
      markdown,
      identity,
      document: docSettings,
    } = (
      event as CustomEvent<{
        lang?: string;
        theme?: string;
        spellcheck?: boolean;
        markdown?: boolean;
        identity?: { name?: string; initials?: string };
        document?: {
          defaultTabStop?: number;
          updateFields?: boolean;
          updateFieldsBeforePrint?: boolean;
          fieldShading?: "never" | "always" | "whenSelected";
          protection?: string;
          compatVersion?: number;
        };
      }>
    ).detail ?? {};
    if (lang && this.getAttribute("lang") !== lang) {
      this.setAttribute("lang", lang);
      this.#emitLangChange(lang);
    }
    if (theme && this.getAttribute("theme") !== theme) {
      this.setAttribute("theme", theme);
      this.#emitThemeChange(theme);
    }
    if (typeof spellcheck === "boolean") this.#spelling.setEnabled(spellcheck);
    if (typeof markdown === "boolean") {
      this.#markdown = markdown;
      // No transaction rides an options commit — re-stamp the ribbon toggle.
      this.#syncFormatButtons();
    }
    if (identity) updateSettings({ identity });
    if (docSettings) this.#applyDocumentSettings(docSettings);
  };

  /** Options → Proofing → AutoCorrect Options: seed the dedicated dialog from
   *  the effective settings (absent user table = the built-in defaults) and
   *  open it over the Options modal. */
  readonly #openAutocorrectDialog = (): void => {
    const dialog = this.shadowRoot?.querySelector("docen-autocorrect-dialog") as {
      show(values: AutocorrectDialogValues): void;
    } | null;
    dialog?.show(autocorrectConfigOf(getSettings().writing.autocorrect));
  };

  /** AutoCorrect Options 确定 — persist the rule toggles and the edited table;
   *  the per-keystroke config reader picks the new values up immediately. */
  readonly #onAutocorrectOk = (event: CustomEvent<AutocorrectDialogValues>): void => {
    const values = event.detail;
    if (!values) return;
    updateSettings({
      writing: {
        autocorrect: {
          smartQuotes: values.smartQuotes,
          emDash: values.emDash,
          ellipsis: values.ellipsis,
          hyperlinkAutoformat: values.hyperlinkAutoformat,
          capitalizeFirstLetter: values.capitalizeFirstLetter,
          ordinalSuperscript: values.ordinalSuperscript,
          table: {
            version: AUTOCORRECT_TABLE_VERSION,
            replacements: values.replacements.map(({ from, to }) => ({ from, to })),
            exceptions: [...values.exceptions],
          },
        },
      },
    });
    this.#bridge?.focus();
  };

  /** Save Selection to Quick Part Gallery… 确定 — the domain appends the
   *  captured selection as a block (one document transaction). */
  readonly #onQuickPartSave = (event: CustomEvent<QuickPartValues>): void => {
    if (event.detail) this.#buildingBlocks?.commitSave(event.detail);
  };

  /** The Building Blocks Organizer's per-row actions. */
  readonly #onBuildingBlockEvent = (event: CustomEvent<{ id?: string; name?: string }>): void => {
    const { id, name } = event.detail ?? {};
    if (!id) return;
    if (event.type === "building-blocks:insert") this.#buildingBlocks?.insert(id);
    else if (event.type === "building-blocks:rename" && name) {
      this.#buildingBlocks?.renameBlock(id, name);
    } else if (event.type === "building-blocks:delete") {
      this.#buildingBlocks?.deleteBlock(id);
    }
  };

  /** Options → Document commit: fold the dialog's values into
   *  documentExtras.settings (the same channel every settings toggle uses) and
   *  re-derive editability. The tab stop rides the layout projection, so a
   *  change needs the transaction; a no-change OK skips it to avoid planting
   *  an undo step. */
  #applyDocumentSettings(d: {
    defaultTabStop?: number;
    updateFields?: boolean;
    updateFieldsBeforePrint?: boolean;
    fieldShading?: "never" | "always" | "whenSelected";
    protection?: string;
    compatVersion?: number;
  }): void {
    const editor = this.editor;
    if (!editor) return;
    if (d.fieldShading) this.setFieldShading(d.fieldShading);
    if (typeof d.updateFieldsBeforePrint === "boolean") {
      this.setUpdateFieldsBeforePrint(d.updateFieldsBeforePrint);
    }
    const prev = this.#documentSettings();
    const prevTab = typeof prev.defaultTabStop === "number" ? prev.defaultTabStop : undefined;
    const prevProtection = (prev.documentProtection as { edit?: string } | undefined)?.edit;
    const prevCompat = (prev.compatibility as { version?: number } | undefined)?.version;
    const prevUpdateBeforePrint = prev.updateFieldsBeforePrint === true;
    const prevFieldShading = prev.fieldShading;
    // Empty tab input = untouched (undefined survives the round-trip compare).
    const tabTwip =
      d.defaultTabStop != null ? convertMillimetersToTwip(d.defaultTabStop * 10) : prevTab;
    if (
      tabTwip !== prevTab ||
      d.updateFields !== (prev.updateFields === true) ||
      d.updateFieldsBeforePrint !== prevUpdateBeforePrint ||
      d.fieldShading !== prevFieldShading ||
      d.protection !== (prevProtection ?? "none") ||
      d.compatVersion !== (prevCompat ?? 15)
    ) {
      const attrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: Record<string, unknown> };
      const extras = attrs.documentExtras ?? {};
      let settings: Record<string, unknown> = { ...prev };
      if (tabTwip != null) settings.defaultTabStop = tabTwip;
      if (d.updateFields) settings.updateFields = true;
      else delete settings.updateFields;
      if (d.updateFieldsBeforePrint) settings.updateFieldsBeforePrint = true;
      else delete settings.updateFieldsBeforePrint;
      if (d.fieldShading) settings.fieldShading = d.fieldShading;
      settings = withProtection(settings, d.protection ?? "none");
      settings.compatibility = {
        ...(prev.compatibility as object | undefined),
        version: d.compatVersion,
      };
      editor.view.dispatch(
        editor.state.tr.setDocAttribute("documentExtras", { ...extras, settings }),
      );
    }
    // Protection folds into #syncEditable's formula; a tracked-changes
    // restriction additionally forces revision tracking on (Word "start
    // enforcement" behavior).
    if (d.protection === "trackedChanges") editor.commands["track-changes"](true);
    applyProtectionMode(this.#protectionView(), d.protection ?? "none");
  }

  /** References → footnotes group launcher: the Word Footnote and Endnote
   *  dialog, prefilled from documentExtras.settings (Word's defaults for
   *  absent fields: footnotes 1,2,3 at page bottom; endnotes i,ii,iii at
   *  section end). */
  #openNoteSettings(): void {
    const s = this.#documentSettings();
    const fn = (s.footnoteProperties ?? {}) as Record<string, unknown>;
    const en = (s.endnoteProperties ?? {}) as Record<string, unknown>;
    const start = (v: unknown): number => (typeof v === "number" && v >= 1 ? v : 1);
    (
      this.shadowRoot?.querySelector("docen-note-settings-dialog") as {
        show(values?: {
          footnote?: Partial<NoteKindSettings>;
          endnote?: Partial<NoteKindSettings>;
        }): void;
      } | null
    )?.show({
      footnote: {
        pos: fn.pos === "beneathText" ? "beneathText" : "pageBottom",
        numFmt: typeof fn.numFmt === "string" ? fn.numFmt : "decimal",
        numStart: start(fn.numStart),
        numRestart:
          fn.numRestart === "eachSect" || fn.numRestart === "eachPage"
            ? fn.numRestart
            : "continuous",
      },
      endnote: {
        pos: en.pos === "docEnd" ? "docEnd" : "sectEnd",
        numFmt: typeof en.numFmt === "string" ? en.numFmt : "lowerRoman",
        numStart: start(en.numStart),
        numRestart:
          en.numRestart === "eachSect" || en.numRestart === "eachPage"
            ? en.numRestart
            : "continuous",
      },
    });
  }

  /** The footnote/endnote dialog's OK — fold both kinds into
   *  documentExtras.settings (w:footnotePr/w:endnotePr on export). The canvas
   *  paints Word's default ordinal formats only, so these ride the export
   *  path like the Options dialog's compatibility version; fields the dialog
   *  doesn't own (format, footnotes/endnotes separators) survive the fold. A
   *  no-change OK skips the transaction (no undo step). */
  readonly onNoteSettingsOk = (event: Event): void => {
    const v = (event as CustomEvent<NoteSettingsValues | undefined>).detail;
    if (!v) return;
    const prev = this.#documentSettings();
    const prevFn = prev.footnoteProperties as Record<string, unknown> | undefined;
    const prevEn = prev.endnoteProperties as Record<string, unknown> | undefined;
    // Only the four fields the dialog owns decide "changed" — parsed
    // documents may carry extra fields (format, separators) that must not
    // force a dispatch nor get dropped.
    const fnChanged =
      !prevFn ||
      prevFn.pos !== v.footnote.pos ||
      prevFn.numFmt !== v.footnote.numFmt ||
      prevFn.numStart !== v.footnote.numStart ||
      prevFn.numRestart !== v.footnote.numRestart;
    const enChanged =
      !prevEn ||
      prevEn.pos !== v.endnote.pos ||
      prevEn.numFmt !== v.endnote.numFmt ||
      prevEn.numStart !== v.endnote.numStart ||
      prevEn.numRestart !== v.endnote.numRestart;
    if (!fnChanged && !enChanged) return;
    const editor = this.editor;
    if (!editor) return;
    const attrs = (editor.state.doc.attrs ?? {}) as { documentExtras?: Record<string, unknown> };
    const extras = attrs.documentExtras ?? {};
    const settings: Record<string, unknown> = { ...prev };
    if (fnChanged) settings.footnoteProperties = { ...prevFn, ...v.footnote };
    if (enChanged) settings.endnoteProperties = { ...prevEn, ...v.endnote };
    editor.view.dispatch(
      editor.state.tr.setDocAttribute("documentExtras", { ...extras, settings }),
    );
  };

  /** Word's Convert Notes: convert footnotes <-> endnotes or swap them. */
  readonly onNoteSettingsConvert = (event: Event): void => {
    const { mode } =
      (
        event as CustomEvent<{
          mode?: "allFootnotesToEndnotes" | "allEndnotesToFootnotes" | "swapNotes";
        }>
      ).detail ?? {};
    if (!mode) return;
    this.#dialogs.convertNotes(mode);
  };

  /** Open Word's Bookmark dialog. */
  #openBookmarkDialog(): void {
    const bookmarks = this.#dialogs.documentBookmarks();
    (
      this.shadowRoot?.querySelector("docen-bookmark-dialog") as {
        show(bookmarks?: BookmarkItem[]): void;
      } | null
    )?.show(bookmarks);
  }

  readonly onBookmarkAdd = (event: Event): void => {
    const { name } = (event as CustomEvent<{ name?: string } | undefined>).detail ?? {};
    if (!name) return;
    this.#dialogs.addBookmark(name);
  };

  readonly onBookmarkDelete = (event: Event): void => {
    const { name } = (event as CustomEvent<{ name?: string } | undefined>).detail ?? {};
    if (!name) return;
    this.#dialogs.deleteBookmark(name);
  };

  readonly onBookmarkGoTo = (event: Event): void => {
    const { from, to } =
      (event as CustomEvent<{ name?: string; from?: number; to?: number } | undefined>).detail ??
      {};
    if (from == null) return;
    const bridge = this.#bridge;
    if (!bridge) return;
    this.#setTextSelection(from, to ?? from);
    bridge.scrollIntoView(from);
  };

  /** Open Word's Go To dialog (Ctrl+G, Find dropdown). */
  #openGoToDialog(initialKind: GoToKind = "page"): void {
    (
      this.shadowRoot?.querySelector("docen-go-to-dialog") as {
        show(initialKind?: GoToKind): void;
      } | null
    )?.show(initialKind);
  }

  readonly onGoToNavigate = (event: Event): void => {
    const payload = (event as CustomEvent<GoToPayload | undefined>).detail;
    if (!payload) return;
    const { kind, target, direction } = payload;
    const editor = this.editor;
    const bridge = this.#bridge;
    if (!editor || !bridge) return;

    const trimmed = target.trim();
    const currentPos = editor.state.selection.from;

    if (kind === "page") {
      const totalPages = this.#pages.length;
      if (totalPages === 0) return;
      const currentPage = bridge.pageOf(currentPos) ?? 0;
      let destPage = currentPage;
      if (trimmed.startsWith("+") || trimmed.startsWith("-")) {
        const delta = parseInt(trimmed, 10);
        if (Number.isFinite(delta)) destPage = currentPage + delta;
      } else if (trimmed) {
        const num = parseInt(trimmed, 10);
        if (Number.isFinite(num)) {
          destPage = direction ? currentPage + direction * num : num - 1;
        }
      } else if (direction) {
        destPage = currentPage + direction;
      }
      destPage = Math.max(0, Math.min(totalPages - 1, destPage));
      const pos = bridge.firstPosOfPage(destPage);
      if (pos != null) {
        this.#setTextSelection(pos);
        bridge.scrollIntoView(pos);
      }
      return;
    }

    if (kind === "section") {
      const sectionStarts: number[] = [0];
      editor.state.doc.descendants((node, pos) => {
        if (
          node.type.name === "paragraph" &&
          (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
        ) {
          sectionStarts.push(pos + node.nodeSize);
        }
      });
      let curSectionIdx = 0;
      for (let i = 0; i < sectionStarts.length; i++) {
        if (currentPos >= sectionStarts[i]!) curSectionIdx = i;
        else break;
      }
      let destSec = curSectionIdx;
      if (trimmed.startsWith("+") || trimmed.startsWith("-")) {
        const delta = parseInt(trimmed, 10);
        if (Number.isFinite(delta)) destSec = curSectionIdx + delta;
      } else if (trimmed) {
        const num = parseInt(trimmed, 10);
        if (Number.isFinite(num)) {
          destSec = direction ? curSectionIdx + direction * num : num - 1;
        }
      } else if (direction) {
        destSec = curSectionIdx + direction;
      }
      destSec = Math.max(0, Math.min(sectionStarts.length - 1, destSec));
      const targetPos = Math.min(sectionStarts[destSec]!, editor.state.doc.content.size);
      this.#setTextSelection(targetPos);
      bridge.scrollIntoView(targetPos);
      return;
    }

    if (kind === "line") {
      const lineCount = bridge.lineCount();
      if (lineCount === 0) return;
      const curLine = bridge.lineIndexAtPos(currentPos) ?? 0;
      let destLine = curLine;
      if (trimmed.startsWith("+") || trimmed.startsWith("-")) {
        const delta = parseInt(trimmed, 10);
        if (Number.isFinite(delta)) destLine = curLine + delta;
      } else if (trimmed) {
        const num = parseInt(trimmed, 10);
        if (Number.isFinite(num)) {
          destLine = direction ? curLine + direction * num : num - 1;
        }
      } else if (direction) {
        destLine = curLine + direction;
      }
      destLine = Math.max(0, Math.min(lineCount - 1, destLine));
      const pos = bridge.firstPosOfLine(destLine);
      if (pos != null) {
        this.#setTextSelection(pos);
        bridge.scrollIntoView(pos);
      }
      return;
    }

    if (kind === "bookmark") {
      if (trimmed) {
        const bookmarks = this.#dialogs.documentBookmarks();
        const found = bookmarks.find((b) => b.name.toLowerCase() === trimmed.toLowerCase());
        if (found) {
          this.#setTextSelection(found.from, found.to);
          bridge.scrollIntoView(found.from);
        }
      } else if (direction) {
        const bookmarks = this.#dialogs.documentBookmarks();
        if (bookmarks.length === 0) return;
        bookmarks.sort((a, b) => a.from - b.from);
        let nextIdx = 0;
        if (direction > 0) {
          const hit = bookmarks.find((b) => b.from > currentPos);
          nextIdx = hit ? bookmarks.indexOf(hit) : 0;
        } else {
          const rev = [...bookmarks].reverse().find((b) => b.from < currentPos);
          nextIdx = rev ? bookmarks.indexOf(rev) : bookmarks.length - 1;
        }
        const targetBm = bookmarks[nextIdx]!;
        this.#setTextSelection(targetBm.from, targetBm.to);
        bridge.scrollIntoView(targetBm.from);
      }
      return;
    }

    if (kind === "footnote" || kind === "endnote") {
      const notePositions: number[] = [];
      editor.state.doc.descendants((child, pos) => {
        if (child.type.name !== "inlinePassthrough") return;
        try {
          const data = JSON.parse(String(child.attrs?.data ?? "{}")) as Record<string, unknown>;
          if (kind === "footnote" && "footnoteReference" in data) notePositions.push(pos + 1);
          if (kind === "endnote" && "endnoteReference" in data) notePositions.push(pos + 1);
        } catch {}
      });
      if (notePositions.length === 0) return;
      if (trimmed.startsWith("+") || trimmed.startsWith("-")) {
        const delta = parseInt(trimmed, 10);
        let curIdx = notePositions.findIndex((p) => p >= currentPos);
        if (curIdx < 0) curIdx = notePositions.length - 1;
        const destIdx = Math.max(0, Math.min(notePositions.length - 1, curIdx + delta));
        const p = notePositions[destIdx]!;
        this.#setTextSelection(p);
        bridge.scrollIntoView(p);
      } else if (trimmed) {
        const num = parseInt(trimmed, 10);
        if (Number.isFinite(num)) {
          const idx = Math.max(0, Math.min(notePositions.length - 1, num - 1));
          const p = notePositions[idx]!;
          this.#setTextSelection(p);
          bridge.scrollIntoView(p);
        }
      } else if (direction) {
        let p: number | undefined;
        if (direction > 0) {
          p = notePositions.find((pos) => pos > currentPos);
        } else {
          p = [...notePositions].reverse().find((pos) => pos < currentPos);
        }
        if (p != null) {
          this.#setTextSelection(p);
          bridge.scrollIntoView(p);
        }
      }
      return;
    }

    if (kind === "heading") {
      const headingPositions: number[] = [];
      editor.state.doc.descendants((node, pos) => {
        if (node.type.name === "heading") {
          headingPositions.push(pos + 1);
        }
      });
      if (headingPositions.length === 0) return;
      if (trimmed.startsWith("+") || trimmed.startsWith("-")) {
        const delta = parseInt(trimmed, 10);
        let curIdx = headingPositions.findIndex((p) => p >= currentPos);
        if (curIdx < 0) curIdx = headingPositions.length - 1;
        const destIdx = Math.max(0, Math.min(headingPositions.length - 1, curIdx + delta));
        const p = headingPositions[destIdx]!;
        this.#setTextSelection(p);
        bridge.scrollIntoView(p);
      } else if (trimmed) {
        const num = parseInt(trimmed, 10);
        if (Number.isFinite(num)) {
          const idx = Math.max(0, Math.min(headingPositions.length - 1, num - 1));
          const p = headingPositions[idx]!;
          this.#setTextSelection(p);
          bridge.scrollIntoView(p);
        }
      } else if (direction) {
        let p: number | undefined;
        if (direction > 0) {
          p = headingPositions.find((pos) => pos > currentPos);
        } else {
          p = [...headingPositions].reverse().find((pos) => pos < currentPos);
        }
        if (p != null) {
          this.#setTextSelection(p);
          bridge.scrollIntoView(p);
        }
      }
      return;
    }
  };

  /** Open Word's Document Properties dialog. */
  #openPropertiesDialog(): void {
    const editor = this.editor;
    if (!editor) return;
    const attrs = (editor.state.doc.attrs ?? {}) as { core?: Record<string, unknown> };
    const core = attrs.core ?? {};
    let bodyText = "";
    let bodyParas = 0;
    const walk = (node: PMNode, inShape: boolean): void => {
      const shape = inShape || node.type.name === "wpsShape" || node.type.name === "textbox";
      if (node.type.name === "paragraph") {
        if (!shape) {
          bodyParas++;
          bodyText += `${node.textContent}\n`;
        }
        return;
      }
      node.forEach((child) => walk(child, shape));
    };
    walk(editor.state.doc, false);
    const stats: DocumentPropertiesStats = {
      pages: this.#pages.length,
      words: wordCounter(bodyText),
      charsWithSpaces: textCounter(bodyText),
      charsNoSpaces: textCounter(bodyText.replace(/\s+/g, "")),
      paragraphs: bodyParas,
      lines: this.#status.layoutLines(),
      revision: typeof core.revision === "number" ? core.revision : 1,
    };
    (
      this.shadowRoot?.querySelector("docen-properties-dialog") as {
        show(core?: DocumentPropertiesCore, stats?: DocumentPropertiesStats): void;
      } | null
    )?.show(core as DocumentPropertiesCore, stats);
  }

  readonly onPropertiesOk = (event: Event): void => {
    const { core } =
      (event as CustomEvent<{ core?: DocumentPropertiesCore } | undefined>).detail ?? {};
    if (!core) return;
    const editor = this.editor;
    if (!editor) return;
    const attrs = (editor.state.doc.attrs ?? {}) as { core?: Record<string, unknown> };
    const prevCore = attrs.core ?? {};
    const nextCore: Record<string, unknown> = { ...prevCore };
    if (core.title !== undefined) nextCore.title = core.title;
    if (core.subject !== undefined) nextCore.subject = core.subject;
    if (core.creator !== undefined) nextCore.creator = core.creator;
    if (core.keywords !== undefined) nextCore.keywords = core.keywords;
    if (core.description !== undefined) nextCore.description = core.description;
    editor.view.dispatch(editor.state.tr.setDocAttribute("core", nextCore));
    this.#dialogs.updateAllFields();
  };

  /** Notify external listeners (framework wrappers like @docen/vue) when the
   *  locale changes from inside the host — status-bar toggle or Options OK.
   *  External `lang` writes (e.g. a Vue prop) set the attribute directly and
   *  don't route through here, so there's no echo cycle. */
  #emitLangChange(lang: string): void {
    this.dispatchEvent(
      new CustomEvent("docen:lang-change", { bubbles: true, composed: true, detail: { lang } }),
    );
  }

  /** Notify external listeners (framework wrappers like @docen/vue) when the
   *  theme changes from inside the host — Options OK. External `theme` writes
   *  (e.g. a Vue prop) set the attribute directly and don't route through
   *  here, so there's no echo cycle. */
  #emitThemeChange(theme: string): void {
    this.dispatchEvent(
      new CustomEvent("docen:theme-change", { bubbles: true, composed: true, detail: { theme } }),
    );
  }

  #pickFile(): void {
    this.#io.pickFile();
  }

  async #saveAs(format: Exclude<SaveFormat, "pdf"> = this.#docxVariant): Promise<void> {
    return this.#io.saveAs(format);
  }

  async #saveAsTemplate(): Promise<void> {
    return this.#io.saveAsTemplate();
  }

  async #saveAsPdf(): Promise<void> {
    return this.#io.saveAsPdf();
  }

  async #share(): Promise<void> {
    return this.#io.share();
  }

  #closeDocument(): void {
    this.#io.closeDocument();
  }

  #inspectFindings(): { comments: number; revisions: number } {
    return this.#io.inspectFindings();
  }

  #inspectDocument(): void {
    this.#io.inspectDocument();
  }

  /** The inspector dialog's removal buttons — clear comments / accept all
   *  revisions, then re-hand the scan so the counts fall to zero in place. */
  readonly #onInspect = (event: Event): void => {
    if (event.type === "inspect:clear-comments") {
      this.#comments.deleteAllComments();
    } else if (event.type === "inspect:accept-revisions") {
      const commands = this.editor?.commands;
      if (commands)
        (commands as unknown as Record<string, () => unknown>)["accept-all-changes"]?.();
    }
    const dialog = this.shadowRoot?.querySelector("docen-inspect-dialog");
    dialog?.setAttribute("findings", JSON.stringify(this.#inspectFindings()));
  };

  #syncDocumentSettings(settings: Record<string, unknown>): void {
    const fs = settings.fieldShading as "never" | "always" | "whenSelected" | undefined;
    if (fs === "never" || fs === "always" || fs === "whenSelected") {
      this.#fieldShading = fs;
    }
    this.#updateFieldsBeforePrint = settings.updateFieldsBeforePrint === true;
    this.#stage?.setFieldShading(this.#fieldShading);
  }

  getFieldShading(): "never" | "always" | "whenSelected" {
    return this.#fieldShading;
  }

  setFieldShading(mode: "never" | "always" | "whenSelected"): void {
    this.#fieldShading = mode;
    this.#stage?.setFieldShading(mode);
  }

  getUpdateFieldsBeforePrint(): boolean {
    return this.#updateFieldsBeforePrint;
  }

  setUpdateFieldsBeforePrint(value: boolean): void {
    this.#updateFieldsBeforePrint = value;
  }

  getPrintMarkup(): boolean {
    return this.#printMarkup;
  }

  setPrintMarkup(value: boolean): void {
    this.#printMarkup = value;
  }

  /**
   * Attach an external .dotx (or .docx) template to this document, importing
   * styles, numbering, docDefaults, and themes without replacing body content.
   */
  attachTemplate(
    data: Uint8Array | ArrayBuffer,
    options: { autoUpdateStyles?: boolean } = {},
  ): DotxTemplatePackage | undefined {
    if (!this.editor) return undefined;
    const pkg = attachTemplate(this.editor, data, options);
    if (pkg.theme) {
      this.#applyDocumentTheme("theme", pkg.theme.id, false);
    }
    this.repaginate();
    return pkg;
  }

  /**
   * Open the Print Preview modal (<docen-print-preview>) with multi-page preview,
   * zoom/navigation, page range, collation, booklet imposition, and markup toggle.
   */
  async openPrintPreview(options?: { markup?: boolean }): Promise<void> {
    if (this.#updateFieldsBeforePrint) {
      updateDynamicFieldsBeforePrint(this.#dialogs, this.editor?.commands, (pos) => {
        const page = this.#bridge?.pageOf(pos);
        return typeof page === "number" ? page + 1 : null;
      });
    }
    const previewEl = this.shadowRoot?.querySelector("docen-print-preview") as
      | (HTMLElement & {
          seed(options: unknown): void;
          show(): void;
        })
      | null;
    if (!previewEl || !this.#stage) return;

    const printMarkup = options?.markup ?? this.#printMarkup;
    const shots = await this.#stage.printSnapshots({ markup: printMarkup });
    const cursor = this.editor?.state.selection.from ?? 0;
    const page = (this.#bridge?.pageOf(cursor) ?? 0) + 1;
    previewEl.seed({
      snapshots: shots,
      filename: this.getAttribute("filename") ?? "Document",
      currentPage: page,
      printMarkup,
    });
    previewEl.show();
  }

  async #print(): Promise<void> {
    if (this.#updateFieldsBeforePrint) {
      updateDynamicFieldsBeforePrint(this.#dialogs, this.editor?.commands, (pos) => {
        const page = this.#bridge?.pageOf(pos);
        return typeof page === "number" ? page + 1 : null;
      });
    }
    return this.#io.print();
  }

  async open(file: File): Promise<void> {
    return this.#io.open(file);
  }

  async openDOCX(
    input: File | ArrayBuffer | Uint8Array,
    variant: DocxVariant = "docx",
  ): Promise<void> {
    return this.#io.openDOCX(input, variant);
  }

  #newFromTemplate(id: string): void {
    this.#io.newFromTemplate(id);
  }
  #openTemplateDialog(): void {
    this.#io.openTemplateDialog();
  }

  async openMarkdown(input: File | string): Promise<void> {
    return this.#io.openMarkdown(input);
  }

  async openRTF(input: File | string): Promise<void> {
    return this.#io.openRTF(input);
  }

  async openPlainText(input: File | string): Promise<void> {
    return this.#io.openPlainText(input);
  }

  #setProgress(label?: string): void {
    this.#io.setProgress(label);
  }

  async saveDOCX(variant: DocxVariant = this.#docxVariant): Promise<Uint8Array> {
    return this.#io.saveDOCX(variant);
  }

  /**
   * Register font bytes for deterministic shaping and export embedding.
   * Pass the family name the document styles reference (e.g. "Calibri") and,
   * for a `.ttc`/`.otc` collection (Noto Sans CJK, msyh.ttc), the face
   * `fontIndex`. Shaping is on by default — a registered family shapes under
   * the engine and feeds the line breaker; families without a registered face
   * keep the canvas fallback (`setShapingEnabled(false)` /
   * `DOCEN_SHAPING_DISABLED=1` rolls back). PDF/DOCX export embeds the font
   * whenever its `fsType` allows; collection faces cannot be subset by the
   * embedded-font writer yet — register an extracted face for export
   * embedding.
   */
  async registerFont(family: string, fontData: Uint8Array, fontIndex = 0): Promise<void> {
    this.#fonts.set(family.toLowerCase(), { family, fontData });
    await initShapingWasm();
    registerShapingFont(family, fontData, fontIndex);
    this.#measurer.clearCache();
  }

  /**
   * Register docen's bundled production faces for the Word default families
   * (Calibri/Calibri Light → Carlito, Cambria → Caladea, Arial → Liberation
   * Sans, Times New Roman → Liberation Serif; all four weight/slant slots).
   * The element calls this automatically on connect, so shaping never silently
   * falls back to canvas for the default families; call it yourself (with
   * `baseUrl` when your bundler does not emit the package's asset URLs, or
   * after overriding the bundled files) to control it. Registered regular
   * faces additionally join the export-embedding set.
   *
   * Failures are non-fatal: the editor warns and keeps the canvas fallback.
   */
  async registerDefaultFonts(options?: RegisterDefaultFontsOptions): Promise<string[]> {
    const faces = await loadDefaultFonts(options);
    await initShapingWasm();
    const labels: string[] = [];
    for (const face of faces) {
      registerShapingFont(face.family, face.bytes, 0, {
        bold: face.bold,
        italic: face.italic,
      });
      // One face per family feeds the export-embedding map (its regular face,
      // matching registerFont's family-keyed contract).
      if (!face.bold && !face.italic) {
        this.#fonts.set(face.family.toLowerCase(), { family: face.family, fontData: face.bytes });
      }
      labels.push(`${face.family}${face.bold ? " bold" : ""}${face.italic ? " italic" : ""}`);
    }
    this.#measurer.clearCache();
    return labels;
  }

  /** Serialize the current document to a Markdown string. */
  saveMarkdown(): string {
    return this.#io.saveMarkdown();
  }

  saveRTF(): string {
    return this.#io.saveRTF();
  }

  saveHTML(options?: HtmlGenerateOptions): string {
    return this.#io.saveHTML(options);
  }

  savePlainText(): string {
    return this.#io.savePlainText();
  }

  async saveODT(): Promise<Uint8Array> {
    return this.#io.saveODT();
  }

  getJSON(): JSONContent {
    return this.#io.getJSON();
  }

  setJSON(json: JSONContent): void {
    this.#io.setJSON(json);
  }

  #loadDoc(doc: JSONContent): void {
    // A new document gets a fresh warning state — a dismissal does not leak
    // from the previous document into this one.
    this.#contentWarningDismissed = false;
    this.#io.loadDoc(doc);
  }

  #documentSettings(): Record<string, unknown> {
    return this.#io.documentSettings();
  }

  /** The underlying Tiptap editor (for advanced, direct control). */
  getEditor(): Editor | undefined {
    return this.editor;
  }

  /** Force a full canvas re-render now. */
  repaginate(): void {
    const editor = this.editor;
    if (editor) this.#renderDoc(editor.getJSON());
  }

  // ── Task pane visibility (Office.addin.showAsTaskpane / hide equivalent) ──

  /** Show a task pane. No-op if already open. */
  showTaskpane(id: TaskPaneId): void {
    this.#setTaskpane(id, true);
  }

  /** Hide a task pane. No-op if already closed. */
  hideTaskpane(id: TaskPaneId): void {
    this.#setTaskpane(id, false);
  }

  /** Whether a task pane is currently open. Returns a boolean for convenience
   *  (callers want open/closed); the `docen:taskpane-visibility-change` event
   *  detail carries the string `VisibilityMode` to mirror `Office.VisibilityMode`. */
  getTaskpaneState(id: TaskPaneId): boolean {
    return !!this.#paneEl(id)?.open;
  }

  #paneEl(id: TaskPaneId): (HTMLElement & { open: boolean }) | null {
    // Panes are looked up by part, not position — several panes share the end
    // rail (properties + comments + clipboard + spelling).
    const part =
      id === "navigation"
        ? "nav-pane"
        : id === "comments"
          ? "comments-pane"
          : id === "clipboard"
            ? "clipboard-pane"
            : id === "proofing"
              ? "proofing-pane"
              : id === "thesaurus"
                ? "thesaurus-pane"
                : id === "translate"
                  ? "translate-pane"
                  : id === "revisions"
                    ? "revisions-pane"
                    : id === "styles"
                      ? "styles-pane"
                      : id === "reveal"
                        ? "reveal-pane"
                        : id === "restrict"
                          ? "restrict-pane"
                          : id === "a11y"
                            ? "a11y-pane"
                            : id === "altText"
                              ? "alt-text-pane"
                              : "props-pane";
    return this.shadowRoot?.querySelector(`docen-task-pane[part="${part}"]`) as
      | (HTMLElement & { open: boolean })
      | null;
  }

  /** Apply a visibility state and dispatch `docen:taskpane-visibility-change`
   *  when it flips. The detail carries `visibilityMode: "taskpane"|"hidden"` to
   *  mirror `Office.VisibilityMode`. Idempotent — no event when state is
   *  unchanged. Opening a pane dismisses its rail-mates (Word's task panes are
   *  mutually exclusive per side — Comments replaces Properties, never
   *  stacks with it). */
  #setTaskpane(id: TaskPaneId, open: boolean): void {
    const pane = this.#paneEl(id);
    if (!pane || pane.open === open) return;
    if (open) {
      const side = pane.getAttribute("position") ?? "start";
      for (const other of this.shadowRoot?.querySelectorAll("docen-task-pane[open]") ?? []) {
        if (other !== pane && (other.getAttribute("position") ?? "start") === side) {
          (other as HTMLElement & { open: boolean }).open = false;
        }
      }
    }
    pane.open = open;
    if (open) {
      if (id === "a11y") {
        (this.shadowRoot?.querySelector("docen-a11y-checker-pane") as any)?.check(this.getJSON());
      } else if (id === "altText") {
        this.#syncAltTextPane();
      } else if (id === "reveal") {
        this.#updateRevealFormatting();
      } else if (id === "proofing") {
        // The pane skips updates while hidden — it needs the current list.
        this.#spelling.syncPane();
      } else if (id === "navigation") {
        if (this.#pages.length > 0) {
          const total = this.#pages.length;
          const thumbs: (string | null)[] = [];
          for (let i = 0; i < total; i++) {
            thumbs.push(this.#stage?.pageThumbnail(i) ?? null);
          }
          this.#status.cacheNavThumbs(thumbs);
          this.#status.pushNavPages(
            total,
            (this.#bridge?.pageOf(this.editor?.state.selection.from ?? 0) ?? 0) + 1,
          );
          // Off-screen pages have no canvas yet — rasterize every page once so
          // the pane shows real thumbnails for the whole document, not just
          // the pages the viewport already reached.
          void this.#refreshNavThumbnails();
        }
      }
    }
    this.dispatchEvent(
      new CustomEvent("docen:taskpane-visibility-change", {
        bubbles: true,
        composed: true,
        detail: { id, visibilityMode: (open ? "taskpane" : "hidden") as VisibilityMode },
      }),
    );
  }

  // ── Zoom (method + event + getter; once `zoom` attr seeds #zoom) ──

  /** Apply a zoom level (percent, clamped 10–500). Idempotent; dispatches
   *  `docen:zoom-change` on a real change (mirrors `Office.Document.zoom.set`). */
  setZoom(pct: number): void {
    this.#setZoom(pct);
  }

  getZoom(): number {
    return this.#status.getZoom();
  }

  // ── Formatting marks (method + event; boolean `show-marks` attribute) ──

  /** Toggle editing/formatting marks on or off. Idempotent; dispatches
   *  `docen:marks-change`. The boolean `show-marks` attribute is the source of
   *  truth; the stage paints the ↵/→/· marks and break rows from it. */
  setShowMarks(on: boolean): void {
    if (this.hasAttribute("show-marks") === on) return;
    this.toggleAttribute("show-marks", on);
    this.#stage?.setShowMarks(on);
    // Marks toggling rides no transaction — re-stamp the ribbon lit state.
    this.#syncFormatButtons();
    this.dispatchEvent(
      new CustomEvent("docen:marks-change", {
        bubbles: true,
        composed: true,
        detail: { showMarks: on },
      }),
    );
  }

  /** Whether editing/formatting marks are currently shown. */
  getShowMarks(): boolean {
    return this.hasAttribute("show-marks");
  }

  // ── Ruler (method + event; boolean `show-ruler` / `ruler` attribute) ──

  /** Toggle ruler on or off. Idempotent; dispatches `docen:ruler-change`. */
  setShowRuler(on: boolean): void {
    if (this.getShowRuler() === on) return;
    this.toggleAttribute("show-ruler", on);
    this.#stage?.setShowRuler(on);
    this.#syncRuler();
    this.dispatchEvent(
      new CustomEvent("docen:ruler-change", {
        bubbles: true,
        composed: true,
        detail: { showRuler: on },
      }),
    );
  }

  /** Whether ruler is currently shown. */
  getShowRuler(): boolean {
    return (
      this.hasAttribute("show-ruler") ||
      this.hasAttribute("ruler") ||
      (this.#stage?.showRuler ?? false)
    );
  }

  // ── Persisted settings (identity + writing toggles) ────────────────────────

  /** The effective settings: the shared persisted store with the `user`
   *  attribute's identity projection applied (the attribute wins while set). */
  get settings(): DocenSettings {
    const stored = getSettings();
    return { ...stored, identity: resolveIdentity(this.getAttribute("user"), stored.identity) };
  }

  /** Merge a settings patch into the shared persisted store. A real change
   *  persists, re-renders the chrome, and bubbles `docen:settings-change`
   *  (`detail: { settings }`) out of the element. */
  setSettings(patch: SettingsPatch): void {
    updateSettings(patch);
  }
}

// Persisted settings + identity store — the same module the autocorrect lane
// (rule config + user table) and host consumers import.
export {
  AUTOCORRECT_TABLE_VERSION,
  SETTINGS_STORAGE_KEY,
  SETTINGS_VERSION,
  createSettingsStore,
  defaultSettings,
  getSettings,
  initialsFromName,
  onSettingsChange,
  resolveIdentity,
  updateSettings,
} from "./settings";
export type {
  AutocorrectReplacement,
  AutocorrectSettings,
  AutocorrectTable,
  DocenSettings,
  IdentitySettings,
  SettingsListener,
  SettingsPatch,
  SettingsStore,
  SettingsStorage,
  WritingSettings,
} from "./settings";

// Building blocks (Quick Parts / AutoText) — the docen model, its document
// persistence (documentExtras.docenBlocks) and the derived Word glossary part
// (word/glossary/document.xml) projection.
export {
  autotextMatch,
  BLOCK_GALLERIES,
  blocksFromGlossary,
  blocksOfDocAttrs,
  BUILDING_BLOCKS_VERSION,
  createBuildingBlock,
  DEFAULT_BLOCK_CATEGORY,
  DEFAULT_BLOCK_GALLERY,
  glossaryOfBlocks,
  groupBlocksByGallery,
  isDuplicateBlockName,
  parseBuildingBlocks,
  parseSlicePayload,
  sortBlocks,
  withBlocks,
} from "./building-blocks";
export type {
  BuildingBlock,
  BuildingBlockInsertMode,
  BuildingBlockSlice,
  BuildingBlocksData,
} from "./building-blocks";
export type { BuildingBlocksSeed } from "../ui/components/workspace/building-blocks-dialog";
export type { QuickPartSeed, QuickPartValues } from "../ui/components/workspace/quick-part-dialog";

/**
 * `<docen-editor>` — Turnkey rich document editor element.
 * Drop-in custom element wrapping DocenDocument with touch, pen, and RTL support.
 */
@customElement({ name: "docen-editor", template: documentTemplate, styles: documentStyles })
export class DocenEditor extends DocenDocument {}

export { DocenDocument };
export default DocenDocument;
