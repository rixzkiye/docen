import {
  t,
  type RibbonControl,
  type RibbonControlOrLayout,
  type RibbonControlSize,
  type RibbonGroup,
  type RibbonLayout,
  type RibbonMenuItem,
  type RibbonTab,
} from "../../ui";

/**
 * Build the ribbon DOM (fluent-tablist + one panel per tab + trailing actions)
 * imperatively from a {@link RibbonTab} schema tree. Same shape as the old
 * HTML-string builder, but typed and data-driven — an addin merges its own
 * tabs/groups into {@link ribbonTabs} before this runs, so the ribbon is
 * externally customizable without host internals. Re-call on a locale change
 * (labels re-resolve in the schema).
 *
 * `tabs` is already the visible subset ({@link ribbonTabs}); the active tab
 * falls back to the first so fluent-tablist never points activeid at a missing id.
 */
export function renderRibbonFromSchema(
  tabs: readonly RibbonTab[],
  actions: readonly RibbonControl[] = [],
  scope: Element = document.documentElement,
): DocumentFragment {
  const frag = document.createDocumentFragment();

  const tablist = document.createElement("fluent-tablist");
  tablist.setAttribute("slot", "tabs");
  tablist.setAttribute("appearance", "transparent");
  const activeId = tabs[0]?.id ?? "";
  if (activeId) tablist.setAttribute("activeid", activeId);
  frag.append(tablist);

  for (const tab of tabs) {
    if (tab.contextual) continue; // the host appends these on selection
    const built = buildContextualTab(tab, scope);
    tablist.append(built.tab);
    frag.append(built.panel);
  }

  for (const c of actions) {
    const el = buildControl(c, scope);
    el.setAttribute("slot", "actions");
    frag.append(el);
  }
  return frag;
}

/** Build one tab's DOM pair — the `docen-ribbon-tab` heading and its
 *  `docen-ribbon-panel` of groups. The static render appends these into the
 *  tablist/panel container; the host reuses the same builder to append/remove
 *  contextual tabs (Table Design/Layout) as the selection enters/leaves their
 *  owning context. */
export function buildContextualTab(
  tab: RibbonTab,
  scope: Element = document.documentElement,
): { tab: HTMLElement; panel: HTMLElement } {
  const tabEl = document.createElement("docen-ribbon-tab");
  tabEl.setAttribute("slot", "tab");
  tabEl.id = tab.id;
  tabEl.textContent = t(tab.label, scope);
  const panel = document.createElement("docen-ribbon-panel");
  panel.setAttribute("value", tab.id);
  for (const g of tab.groups) panel.append(buildGroup(g, scope));
  return { tab: tabEl, panel };
}

function buildGroup(g: RibbonGroup, scope: Element): HTMLElement {
  const el = document.createElement("docen-ribbon-group");
  el.setAttribute("label", t(g.label, scope));
  if (g.launcher) el.setAttribute("launcher", g.launcher);
  for (const c of g.controls) el.append(buildControlOrLayout(c, scope));
  return el;
}

function buildControlOrLayout(c: RibbonControlOrLayout, scope: Element): HTMLElement {
  return c.type === "layout" ? buildLayout(c, scope) : buildControl(c, scope);
}

function buildLayout(l: RibbonLayout, scope: Element): HTMLElement {
  const el = document.createElement("div");
  el.className = l.layout === "column" ? "rb-col" : l.layout === "row" ? "rb-row" : "rb-grid";
  if (l.layout === "grid" && l.columns) {
    el.dataset.columns = "";
    el.style.setProperty("--rb-grid-cols", String(l.columns));
  }
  for (const c of l.controls) el.append(buildControlOrLayout(c, scope));
  return el;
}

/** Stamp the shared base attrs (icon/label/event/value/iconOnly/size/disabled)
 *  every control component reads. */
function applyBase(
  el: HTMLElement,
  c: {
    icon?: string;
    label?: string;
    event?: string;
    value?: string;
    iconOnly?: boolean;
    size?: RibbonControlSize;
    disabled?: boolean;
  },
  scope: Element,
): void {
  if (c.icon) el.setAttribute("icon", c.icon);
  if (c.label) el.setAttribute("label", t(c.label, scope));
  if (c.event) el.setAttribute("event", c.event);
  if (c.value) el.setAttribute("value", c.value);
  if (c.iconOnly) el.setAttribute("icon-only", "");
  if (c.size) el.setAttribute("size", c.size);
  if (c.disabled) el.setAttribute("disabled", "");
}

/** Resolve each item's `text` i18n key to the active locale — the schema stores
 *  keys, the render pass is the single translate point (mirrors label).
 *  Submenus recurse: nested `children`/`items` carry keys too. */
const translateItems = (items: readonly RibbonMenuItem[], scope: Element): RibbonMenuItem[] =>
  items.map((it) => ({
    ...it,
    text: it.text ? t(it.text, scope) : it.text,
    ...(it.children ? { children: translateItems(it.children, scope) } : {}),
    ...(it.items ? { items: translateItems(it.items, scope) } : {}),
  }));

function buildControl(c: RibbonControl, scope: Element): HTMLElement {
  switch (c.type) {
    case "separator": {
      const el = document.createElement("span");
      el.className = "rb-vsep";
      return el;
    }
    case "button": {
      const el = document.createElement(
        c.toggle ? "docen-ribbon-toggle-button" : "docen-ribbon-button",
      );
      applyBase(el, c, scope);
      return el;
    }
    case "checkbox": {
      const el = document.createElement("docen-ribbon-checkbox");
      applyBase(el, c, scope);
      if (c.checked) el.setAttribute("checked", "");
      return el;
    }
    case "menu": {
      const el = document.createElement("docen-ribbon-menu");
      applyBase(el, c, scope);
      el.setAttribute("items", JSON.stringify(translateItems(c.items ?? [], scope)));
      return el;
    }
    case "split": {
      const el = document.createElement("docen-ribbon-split-button");
      applyBase(el, c, scope);
      el.setAttribute("items", JSON.stringify(translateItems(c.items ?? [], scope)));
      // Opt the face into the host's lit sync; pressed itself is the sync's
      // write, not the render's (single writer keeps render idempotent).
      if (c.toggle) el.setAttribute("pressed-face", "");
      return el;
    }
    case "combobox": {
      const el = document.createElement("docen-ribbon-combobox");
      applyBase(el, c, scope);
      if (c.value != null) el.setAttribute("value", c.value);
      el.setAttribute("items", JSON.stringify(translateItems(c.items ?? [], scope)));
      if (c.source) el.setAttribute("source", c.source);
      if (c.comboboxSize === "short") el.setAttribute("size", "short");
      return el;
    }
    case "input": {
      const el = document.createElement("docen-ribbon-input");
      applyBase(el, c, scope);
      if (c.value != null) el.setAttribute("value", c.value);
      return el;
    }
    case "color-picker": {
      const el = document.createElement("docen-color-picker");
      applyBase(el, c, scope);
      if (c.defaultColor) el.setAttribute("default-color", c.defaultColor);
      if (c.palette) el.setAttribute("palette", c.palette);
      if (c.panel) el.setAttribute("panel", c.panel);
      return el;
    }
    case "gallery": {
      const el = document.createElement("docen-ribbon-gallery");
      applyBase(el, c, scope);
      el.setAttribute("items", JSON.stringify(translateItems(c.items ?? [], scope)));
      if (c.visibleCount != null) el.setAttribute("visible-count", String(c.visibleCount));
      return el;
    }
  }
}
