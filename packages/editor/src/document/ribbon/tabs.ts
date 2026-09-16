import type { StylesOptions } from "@docen/docx";

import type { RibbonControl, RibbonTab } from "../../ui";
import { designTab } from "./design";
import { developerTab } from "./developer";
import { drawTab } from "./draw";
import { homeTab } from "./home";
import { insertTab } from "./insert";
import { layoutTab } from "./layout";
import { mailingsTab } from "./mailings";
import { referencesTab } from "./references";
import { reviewTab } from "./review";
import { btn, cmd, menu, opt, parsedItems, RIBBON_TAB_IDS, type RibbonTabId } from "./shared";
import { viewTab } from "./view";

// Edit / View mode pick — the tab-row "Editing" trailing action. Default is
// Edit checked; the host (#syncEditModeMenu in document/index.ts) rewrites the
// label + checked state to match the live editable state, so this is only the
// initial stamp.
const editItems = (): string =>
  JSON.stringify([
    { text: opt("editing"), event: "edit-mode", value: "edit", checked: true },
    { text: opt("viewing"), event: "edit-mode", value: "view" },
  ]);

/** Options for {@link buildRibbonInnerHTML}. */
export interface RibbonOptions {
  /** Whitelist of tab ids to render; omitted/empty = all tabs (back-compat). */
  tabs?: readonly RibbonTabId[];
  /** The document's revision authors (w:ins/@w:author values, document order)
   *  — fills the markup view's Specific People menu; absent/empty greys it. */
  revisionAuthors?: readonly string[];
}

/** Default ribbon tabs for the active locale (and the loaded document's styles,
 *  for the Styles gallery). Pass `{ tabs }` to render a subset. */
export function ribbonTabs(styles?: StylesOptions | null, opts: RibbonOptions = {}): RibbonTab[] {
  const visible: readonly RibbonTabId[] =
    opts.tabs && opts.tabs.length > 0 ? opts.tabs : RIBBON_TAB_IDS;
  const show = (id: RibbonTabId): boolean => visible.includes(id);
  const tabs: RibbonTab[] = [];
  if (show("home")) tabs.push(homeTab(styles));
  if (show("insert")) tabs.push(insertTab());
  if (show("draw")) tabs.push(drawTab());
  if (show("design")) tabs.push(designTab());
  if (show("layout")) tabs.push(layoutTab());
  if (show("references")) tabs.push(referencesTab());
  if (show("mailings")) tabs.push(mailingsTab());
  if (show("review")) tabs.push(reviewTab(opts.revisionAuthors));
  if (show("view")) tabs.push(viewTab());
  if (show("developer")) tabs.push(developerTab());
  return tabs;
}

/** Trailing ribbon actions (right of the tabs): comment / edit-mode / share. */
export function ribbonActions(): RibbonControl[] {
  return [
    btn("comment", "comment"),
    menu("edit", "edit-mode", parsedItems(editItems()), { label: cmd("editing") }),
    btn("share", "share"),
  ];
}
