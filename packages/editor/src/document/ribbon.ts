/**
 * Default MS Office Word ribbon for `<docen-document>` — all 9 standard tabs
 * (Home/Insert/Draw/Design/Layout/References/Mailings/Review/View) with the
 * canonical groups and primary commands.
 *
 * `ribbonTabs()` builds the RibbonTab schema with i18n keys (not translated
 * strings); `renderRibbonFromSchema()` consumes that tree and resolves every
 * label via `t("ribbon.*")` when building the ribbon DOM. The
 * host stamps the result into its `<docen-ribbon>` and re-runs it on language
 * change. Callers wanting a tailored ribbon merge their own tabs/groups into
 * the schema before render.
 *
 * Layout helpers (`.rb-col` / `.rb-row` / `.rb-vsep`) are injected by the host
 * style — Office groups stack a large button beside rows/columns of small
 * `icon-only` buttons.
 *
 * Each command carries an `event` name. `DocumentCommands` (extensions/commands)
 * wires the ones the Tiptap engine supports today (marks, lists, alignment,
 * styles, breaks, history); the rest render as a complete visual skeleton and
 * no-op on click until wired.
 *
 * The tab builders live in `ribbon/<tab>.ts`; this module is the public
 * surface (same names and signatures as before the split).
 */
export { chartDesignTab } from "./ribbon/chart-design";
export { equationContextTab } from "./ribbon/equation";
export { headerFooterContextTab } from "./ribbon/header-footer";
export { pictureFormatTab } from "./ribbon/picture-format";
export { buildContextualTab, renderRibbonFromSchema } from "./ribbon/render";
export { shapeFormatTab } from "./ribbon/shape-format";
export {
  DEFAULT_RIBBON_TAB,
  RIBBON_TAB_IDS,
  styleGalleryItems,
  stylePreviewCss,
  type RibbonTabId,
} from "./ribbon/shared";
export { formatMeasureTwip, tableContextTabs, useCmUnits } from "./ribbon/table";
export { ribbonActions, ribbonTabs, type RibbonOptions } from "./ribbon/tabs";
