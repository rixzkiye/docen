/**
 * @docen/editor — assembly layer that bundles the Fluent UI shell with
 * @docen/docx into turnkey editor elements. The UI layer lives in ./ui
 * (inlined so i18n side effects stay in one bundle, with no cross-package
 * tree-shaking hazard).
 *
 * @module
 */

// Editor elements (register their custom elements on import)
export { default as DocenDocument, DocenEditor } from "./document";
export { default as DocenPresentation } from "./presentation";
export { default as DocenWorkbook } from "./workbook";

// Re-export the UI bootstrap so demos/consumers import everything from one
// entry — matches importing ./src/index.ts directly.
export {
  applyTheme,
  availableLanguages,
  builtinThemes,
  getUiDirection,
  notifyLocaleChange,
  observeLang,
  registerComponents,
  registerLocalization,
  registerTheme,
  registerTranslation,
  resolveDir,
  resolveTheme,
  setUiDirection,
  t,
} from "./ui";

// Public types for add-in authors and host consumers.
export type { TaskPaneId, VisibilityMode } from "./document";
export type { AddinHost } from "./ui/addin/host";
export type { DocenHost, DocenAddin, RibbonTab } from "./ui/addin/types";
export type { AdditionalLanguage, LanguageOption, LocalizationInfo } from "./ui";

// Persisted editor settings + the user identity store (localStorage-backed,
// shared by every host element on the page).
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
} from "./document";
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
} from "./document";

// Building blocks (Quick Parts / AutoText): the docen model, the document
// persistence helpers (documentExtras.docenBlocks) and the Word glossary
// projection. DocPartGallery/DocPartOptions come from @docen/docx.
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
  sortBlocks,
  withBlocks,
} from "./document";
export type {
  BuildingBlock,
  BuildingBlockInsertMode,
  BuildingBlockSlice,
  BuildingBlocksData,
} from "./document";

// Fluent theme factories re-exported so registerTheme() callers build brand
// themes (createLightTheme/createDarkTheme) from @docen/editor alone — no need
// to depend on @fluentui/tokens directly. Mirrors how an Office.js host hands
// add-ins a Fluent Theme object (fluentThemeData).
export { createDarkTheme, createHighContrastTheme, createLightTheme } from "@fluentui/tokens";
export type { BrandVariants, Theme } from "@fluentui/tokens";

// PDF export engine and types (vector, PDF/A-2b, PDF/UA-1, fonts)
export {
  pagesToPdf,
  buildEmbeddedPdfFonts,
  DEFAULT_EDITOR_TEXT_MODE,
  DEFAULT_SERVER_TEXT_MODE,
} from "./document";
export type {
  PdfExportOptions,
  PdfPageShot,
  PdfEmbeddedFont,
  PdfEmbeddableFontSource,
  PdfViewerPreferences,
  PdfFormField,
  PdfStructElement,
  PdfOutlineItem,
  PdfPageLabelRange,
  PdfDestination,
} from "./document";
