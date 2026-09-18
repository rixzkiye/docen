import type { JSONContent } from "@docen/docx";
import type { Editor } from "@docen/docx/core";

/**
 * Design → Set as Default: the document formatting a newly created document
 * starts from. Word persists this on the Normal template; the browser element
 * has no filesystem template, so the pair (theme + style set) is stored in
 * `localStorage` under a versioned key and re-applied by {@link newDocumentJSON}
 * / {@link applyDocumentDefaults} whenever the editor starts a blank document.
 *
 * Storage is treated as a cache exactly like the settings store: a corrupt
 * entry, a foreign version, or blocked storage falls back to "no defaults"
 * (the factory Office theme) instead of throwing.
 */

/** The remembered formatting of `Set as Default`. */
export interface DocumentDefaults {
  /** The theme id (see THEMES — "office", "facet", …). */
  theme: string;
  /** The theme resolver kind the theme was applied through ("theme" /
   *  "theme-color" / "theme-font"). */
  themeKind: string;
  /** The style-set preset id ("modern" / "classic" / "elegant"), or null when
   *  the document never applied one. */
  styleSet: string | null;
}

/** The versioned localStorage key. */
export const DOCUMENT_DEFAULTS_STORAGE_KEY = "docen.document-defaults.v1";

/** The schema version persisted alongside the defaults. */
export const DOCUMENT_DEFAULTS_VERSION = 1;

/** The storage surface this module needs — structural subset of `Storage`. */
export interface DefaultsStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

/** Defensive parse: bad JSON, wrong shape, or a foreign version yield null. */
export function parseDocumentDefaults(raw: string | null): DocumentDefaults | null {
  if (!raw) return null;
  let data: unknown;
  try {
    data = JSON.parse(raw);
  } catch {
    return null;
  }
  if (typeof data !== "object" || data === null) return null;
  const record = data as Record<string, unknown>;
  if (record.version !== DOCUMENT_DEFAULTS_VERSION) return null;
  const theme = typeof record.theme === "string" ? record.theme.trim() : "";
  if (!theme) return null;
  const themeKind = typeof record.themeKind === "string" ? record.themeKind : "theme";
  const styleSet = typeof record.styleSet === "string" && record.styleSet ? record.styleSet : null;
  return { theme, themeKind, styleSet };
}

/** Resolve the storage target: an explicitly injected storage (null = memory
 *  only), otherwise `localStorage` — resolved lazily so importing this module
 *  never touches blocked storage. */
function resolveStorage(provided?: DefaultsStorage | null): DefaultsStorage | null {
  if (provided !== undefined) return provided;
  try {
    return typeof localStorage === "undefined" ? null : localStorage;
  } catch {
    return null;
  }
}

/** The stored defaults, or null when none were ever set (or storage is
 *  blocked/corrupt — Word opens a fresh document on its factory template). */
export function readDocumentDefaults(storage?: DefaultsStorage | null): DocumentDefaults | null {
  try {
    return parseDocumentDefaults(
      resolveStorage(storage)?.getItem(DOCUMENT_DEFAULTS_STORAGE_KEY) ?? null,
    );
  } catch {
    return null;
  }
}

/** Persist the defaults (`Set as Default`). */
export function writeDocumentDefaults(
  defaults: DocumentDefaults,
  storage?: DefaultsStorage | null,
): void {
  try {
    resolveStorage(storage)?.setItem(
      DOCUMENT_DEFAULTS_STORAGE_KEY,
      JSON.stringify({ version: DOCUMENT_DEFAULTS_VERSION, ...defaults }),
    );
  } catch {
    // Quota exceeded / blocked storage — the set is a no-op this session.
  }
}

/** Reset to the factory defaults (no stored document formatting). The public
 *  reset path for `Set as Default` — there is no ribbon control for it (Word
 *  re-sets the default by setting a blank document's formatting), so callers
 *  invoke this programmatically. */
export function clearDocumentDefaults(storage?: DefaultsStorage | null): void {
  try {
    resolveStorage(storage)?.removeItem(DOCUMENT_DEFAULTS_STORAGE_KEY);
  } catch {
    // Blocked storage — nothing to clear.
  }
}

/** Read the current document's formatting as the `Set as Default` pair — the
 *  persisted theme (documentExtras.settings.theme) and the style set recorded
 *  when the last preset was applied. */
export function documentDefaultsOf(attrs: unknown): DocumentDefaults {
  const extras = (attrs as { documentExtras?: { settings?: Record<string, unknown> } } | undefined)
    ?.documentExtras;
  const settings = extras?.settings ?? {};
  const theme = settings.theme as { id?: unknown; kind?: unknown } | undefined;
  const styleSet = settings.styleSet;
  return {
    theme: typeof theme?.id === "string" && theme.id ? theme.id : "office",
    themeKind: typeof theme?.kind === "string" && theme.kind ? theme.kind : "theme",
    styleSet: typeof styleSet === "string" && styleSet ? styleSet : null,
  };
}

/** Re-apply the stored defaults to a freshly loaded document: the style set
 *  is a command (it rewrites the styles model), the theme already rode the
 *  doc attrs through `loadDoc`. Returns false when nothing applied. */
export function applyDocumentDefaults(
  editor: Editor | null | undefined,
  defaults: DocumentDefaults | null,
): boolean {
  if (!editor || !defaults?.styleSet) return false;
  return editor.commands["style-set"](defaults.styleSet) as boolean;
}

/** The blank document a new editor session starts from — carrying the stored
 *  theme (and the style set, which {@link applyDocumentDefaults} re-applies
 *  after the load) in documentExtras.settings so the load path resolves it
 *  exactly like an opened document's persisted theme. */
export function newDocumentJSON(defaults: DocumentDefaults | null): JSONContent {
  const json: JSONContent = { type: "doc", content: [{ type: "paragraph" }] };
  if (defaults) {
    json.attrs = {
      documentExtras: {
        settings: {
          theme: { id: defaults.theme, kind: defaults.themeKind },
          ...(defaults.styleSet ? { styleSet: defaults.styleSet } : {}),
        },
      },
    };
  }
  return json;
}
