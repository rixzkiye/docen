/**
 * Persisted editor settings + the user identity store (`<docen-document>`).
 *
 * One small typed document under a versioned localStorage key
 * ({@link SETTINGS_STORAGE_KEY}), shared by every host element on the page:
 * identity (the Word Options → General "User name"/"Initials" pair) and the
 * writing toggles the lanes consume — D1 autocorrect (its options + the user
 * replacement table), B4 proofing language, C1 hidden-text display (the
 * `<docen-document>` projection reads `writing.showHiddenText` per render and
 * re-renders when it flips). This module only stores the options; the document
 * behavior rides the engine.
 *
 * Storage is treated as a cache: a corrupt entry, a foreign version, blocked
 * storage (private mode / SecurityError) or a throwing write all fall back to
 * the in-memory value, never to an exception.
 */

/** The storage surface the store needs — structural subset of `Storage`, so a
 *  headless run (or a test) can inject a fake. */
export interface SettingsStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

/** The Word Options "User name" / "Initials" pair. */
export interface IdentitySettings {
  name: string;
  initials: string;
}

/** One user-editable replacement-table entry (AutoCorrect Options → table). */
export interface AutocorrectReplacement {
  from: string;
  to: string;
}

/** The table shape version — a stored table with another version is ignored
 *  and the built-in defaults apply. */
export const AUTOCORRECT_TABLE_VERSION = 1;

/** The versioned user table behind AutoCorrect Options: the complete
 *  replacement list (deleting an entry removes it — absent `table` means the
 *  built-in defaults) plus the no-correct exceptions list. */
export interface AutocorrectTable {
  version: number;
  replacements: AutocorrectReplacement[];
  exceptions: string[];
}

/** AutoCorrect options (D1) — persisted; the rule engine consumes them. */
export interface AutocorrectSettings {
  smartQuotes: boolean;
  emDash: boolean;
  ellipsis: boolean;
  hyperlinkAutoformat: boolean;
  capitalizeFirstLetter: boolean;
  ordinalSuperscript: boolean;
  /** User-edited table; absent = the built-in defaults. */
  table?: AutocorrectTable;
}

/** Writing options (stored for B4/C1 — no behavior is wired here). */
export interface WritingSettings {
  showHiddenText: boolean;
  autocorrect: AutocorrectSettings;
  markdownInput: boolean;
  proofingLanguage: string;
}

/** The persisted settings document. */
export interface DocenSettings {
  version: number;
  identity: IdentitySettings;
  writing: WritingSettings;
}

/** A merge patch — every level optional; absent groups/keys keep their value. */
export interface SettingsPatch {
  identity?: Partial<IdentitySettings>;
  writing?: Partial<Omit<WritingSettings, "autocorrect">> & {
    autocorrect?: Partial<AutocorrectSettings>;
  };
}

/** Store change subscriber. */
export type SettingsListener = (settings: DocenSettings) => void;

/** The versioned localStorage key. */
export const SETTINGS_STORAGE_KEY = "docen.settings.v1";

/** The schema version persisted alongside the settings. */
export const SETTINGS_VERSION = 1;

/** The default settings document (a fresh object each call — callers own it). */
export function defaultSettings(): DocenSettings {
  return {
    version: SETTINGS_VERSION,
    identity: { name: "", initials: "" },
    writing: {
      showHiddenText: false,
      autocorrect: {
        smartQuotes: true,
        emDash: true,
        ellipsis: true,
        hyperlinkAutoformat: true,
        capitalizeFirstLetter: true,
        // Word leaves ordinals (1st → 1ˢᵗ) off by default.
        ordinalSuperscript: false,
      },
      markdownInput: true,
      proofingLanguage: "en-US",
    },
  };
}

/** Name → initials (Word's avatar initials): the first letter of the name's
 *  first and last words, upper-cased ("John Doe" → "JD", "madonna" → "M"). */
export function initialsFromName(name: string): string {
  const words = name.trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return "";
  const first = words[0]!;
  const last = words[words.length - 1]!;
  return (first.charAt(0) + (last === first ? "" : last.charAt(0))).toUpperCase();
}

/**
 * The effective identity for a `<docen-document>`: the `user` attribute wins
 * while explicitly set (its initials derive from it); otherwise the stored
 * identity — custom initials included.
 */
export function resolveIdentity(
  attributeUser: string | null | undefined,
  stored: IdentitySettings,
): IdentitySettings {
  const name = attributeUser?.trim();
  if (!name) return stored;
  return { name, initials: initialsFromName(name) };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function pickString(value: unknown, fallback: string): string {
  return typeof value === "string" ? value : fallback;
}

function pickBoolean(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

/** Defensive parse of the user AutoCorrect table: a wrong/missing version or a
 *  non-array replacement list drops the whole table (built-in defaults apply);
 *  malformed entries/words within a valid table are skipped. */
function parseAutocorrectTable(value: unknown): AutocorrectTable | undefined {
  if (!isRecord(value) || value.version !== AUTOCORRECT_TABLE_VERSION) return undefined;
  if (!Array.isArray(value.replacements)) return undefined;
  const replacements: AutocorrectReplacement[] = [];
  for (const entry of value.replacements) {
    if (!isRecord(entry)) continue;
    const from = typeof entry.from === "string" ? entry.from.trim() : "";
    const to = typeof entry.to === "string" ? entry.to.trim() : "";
    if (from && to) replacements.push({ from, to });
  }
  const exceptions = Array.isArray(value.exceptions)
    ? value.exceptions.flatMap((word) => {
        const trimmed = typeof word === "string" ? word.trim() : "";
        return trimmed ? [trimmed] : [];
      })
    : [];
  return { version: AUTOCORRECT_TABLE_VERSION, replacements, exceptions };
}

/** Defensive parse: any malformed entry — bad JSON, wrong shape, foreign
 *  version — yields the defaults for the offending fields (never throws). */
function parseSettings(raw: string | null): DocenSettings {
  const base = defaultSettings();
  if (!raw) return base;
  let data: unknown;
  try {
    data = JSON.parse(raw);
  } catch {
    return base;
  }
  if (!isRecord(data) || data.version !== SETTINGS_VERSION) return base;
  const identity = isRecord(data.identity) ? data.identity : {};
  const writing = isRecord(data.writing) ? data.writing : {};
  const autocorrect = isRecord(writing.autocorrect) ? writing.autocorrect : {};
  const table = parseAutocorrectTable(autocorrect.table);
  return {
    version: SETTINGS_VERSION,
    identity: {
      name: pickString(identity.name, base.identity.name),
      initials: pickString(identity.initials, base.identity.initials),
    },
    writing: {
      showHiddenText: pickBoolean(writing.showHiddenText, base.writing.showHiddenText),
      autocorrect: {
        smartQuotes: pickBoolean(autocorrect.smartQuotes, base.writing.autocorrect.smartQuotes),
        emDash: pickBoolean(autocorrect.emDash, base.writing.autocorrect.emDash),
        ellipsis: pickBoolean(autocorrect.ellipsis, base.writing.autocorrect.ellipsis),
        hyperlinkAutoformat: pickBoolean(
          autocorrect.hyperlinkAutoformat,
          base.writing.autocorrect.hyperlinkAutoformat,
        ),
        capitalizeFirstLetter: pickBoolean(
          autocorrect.capitalizeFirstLetter,
          base.writing.autocorrect.capitalizeFirstLetter,
        ),
        ordinalSuperscript: pickBoolean(
          autocorrect.ordinalSuperscript,
          base.writing.autocorrect.ordinalSuperscript,
        ),
        ...(table ? { table } : {}),
      },
      markdownInput: pickBoolean(writing.markdownInput, base.writing.markdownInput),
      proofingLanguage: pickString(writing.proofingLanguage, base.writing.proofingLanguage),
    },
  };
}

/** Merge a patch onto `current`: an empty identity name clears the identity,
 *  otherwise empty initials derive from the (patched) name. Canonical key
 *  order is what `update` compares on. */
function mergeSettings(current: DocenSettings, patch: SettingsPatch): DocenSettings {
  const name = (patch.identity?.name ?? current.identity.name).trim();
  const initials = (patch.identity?.initials ?? current.identity.initials).trim();
  const w = patch.writing ?? {};
  const a = w.autocorrect ?? {};
  const table = a.table ?? current.writing.autocorrect.table;
  return {
    version: SETTINGS_VERSION,
    identity: {
      name,
      initials: name ? initials || initialsFromName(name) : "",
    },
    writing: {
      showHiddenText: w.showHiddenText ?? current.writing.showHiddenText,
      autocorrect: {
        smartQuotes: a.smartQuotes ?? current.writing.autocorrect.smartQuotes,
        emDash: a.emDash ?? current.writing.autocorrect.emDash,
        ellipsis: a.ellipsis ?? current.writing.autocorrect.ellipsis,
        hyperlinkAutoformat:
          a.hyperlinkAutoformat ?? current.writing.autocorrect.hyperlinkAutoformat,
        capitalizeFirstLetter:
          a.capitalizeFirstLetter ?? current.writing.autocorrect.capitalizeFirstLetter,
        ordinalSuperscript: a.ordinalSuperscript ?? current.writing.autocorrect.ordinalSuperscript,
        ...(table ? { table } : {}),
      },
      markdownInput: w.markdownInput ?? current.writing.markdownInput,
      proofingLanguage: w.proofingLanguage ?? current.writing.proofingLanguage,
    },
  };
}

/** Resolve the storage target: an explicitly injected storage (null = memory
 *  only), otherwise the global `localStorage` — resolved lazily so importing
 *  this module never touches blocked storage. */
function resolveStorage(provided?: SettingsStorage | null): SettingsStorage | null {
  if (provided !== undefined) return provided;
  try {
    return typeof localStorage === "undefined" ? null : localStorage;
  } catch {
    return null;
  }
}

/** A settings store instance ({@link createSettingsStore}) — `get` loads once,
 *  `update` merges/persists/notifies, `onChange` subscribes. */
export interface SettingsStore {
  /** Current settings (loaded from storage on first call; treat the returned
   *  document as read-only — mutate through `update`). */
  get(): DocenSettings;
  /** Merge a patch; persist + notify only when something changed. */
  update(patch: SettingsPatch): DocenSettings;
  /** Subscribe to changes; returns the unsubscribe function. */
  onChange(listener: SettingsListener): () => void;
}

/**
 * Build an isolated settings store. Tests (and consumers that want their own
 * persistence channel) pass a storage; `null` forces the in-memory fallback.
 */
export function createSettingsStore(storage?: SettingsStorage | null): SettingsStore {
  const listeners = new Set<SettingsListener>();
  let cached: DocenSettings | undefined;

  const read = (): DocenSettings => {
    if (!cached) {
      try {
        cached = parseSettings(resolveStorage(storage)?.getItem(SETTINGS_STORAGE_KEY) ?? null);
      } catch {
        // Blocked storage — the session runs on defaults.
        cached = defaultSettings();
      }
    }
    return cached;
  };

  const update = (patch: SettingsPatch): DocenSettings => {
    const current = read();
    const next = mergeSettings(current, patch);
    // Canonical construction keeps the serialized form comparable — an
    // unchanged patch is a no-op (no write, no notification).
    if (JSON.stringify(next) === JSON.stringify(current)) return current;
    cached = next;
    try {
      resolveStorage(storage)?.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(next));
    } catch {
      // Quota exceeded / blocked storage — the session keeps the value.
    }
    for (const listener of listeners) listener(next);
    return next;
  };

  const onChange = (listener: SettingsListener): (() => void) => {
    listeners.add(listener);
    return () => {
      listeners.delete(listener);
    };
  };

  return { get: read, update, onChange };
}

/** The app-wide store — `localStorage`-backed, shared by every host element. */
const store = createSettingsStore();

/** Current settings (in-memory cache; first read loads from storage). */
export function getSettings(): DocenSettings {
  return store.get();
}

/** Merge a settings patch into the app-wide store (persist + notify). */
export function updateSettings(patch: SettingsPatch): DocenSettings {
  return store.update(patch);
}

/** Subscribe to app-wide settings changes; returns the unsubscribe function. */
export function onSettingsChange(listener: SettingsListener): () => void {
  return store.onChange(listener);
}
