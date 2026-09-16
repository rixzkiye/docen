import { describe, expect, it } from "vitest";

import {
  AUTOCORRECT_TABLE_VERSION,
  SETTINGS_STORAGE_KEY,
  SETTINGS_VERSION,
  createSettingsStore,
  defaultSettings,
  type AutocorrectTable,
  type SettingsStorage,
} from "./settings";

/** Minimal `Storage` stand-in — the store only needs getItem/setItem. */
class MockStorage implements SettingsStorage {
  readonly entries = new Map<string, string>();
  getItem(key: string): string | null {
    return this.entries.get(key) ?? null;
  }
  setItem(key: string, value: string): void {
    this.entries.set(key, value);
  }
}

const table = (replacements: AutocorrectTable["replacements"], exceptions: string[] = []) => ({
  version: AUTOCORRECT_TABLE_VERSION,
  replacements,
  exceptions,
});

describe("AutoCorrect table persistence", () => {
  it("starts with no stored table (the built-ins apply)", () => {
    const settings = createSettingsStore(new MockStorage()).get();
    expect(settings.writing.autocorrect.table).toBeUndefined();
    expect(defaultSettings().writing.autocorrect.table).toBeUndefined();
  });

  it("round-trips a user table under the versioned settings key", () => {
    const storage = new MockStorage();
    createSettingsStore(storage).update({
      writing: {
        autocorrect: { table: table([{ from: "teh", to: "teh!" }], ["teh"]) },
      },
    });
    expect(storage.entries.has(SETTINGS_STORAGE_KEY)).toBe(true);
    const reloaded = createSettingsStore(storage).get();
    expect(reloaded.writing.autocorrect.table).toEqual(
      table([{ from: "teh", to: "teh!" }], ["teh"]),
    );
  });

  it("keeps an explicitly emptied table empty (no built-in fallback)", () => {
    const storage = new MockStorage();
    createSettingsStore(storage).update({ writing: { autocorrect: { table: table([]) } } });
    expect(createSettingsStore(storage).get().writing.autocorrect.table).toEqual(table([]));
  });

  it("drops a corrupt or foreign-version table back to the built-ins", () => {
    const storage = new MockStorage();
    for (const bad of [
      { version: 99, replacements: [{ from: "a", to: "b" }], exceptions: [] },
      { version: AUTOCORRECT_TABLE_VERSION, replacements: "nope", exceptions: [] },
      { version: AUTOCORRECT_TABLE_VERSION },
      "garbage",
      42,
      null,
    ]) {
      storage.entries.set(
        SETTINGS_STORAGE_KEY,
        JSON.stringify({
          version: SETTINGS_VERSION,
          writing: { autocorrect: { table: bad } },
        }),
      );
      expect(createSettingsStore(storage).get().writing.autocorrect.table).toBeUndefined();
    }
  });

  it("skips malformed entries and words inside a valid table", () => {
    const storage = new MockStorage();
    storage.entries.set(
      SETTINGS_STORAGE_KEY,
      JSON.stringify({
        version: SETTINGS_VERSION,
        writing: {
          autocorrect: {
            table: {
              version: AUTOCORRECT_TABLE_VERSION,
              replacements: [
                { from: "ok", to: "yes" },
                { from: 42, to: "x" },
                { from: "", to: "y" },
                { from: "z", to: "" },
                null,
              ],
              exceptions: ["a", 7, "", "  b "],
            },
          },
        },
      }),
    );
    expect(createSettingsStore(storage).get().writing.autocorrect.table).toEqual(
      table([{ from: "ok", to: "yes" }], ["a", "b"]),
    );
  });

  it("preserves the table across unrelated boolean patches", () => {
    const store = createSettingsStore(null);
    store.update({ writing: { autocorrect: { table: table([{ from: "x", to: "y" }], ["x"]) } } });
    store.update({ writing: { autocorrect: { emDash: false } } });
    expect(store.get().writing.autocorrect.emDash).toBe(false);
    expect(store.get().writing.autocorrect.table).toEqual(table([{ from: "x", to: "y" }], ["x"]));
  });

  it("replaces the table wholesale on a table patch", () => {
    const store = createSettingsStore(null);
    store.update({ writing: { autocorrect: { table: table([{ from: "a", to: "b" }], ["a"]) } } });
    store.update({ writing: { autocorrect: { table: table([{ from: "c", to: "d" }]) } } });
    expect(store.get().writing.autocorrect.table).toEqual(table([{ from: "c", to: "d" }]));
  });
});
