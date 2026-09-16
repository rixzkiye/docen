import { describe, expect, it } from "vitest";

import {
  SETTINGS_STORAGE_KEY,
  SETTINGS_VERSION,
  createSettingsStore,
  defaultSettings,
  initialsFromName,
  resolveIdentity,
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

describe("defaultSettings", () => {
  it("carries the versioned defaults", () => {
    expect(defaultSettings()).toEqual({
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
          ordinalSuperscript: false,
        },
        markdownInput: true,
        proofingLanguage: "en-US",
      },
    });
  });

  it("builds a fresh object per call", () => {
    const a = defaultSettings();
    const b = defaultSettings();
    expect(a).not.toBe(b);
    expect(a).toEqual(b);
    a.writing.autocorrect.emDash = false;
    expect(b.writing.autocorrect.emDash).toBe(true);
  });
});

describe("createSettingsStore", () => {
  it("starts from defaults when nothing is stored", () => {
    const store = createSettingsStore(new MockStorage());
    expect(store.get()).toEqual(defaultSettings());
  });

  it("round-trips through storage under the versioned key", () => {
    const storage = new MockStorage();
    createSettingsStore(storage).update({
      identity: { name: "Demo Macro" },
      writing: { showHiddenText: true, markdownInput: false, proofingLanguage: "zh-CN" },
    });
    expect(storage.entries.has(SETTINGS_STORAGE_KEY)).toBe(true);

    const reloaded = createSettingsStore(storage).get();
    expect(reloaded.identity).toEqual({ name: "Demo Macro", initials: "DM" });
    expect(reloaded.writing.showHiddenText).toBe(true);
    expect(reloaded.writing.markdownInput).toBe(false);
    expect(reloaded.writing.proofingLanguage).toBe("zh-CN");
    expect(reloaded.writing.autocorrect).toEqual(defaultSettings().writing.autocorrect);
  });

  it("falls back to defaults on corrupt entries, never throwing", () => {
    const storage = new MockStorage();
    for (const raw of [
      "not json",
      "null",
      "[]",
      '"settings"',
      JSON.stringify({ version: 99, identity: { name: "Leftover" } }),
    ]) {
      storage.entries.set(SETTINGS_STORAGE_KEY, raw);
      expect(createSettingsStore(storage).get()).toEqual(defaultSettings());
    }
  });

  it("keeps valid fields around malformed siblings", () => {
    const storage = new MockStorage();
    storage.entries.set(
      SETTINGS_STORAGE_KEY,
      JSON.stringify({
        version: SETTINGS_VERSION,
        identity: { name: 42, initials: "DM" },
        writing: { showHiddenText: "yes", autocorrect: { emDash: false }, markdownInput: null },
      }),
    );
    const settings = createSettingsStore(storage).get();
    expect(settings.identity).toEqual({ name: "", initials: "DM" });
    expect(settings.writing.showHiddenText).toBe(false);
    expect(settings.writing.markdownInput).toBe(true);
    expect(settings.writing.autocorrect.emDash).toBe(false);
    expect(settings.writing.autocorrect.smartQuotes).toBe(true);
  });

  it("survives storage that throws on read or write", () => {
    const throwing: SettingsStorage = {
      getItem: () => {
        throw new Error("SecurityError");
      },
      setItem: () => {
        throw new Error("QuotaExceededError");
      },
    };
    const store = createSettingsStore(throwing);
    expect(store.get()).toEqual(defaultSettings());
    const events: unknown[] = [];
    store.onChange((s) => events.push(s));
    expect(store.update({ writing: { showHiddenText: true } }).writing.showHiddenText).toBe(true);
    expect(store.get().writing.showHiddenText).toBe(true);
    expect(events).toHaveLength(1);
  });

  it("runs in-memory when no storage is available", () => {
    const store = createSettingsStore(null);
    store.update({ identity: { name: "Solo" } });
    expect(store.get().identity).toEqual({ name: "Solo", initials: "S" });
  });

  it("notifies subscribers and stops after unsubscribe", () => {
    const store = createSettingsStore(null);
    const seen: string[] = [];
    const off = store.onChange((s) => seen.push(s.identity.name));
    store.update({ identity: { name: "Jane Doe" } });
    store.update({ identity: { name: "Jane Roe" } });
    off();
    store.update({ identity: { name: "Never Seen" } });
    expect(seen).toEqual(["Jane Doe", "Jane Roe"]);
  });

  it("does not notify or rewrite for a no-op patch", () => {
    const store = createSettingsStore(null);
    const before = store.get();
    let fired = 0;
    store.onChange(() => fired++);
    const after = store.update({
      identity: { name: before.identity.name },
      writing: { autocorrect: { emDash: before.writing.autocorrect.emDash } },
    });
    expect(after).toBe(before);
    expect(fired).toBe(0);
  });

  it("merges nested patches without dropping sibling keys", () => {
    const store = createSettingsStore(null);
    store.update({ writing: { autocorrect: { emDash: false, smartQuotes: false } } });
    store.update({ writing: { autocorrect: { emDash: true } } });
    expect(store.get().writing.autocorrect).toEqual({
      ...defaultSettings().writing.autocorrect,
      emDash: true,
      smartQuotes: false,
    });
  });

  it("derives initials from the name when initials are empty", () => {
    const store = createSettingsStore(null);
    store.update({ identity: { name: "John Doe" } });
    expect(store.get().identity).toEqual({ name: "John Doe", initials: "JD" });
    // Clearing the name clears the derived initials too.
    store.update({ identity: { name: "" } });
    expect(store.get().identity).toEqual({ name: "", initials: "" });
  });

  it("keeps explicitly stored initials when the name changes", () => {
    const store = createSettingsStore(null);
    store.update({ identity: { name: "John Doe", initials: "XY" } });
    store.update({ identity: { name: "Jane Roe" } });
    expect(store.get().identity).toEqual({ name: "Jane Roe", initials: "XY" });
  });
});

describe("initialsFromName", () => {
  it("takes the first letter of the first and last words", () => {
    expect(initialsFromName("John Doe")).toBe("JD");
    expect(initialsFromName("  Mary   Jane   Watson ")).toBe("MW");
  });

  it("uses a single letter for one-word names", () => {
    expect(initialsFromName("madonna")).toBe("M");
  });

  it("returns empty for empty or whitespace names", () => {
    expect(initialsFromName("")).toBe("");
    expect(initialsFromName("   ")).toBe("");
  });
});

describe("resolveIdentity", () => {
  const stored = { name: "Jane Roe", initials: "JR" };

  it("lets an explicit user attribute win, deriving its initials", () => {
    expect(resolveIdentity("John Doe", stored)).toEqual({ name: "John Doe", initials: "JD" });
  });

  it("falls back to the stored identity when the attribute is absent or blank", () => {
    expect(resolveIdentity(null, stored)).toBe(stored);
    expect(resolveIdentity(undefined, stored)).toBe(stored);
    expect(resolveIdentity("   ", stored)).toBe(stored);
  });
});
