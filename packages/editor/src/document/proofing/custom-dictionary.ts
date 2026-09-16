/**
 * Custom Dictionary: persistent storage for user-added words (Add to Dictionary).
 * Persists to localStorage ("docen:custom-dictionary") with graceful in-memory fallback.
 */

const STORAGE_KEY = "docen:custom-dictionary";

export class CustomDictionaryManager {
  private words = new Set<string>();

  constructor(private readonly storageKey = STORAGE_KEY) {
    this.load();
  }

  load(): void {
    if (typeof localStorage === "undefined") return;
    try {
      const raw = localStorage.getItem(this.storageKey);
      if (raw) {
        const list = JSON.parse(raw);
        if (Array.isArray(list)) {
          this.words = new Set(list.map((w) => String(w).toLowerCase()));
        }
      }
    } catch {
      // Storage unavailable or disabled
    }
  }

  save(): void {
    if (typeof localStorage === "undefined") return;
    try {
      localStorage.setItem(this.storageKey, JSON.stringify([...this.words]));
    } catch {
      // Storage unavailable or full
    }
  }

  has(word: string): boolean {
    return this.words.has(word.toLowerCase());
  }

  add(word: string): void {
    const w = word.trim().toLowerCase();
    if (!w) return;
    this.words.add(w);
    this.save();
  }

  remove(word: string): void {
    const w = word.trim().toLowerCase();
    if (this.words.delete(w)) {
      this.save();
    }
  }

  all(): string[] {
    return [...this.words].sort();
  }

  clear(): void {
    this.words.clear();
    if (typeof localStorage !== "undefined") {
      try {
        localStorage.removeItem(this.storageKey);
      } catch {
        // Storage unavailable
      }
    }
  }

  size(): number {
    return this.words.size;
  }
}

export const customDictionary = new CustomDictionaryManager();
