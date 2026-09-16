/**
 * Proofing Pack Manager: manages preinstalled and downloaded proofing packs
 * (dictionaries, grammar, thesaurus).
 */

import { englishWords } from "../spelling-dictionary";
import { indonesianWords } from "./id-words";
import type { ProofingLanguagePack, ProofingPackInfo } from "./types";

const PREINSTALLED_EN: ProofingLanguagePack = {
  id: "en",
  name: "English",
  spellWords: englishWords,
};

const PREINSTALLED_ID: ProofingLanguagePack = {
  id: "id",
  name: "Bahasa Indonesia",
  spellWords: indonesianWords,
};

export class ProofingPackManager {
  private installedPacks = new Map<string, ProofingLanguagePack>();

  constructor() {
    this.installedPacks.set("en", PREINSTALLED_EN);
    this.installedPacks.set("id", PREINSTALLED_ID);
  }

  getPack(langTag?: string): ProofingLanguagePack | null {
    if (!langTag) return this.installedPacks.get("en") ?? null;
    const clean = langTag.toLowerCase().trim();
    // Match exact tag e.g. "en-US" or prefix e.g. "en"
    if (this.installedPacks.has(clean)) return this.installedPacks.get(clean)!;
    const prefix = clean.split("-")[0]!;
    return this.installedPacks.get(prefix) ?? null;
  }

  isInstalled(id: string): boolean {
    const clean = id.toLowerCase().trim();
    return this.installedPacks.has(clean);
  }

  installPack(pack: ProofingLanguagePack): void {
    this.installedPacks.set(pack.id.toLowerCase(), pack);
  }

  uninstallPack(id: string): boolean {
    const clean = id.toLowerCase();
    // Do not allow uninstalling the default English pack
    if (clean === "en") return false;
    return this.installedPacks.delete(clean);
  }

  listInstalledPacks(): Array<{ id: string; name: string }> {
    return Array.from(this.installedPacks.values()).map((p) => ({
      id: p.id,
      name: p.name,
    }));
  }

  listAvailablePacks(): ProofingPackInfo[] {
    const catalog: ProofingPackInfo[] = [
      { id: "en", name: "English (US / UK)", installed: this.isInstalled("en"), size: "155 KB" },
      { id: "id", name: "Bahasa Indonesia", installed: this.isInstalled("id"), size: "45 KB" },
      { id: "fr", name: "Français", installed: this.isInstalled("fr"), size: "120 KB" },
      { id: "de", name: "Deutsch", installed: this.isInstalled("de"), size: "140 KB" },
      { id: "es", name: "Español", installed: this.isInstalled("es"), size: "115 KB" },
    ];
    return catalog;
  }
}

export const packManager = new ProofingPackManager();
