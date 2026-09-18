import { describe, expect, it } from "vitest";

import { TranslationEngine, detectLanguage, preserveCasing, translateText } from "./engine";
import { TranslationPackManager } from "./pack-manager";

describe("Translation Engine & Packs", () => {
  describe("preserveCasing", () => {
    it("preserves ALL UPPERCASE", () => {
      expect(preserveCasing("DOCUMENT", "dokumen")).toBe("DOKUMEN");
      expect(preserveCasing("TABLE OF CONTENTS", "daftar isi")).toBe("DAFTAR ISI");
    });

    it("preserves Title / Capitalized Case", () => {
      expect(preserveCasing("Document", "dokumen")).toBe("Dokumen");
      expect(preserveCasing("Important", "penting")).toBe("Penting");
    });

    it("preserves lowercase", () => {
      expect(preserveCasing("document", "dokumen")).toBe("dokumen");
    });
  });

  describe("detectLanguage", () => {
    it("detects English source text", () => {
      expect(detectLanguage("This is an important document and report.")).toBe("en");
      expect(detectLanguage("Please save the file")).toBe("en");
    });

    it("detects Indonesian source text", () => {
      expect(detectLanguage("Ini adalah dokumen penting dan laporan.")).toBe("id");
      expect(detectLanguage("Harap simpan berkas ini")).toBe("id");
    });
  });

  describe("Single word lookup", () => {
    const engine = new TranslationEngine();

    it("translates common document words en -> id", () => {
      expect(translateText("document", "en", "id", engine)).toBe("dokumen");
      expect(translateText("paragraph", "en", "id", engine)).toBe("paragraf");
      expect(translateText("table", "en", "id", engine)).toBe("tabel");
      expect(translateText("column", "en", "id", engine)).toBe("kolom");
      expect(translateText("save", "en", "id", engine)).toBe("simpan");
      expect(translateText("close", "en", "id", engine)).toBe("tutup");
    });

    it("translates common document words id -> en", () => {
      expect(translateText("dokumen", "id", "en", engine)).toBe("document");
      expect(translateText("paragraf", "id", "en", engine)).toBe("paragraph");
      expect(translateText("tabel", "id", "en", engine)).toBe("table");
      expect(translateText("simpan", "id", "en", engine)).toBe("save");
      expect(translateText("tutup", "id", "en", engine)).toBe("close");
    });

    it("preserves casing on single word translations", () => {
      expect(translateText("Document", "en", "id", engine)).toBe("Dokumen");
      expect(translateText("DOCUMENT", "en", "id", engine)).toBe("DOKUMEN");
      expect(translateText("Dokumen", "id", "en", engine)).toBe("Document");
      expect(translateText("DOKUMEN", "id", "en", engine)).toBe("DOCUMENT");
    });

    it("leaves unknown words untranslated and preserves case", () => {
      expect(translateText("Kubernetes cluster", "en", "id", engine)).toBe("Kubernetes cluster");
      expect(translateText("Xylophone999", "en", "id", engine)).toBe("Xylophone999");
    });
  });

  describe("Phrase matching", () => {
    const engine = new TranslationEngine();

    it("matches multi-word phrases en -> id", () => {
      expect(translateText("table of contents", "en", "id", engine)).toBe("daftar isi");
      expect(translateText("executive summary", "en", "id", engine)).toBe("ringkasan eksekutif");
      expect(translateText("terms and conditions", "en", "id", engine)).toBe(
        "syarat dan ketentuan",
      );
      expect(translateText("privacy policy", "en", "id", engine)).toBe("kebijakan privasi");
      expect(translateText("annual report", "en", "id", engine)).toBe("laporan tahunan");
      expect(translateText("track changes", "en", "id", engine)).toBe("lacak perubahan");
      expect(translateText("user manual", "en", "id", engine)).toBe("panduan pengguna");
    });

    it("matches multi-word phrases id -> en", () => {
      expect(translateText("daftar isi", "id", "en", engine)).toBe("table of contents");
      expect(translateText("ringkasan eksekutif", "id", "en", engine)).toBe("executive summary");
      expect(translateText("syarat dan ketentuan", "id", "en", engine)).toBe(
        "terms and conditions",
      );
      expect(translateText("kebijakan privasi", "id", "en", engine)).toBe("privacy policy");
      expect(translateText("laporan tahunan", "id", "en", engine)).toBe("annual report");
      expect(translateText("lacak perubahan", "id", "en", engine)).toBe("track changes");
    });

    it("preserves casing in multi-word phrases", () => {
      expect(translateText("Table of contents", "en", "id", engine)).toBe("Daftar isi");
      expect(translateText("TABLE OF CONTENTS", "en", "id", engine)).toBe("DAFTAR ISI");
      expect(translateText("Daftar isi", "id", "en", engine)).toBe("Table of contents");
      expect(translateText("DAFTAR ISI", "id", "en", engine)).toBe("TABLE OF CONTENTS");
    });
  });

  describe("Grammar transformations & sentence structure", () => {
    const engine = new TranslationEngine();

    it("swaps English [ADJECTIVE] [NOUN] to Indonesian [NOUN] [ADJECTIVE]", () => {
      expect(translateText("blue car", "en", "id", engine)).toBe("mobil biru");
      expect(translateText("important document", "en", "id", engine)).toBe("dokumen penting");
      expect(translateText("new project", "en", "id", engine)).toBe("proyek baru");
      expect(translateText("Important document", "en", "id", engine)).toBe("Dokumen penting");
      expect(translateText("Blue car", "en", "id", engine)).toBe("Mobil biru");
    });

    it("swaps Indonesian [NOUN] [ADJECTIVE] to English [ADJECTIVE] [NOUN]", () => {
      expect(translateText("mobil biru", "id", "en", engine)).toBe("blue car");
      expect(translateText("dokumen penting", "id", "en", engine)).toBe("important document");
      expect(translateText("proyek baru", "id", "en", engine)).toBe("new project");
      expect(translateText("Dokumen penting", "id", "en", engine)).toBe("Important document");
    });

    it("translates Indonesian reduplicated plurals to English plurals", () => {
      expect(translateText("buku-buku", "id", "en", engine)).toBe("books");
      expect(translateText("dokumen-dokumen", "id", "en", engine)).toBe("documents");
    });

    it("preserves punctuation, whitespace, and sentence structure", () => {
      const input = "Please save the document! Is the report ready? (Yes, it is).";
      const result = translateText(input, "en", "id", engine);
      expect(result).toContain("!");
      expect(result).toContain("?");
      expect(result).toContain("(");
      expect(result).toContain(")");
      expect(result).toContain(".");
      expect(result).toContain("dokumen");
      expect(result).toContain("laporan");
    });

    it("works with auto-detect", () => {
      const enRes = engine.translate("This is a document", { from: "auto", to: "id" });
      expect(enRes.detectedSourceLang).toBe("en");
      expect(enRes.translatedText).toContain("dokumen");

      const idRes = engine.translate("Ini adalah dokumen", { from: "auto", to: "en" });
      expect(idRes.detectedSourceLang).toBe("id");
      expect(idRes.translatedText).toContain("document");
    });
  });

  describe("TranslationPackManager operations", () => {
    it("has preinstalled en-id and id-en language packs", () => {
      const pm = new TranslationPackManager();
      expect(pm.isInstalled("en-id")).toBe(true);
      expect(pm.isInstalled("id-en")).toBe(true);
      expect(pm.isInstalled("en")).toBe(true);
      expect(pm.isInstalled("id")).toBe(true);
      expect(pm.getPack("en", "id")).not.toBeNull();
      expect(pm.getPack("id", "en")).not.toBeNull();
    });

    it("cannot uninstall built-in packs", () => {
      const pm = new TranslationPackManager();
      expect(pm.uninstallPack("en-id")).toBe(false);
      expect(pm.uninstallPack("en")).toBe(false);
      expect(pm.isInstalled("en-id")).toBe(true);
    });

    it("lists available packs including fr, de, es, zh", () => {
      const pm = new TranslationPackManager();
      const catalog = pm.listAvailablePacks();
      const ids = catalog.map((p) => p.id);
      expect(ids).toContain("en-id");
      expect(ids).toContain("fr");
      expect(ids).toContain("de");
      expect(ids).toContain("es");
      expect(ids).toContain("zh");
    });

    it("downloads and installs on-demand language pack", async () => {
      const pm = new TranslationPackManager();
      expect(pm.isInstalled("fr")).toBe(false);

      const pack = await pm.downloadPack("fr");
      expect(pack.id).toBe("en-fr");
      expect(pm.isInstalled("fr")).toBe(true);

      const engine = new TranslationEngine(pm);
      const res = translateText("document", "en", "fr", engine);
      expect(res).toBe("document");

      const res2 = translateText("file", "en", "fr", engine);
      expect(res2).toBe("fichier");

      // Verify uninstallation of downloaded pack
      expect(pm.uninstallPack("fr")).toBe(true);
      expect(pm.isInstalled("fr")).toBe(false);
    });

    it("returns supported languages including newly installed packs", async () => {
      const pm = new TranslationPackManager();
      const initialLangs = pm.getSupportedLanguages().map((l) => l.code);
      expect(initialLangs).toContain("en");
      expect(initialLangs).toContain("id");

      await pm.downloadPack("de");
      const updatedLangs = pm.getSupportedLanguages().map((l) => l.code);
      expect(updatedLangs).toContain("de");
    });
  });
});
