// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from "vitest";

import { DocenTranslatePane } from "./translate-pane";

/** Let FAST render before asserting on DOM */
async function settle(): Promise<void> {
  for (let i = 0; i < 5; i++) await new Promise((r) => setTimeout(r, 0));
}

const created: DocenTranslatePane[] = [];
afterEach(() => {
  while (created.length) created.pop()!.remove();
});

async function mount(): Promise<DocenTranslatePane> {
  const pane = new DocenTranslatePane();
  created.push(pane);
  document.body.append(pane);
  await settle();
  return pane;
}

describe("DocenTranslatePane", () => {
  describe("Tabs & Selection Translation", () => {
    it("renders selection tab by default and switches between tabs", async () => {
      const pane = await mount();
      expect(pane.tab).toBe("selection");
      expect(
        pane.shadowRoot!.querySelector("[part='tab-selection']")?.classList.contains("active"),
      ).toBe(true);
      expect(
        pane.shadowRoot!.querySelector("[part='panel-selection']")?.hasAttribute("hidden"),
      ).toBe(false);

      pane.selectTab("document");
      await settle();
      expect(pane.tab).toBe("document");
      expect(
        pane.shadowRoot!.querySelector("[part='tab-document']")?.classList.contains("active"),
      ).toBe(true);
      expect(
        pane.shadowRoot!.querySelector("[part='panel-document']")?.hasAttribute("hidden"),
      ).toBe(false);
      expect(
        pane.shadowRoot!.querySelector("[part='panel-selection']")?.hasAttribute("hidden"),
      ).toBe(true);
      expect(pane.shadowRoot!.querySelector("[part='translate-doc-btn']")).not.toBeNull();
    });

    it("translates source text and updates translation on setSelectionText", async () => {
      const pane = await mount();
      pane.setSelectionText("important document");
      await settle();

      expect(pane.sourceText).toBe("important document");
      expect(pane.translatedText).toBe("dokumen penting");
      expect(pane.shadowRoot!.querySelector("[part='target-text']")?.textContent).toContain(
        "dokumen penting",
      );
    });

    it("handles source input changes", async () => {
      const pane = await mount();
      const textarea = pane.shadowRoot!.querySelector<HTMLTextAreaElement>("[part='source-text']")!;
      textarea.value = "table of contents";
      textarea.dispatchEvent(new Event("input"));
      await settle();

      expect(pane.sourceText).toBe("table of contents");
      expect(pane.translatedText).toBe("daftar isi");
    });

    it("swaps languages when swap button is clicked", async () => {
      const pane = await mount();
      pane.sourceLang = "en";
      pane.targetLang = "id";
      pane.setSelectionText("report");
      await settle();

      expect(pane.translatedText).toBe("laporan");

      pane.swapLanguages();
      await settle();

      expect(pane.sourceLang).toBe("id");
      expect(pane.targetLang).toBe("en");
      expect(pane.sourceText).toBe("laporan");
      expect(pane.translatedText).toBe("report");
    });

    it("emits translate:insert event when Insert button is clicked", async () => {
      const pane = await mount();
      pane.setSelectionText("financial report");
      await settle();

      let emittedDetail = "";
      pane.addEventListener("translate:insert", ((e: CustomEvent<string>) => {
        emittedDetail = e.detail;
      }) as EventListener);

      const insertBtn = pane.shadowRoot!.querySelector<HTMLElement>("[part='insert-btn']")!;
      insertBtn.click();
      await settle();

      expect(emittedDetail).toBe("laporan keuangan");
    });
  });

  describe("Document Translation Tab", () => {
    it("dispatches translate:document event when Translate Document is clicked", async () => {
      const pane = await mount();
      pane.selectTab("document");
      pane.docSourceLang = "en";
      pane.docTargetLang = "id";
      await settle();

      let emittedDetail: { from?: string; to?: string } | null = null;
      pane.addEventListener("translate:document", ((
        e: CustomEvent<{ from: string; to: string }>,
      ) => {
        emittedDetail = e.detail;
      }) as EventListener);

      const btn = pane.shadowRoot!.querySelector<HTMLElement>("[part='translate-doc-btn']")!;
      btn.click();
      await settle();

      expect(pane.isTranslatingDoc).toBe(true);
      expect(emittedDetail).toEqual({ from: "en", to: "id" });

      pane.onDocumentTranslated(5);
      await settle();

      expect(pane.isTranslatingDoc).toBe(false);
      expect(pane.docStatusMessage).toContain("5 paragraphs");
      expect(pane.shadowRoot!.querySelector("[part='doc-status']")?.textContent).toContain(
        "5 paragraphs",
      );
    });
  });

  describe("Language Pack Status & On-Demand Download Simulator", () => {
    it("shows installed packs status in footer", async () => {
      const pane = await mount();
      const statusText = pane.shadowRoot!.querySelector("[part='pack-status-text']")?.textContent;
      expect(statusText).toContain("Packs:");
      expect(statusText).toContain("installed");
    });

    it("toggles available packs drawer and downloads on-demand pack", async () => {
      const pane = await mount();
      expect(pane.showPacksModal).toBe(false);
      expect(pane.shadowRoot!.querySelector("[part='pack-drawer']")?.hasAttribute("hidden")).toBe(
        true,
      );

      pane.togglePacksModal();
      await settle();

      expect(pane.showPacksModal).toBe(true);
      expect(pane.shadowRoot!.querySelector("[part='pack-drawer']")?.hasAttribute("hidden")).toBe(
        false,
      );

      // Download a pack
      await pane.downloadPack("es");
      await settle();

      expect(pane.packManager.isInstalled("es")).toBe(true);
      expect(pane.supportedLanguages.some((l) => l.code === "es")).toBe(true);

      // Clean up downloaded pack
      pane.packManager.uninstallPack("es");
    });
  });
});
