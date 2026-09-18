import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    // The subsetter is export-path-only (PDF/DOCX embedding). Keeping it out
    // of the main entry lets the editor load it on demand — the runtime
    // shaping path (rustybuzz + fontations) never needs it.
    entry: ["src/index.ts", "src/subsetter.ts"],
  },
});
