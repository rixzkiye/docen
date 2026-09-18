import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: [
      "src/index.ts",
      "src/core.ts",
      "src/editor.ts",
      "src/layout/index.ts",
      "src/converters/**/*",
      "src/extensions/**/*",
      // Co-located vitest specs (*.spec.ts) live under src/ for name parity with
      // the module under test, but must not ship in dist (they import vitest).
      "!src/**/*.spec.ts",
      "!src/**/*.test.ts",
    ],
    // @docen/markdown is an internal, never-published workspace package (the
    // published surface stays self-contained): its code is bundled into the
    // converter entry that imports it, so `npm i @docen/docx` never has to
    // resolve it. It stays a devDependency for workspace builds/tests.
    deps: {
      alwaysBundle: ["@docen/markdown"],
    },
  },
});
