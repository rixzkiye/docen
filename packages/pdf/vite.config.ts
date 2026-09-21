import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    // Two entries: the browser-safe writer/IR surface (main) and the Node
    // headless renderer (`@docen/pdf/node`) — the latter is the only one that
    // may import Node built-ins (`node:zlib` image decoding).
    entry: ["src/index.ts", "src/node.ts"],
  },
});
