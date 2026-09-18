#!/usr/bin/env node
// Fresh-process determinism probe for @docen/docx.
//
// Runs the same generation in a brand-new Node process so the parent test can
// prove that reproducibility does not depend on in-process state (counters,
// caches, the module graph). Prints the SHA-256 of the generated DOCX, or
// writes the raw bytes with `--dump` so the parent can inspect ZIP headers.
//
// Set DOCEN_FORCE_JS_DEFLATE=1 to force the pure-JS (fflate) ZIP writer — the
// browser/Deno fallback whose fixed-mtime normalization is otherwise shadowed
// by Node's native zeroed-timestamp writer.

import { createHash } from "node:crypto";

if (process.env.DOCEN_FORCE_JS_DEFLATE === "1") {
  globalThis[Symbol.for("docen.ooxml.force-js-deflate")] = true;
}

const { generateDOCXSync } = await import("@docen/docx");

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const model = {
  type: "doc",
  attrs: { core: { title: "Reproducible", creator: "docen" } },
  content: [
    {
      type: "paragraph",
      content: [
        { type: "text", text: "deterministic output " },
        { type: "image", attrs: { src: PNG, width: 1, height: 1 } },
        { type: "image", attrs: { src: PNG, width: 2, height: 2 } },
      ],
    },
    { type: "paragraph", content: [{ type: "text", text: "second paragraph" }] },
  ],
};

const bytes = generateDOCXSync(model);
if (process.argv[2] === "--dump") {
  process.stdout.write(Buffer.from(bytes));
} else {
  console.log(createHash("sha256").update(bytes).digest("hex"));
}
