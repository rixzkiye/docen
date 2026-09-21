// `docen/pdf` subpath — full facade over @docen/pdf, the headless PDF engine.
//
// Re-exports everything: the headless server-side renderer (`renderPdf`),
// vector scene serializer, low-level PDF primitives, font embedding,
// structure tagging, PDF/A-2b / PDF/UA-1 conformance, and image decoding.
export * from "@docen/pdf";
export * from "@docen/pdf/node";
export * from "./pdf-adapter";
