// Node-only entry (`@docen/pdf/node`): headless rendering, image decoding via
// `node:zlib`, and the leafer-free scene serializer. Kept out of the main
// browser-safe entry so the editor bundle never pulls Node built-ins.
export * from "./render";
export * from "./node-image";
export * from "./node-scene";
