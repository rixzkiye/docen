/**
 * Context cursor pack for Word parity.
 */

/** Word's margin selection cursor: right-pointing (northeast) arrow for line/para selection */
export const MARGIN_SELECTION_CURSOR = (() => {
  const svg =
    '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">' +
    '<path d="M18 2 L18 17 L14 13 L10 21 L7 19.5 L11 12 L6 12 Z" fill="#ffffff" stroke="#000000" stroke-width="1.5" stroke-linejoin="round"/>' +
    "</svg>";
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") 18 2, default`;
})();

/** Word's rectangular block selection cursor */
export const BLOCK_SELECT_CURSOR = "crosshair";
