import type { RibbonTab } from "../../ui";
import {
  accessibilityGroup,
  arrangeGroup,
  btn,
  col,
  grid,
  group,
  menu,
  opt,
  parsedItems,
  picker,
  sizeGroup,
  tab,
} from "./shared";

// Word's Shape Effects menu — the Shadow section (Off + the eight outer
// presets in Word's gallery ring: lower-right → bottom → … → right). The
// inner/perspective shadows and the other effect families have no painter
// mapping yet and grey out honestly.
export const shapeEffectItems = (): string =>
  JSON.stringify([
    { text: opt("no-shadow"), value: "none" },
    { text: opt("shadow-lower-right"), value: "shadow-lower-right" },
    { text: opt("shadow-bottom"), value: "shadow-bottom" },
    { text: opt("shadow-lower-left"), value: "shadow-lower-left" },
    { text: opt("shadow-left"), value: "shadow-left" },
    { text: opt("shadow-upper-left"), value: "shadow-upper-left" },
    { text: opt("shadow-top"), value: "shadow-top" },
    { text: opt("shadow-upper-right"), value: "shadow-upper-right" },
    { text: opt("shadow-right"), value: "shadow-right" },
    { text: opt("glow"), disabled: true },
    { text: opt("soft-edges"), disabled: true },
    { text: opt("bevel"), disabled: true },
    { text: opt("rotation-3d"), disabled: true },
  ]);

// Word's Text Direction menu (Drawing Tools > Text): the two rotated
// layouts the renderer maps plus horizontal (the cleared state). Stacked
// needs per-glyph upright layout — greyed, a registered gap.
export const textDirectionItems = (): string =>
  JSON.stringify([
    { text: opt("text-horizontal"), value: "horizontal" },
    { text: opt("text-rotate-90"), value: "vertical" },
    { text: opt("text-rotate-270"), value: "vertical270" },
    { text: opt("text-stacked"), disabled: true },
  ]);

/** Word's Drawing Tools — the contextual tab while a shape or a group
 *  carries the selection. Shape Styles is live (fill picker, outline panel,
 *  the Shadow effects menu); Insert Shapes and the WordArt group grey until
 *  their commands land; Accessibility, Arrange, and Size are live (their
 *  commands accept all three drawing kinds). Marked `contextual` like the
 *  picture tab. */
export function shapeFormatTab(): RibbonTab {
  return {
    id: "shape-format",
    label: tab("shape-format"),
    contextual: true,
    groups: [
      // The shape gallery + Edit Shape / Draw Text Box — shape insertion is a
      // later batch, so the whole group greys.
      group("insert-shapes", [
        btn("ink-shape", "shapes", { size: "large" }),
        col([grid([btn("format-painter", "edit-shape"), btn("text-box", "text-box")])]),
      ]),
      group("shape-styles", [
        // Fill is the theme palette picker; Outline is Word's outline panel
        // (palette + Weight / Dashes); Effects is the Shadow menu — all three
        // commit onto the shape's attrs.
        btn("picture-styles", "shape-styles", { size: "large" }),
        col([
          grid([
            picker("page-color", "shape-fill", "4472C4", { withLabel: true }),
            picker("page-border", "shape-outline", "000000", {
              panel: "outline",
              withLabel: true,
            }),
            menu("artistic-effects", "shape-effects", parsedItems(shapeEffectItems())),
          ]),
        ]),
      ]),
      group("wordart-styles", [
        btn("picture-styles", "wordart-styles", { size: "large" }),
        col([
          grid([
            btn("font-color", "text-fill"),
            btn("page-border", "text-outline"),
            btn("text-effects", "text-effects"),
          ]),
        ]),
      ]),
      group("text", [
        col([
          grid([
            menu("text-direction", "shape-text-direction", parsedItems(textDirectionItems())),
            btn("align-distribute", "align-text"),
            btn("text-link", "text-link"),
          ]),
        ]),
      ]),
      accessibilityGroup(),
      arrangeGroup(),
      sizeGroup("shape-size", false),
    ],
  };
}
