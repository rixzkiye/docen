import { resolveLang, type RibbonTab } from "../../ui";
import {
  accessibilityGroup,
  arrangeGroup,
  btn,
  col,
  grid,
  group,
  menu,
  parsedItems,
  picker,
  resetPictureItems,
  sizeGroup,
  split,
  tab,
} from "./shared";

// Picture Format > Adjust: Word's preset grids flattened into menus. The
// labels assemble a localized lead word with the plain percent (the font-
// size ladder's pattern); the values are the picture-correction /
// picture-color / picture-transparency / picture-border command values.
export const PICTURE_PRESET_LEVELS = [40, 20, 0, -20, -40];

export const correctionItems = (): string => {
  const zh = resolveLang().toLowerCase().startsWith("zh");
  const pct = (n: number): string => `${n > 0 ? "+" : ""}${n}%`;
  const lead = zh ? "亮度" : "Brightness";
  const lead2 = zh ? "对比度" : "Contrast";
  return JSON.stringify([
    ...PICTURE_PRESET_LEVELS.map((n) => ({
      text: `${lead}: ${pct(n)}`,
      value: `bright:${n}`,
    })),
    { text: "-" },
    ...PICTURE_PRESET_LEVELS.map((n) => ({
      text: `${lead2}: ${pct(n)}`,
      value: `contrast:${n}`,
    })),
  ]);
};

export const pictureColorItems = (): string => {
  const zh = resolveLang().toLowerCase().startsWith("zh");
  const lead = zh ? "饱和度" : "Saturation";
  return JSON.stringify([
    { text: zh ? "没有重新着色" : "No Recolor", value: "none" },
    { text: "-" },
    ...[0, 33, 66, 100, 200].map((n) => ({
      text: `${lead}: ${n}%`,
      value: `saturation:${n}`,
    })),
    { text: "-" },
    {
      // Word's Set Transparent Color: arms the eyedropper (the next canvas
      // press samples the clicked picture's pixel) — a local host action,
      // not a command value.
      text: zh ? "设置透明色" : "Set Transparent Color",
      value: "transparent-pick",
      event: "picture-transparent-pick",
    },
  ]);
};

export const transparencyItems = (): string => {
  const zh = resolveLang().toLowerCase().startsWith("zh");
  const lead = zh ? "透明度" : "Transparency";
  return JSON.stringify(
    [0, 15, 30, 50, 65, 80, 95].map((n) => ({ text: `${lead}: ${n}%`, value: String(n) })),
  );
};

/** Word's Picture Tools — the contextual tab while a picture carries the
 *  selection. Groups mirror Word's Picture Format tab: Adjust (the pixel
 *  tools the engine doesn't model yet — greys until the image-tools batch),
 *  Picture Styles, Accessibility (Alt Text rides the size-and-position
 *  dialog), Arrange (the Layout tab's floating-drawing commands), and Size
 *  (the numeric Height/Width boxes + the crop split). Marked `contextual` so
 *  {@link ribbonTabs} excludes it from the static render; the host appends it
 *  via {@link buildContextualTab} as the selection enters/leaves a picture. */
export function pictureFormatTab(): RibbonTab {
  return {
    id: "picture-format",
    label: tab("picture-format"),
    contextual: true,
    groups: [
      // Word's Adjust group — the pixel tools the projection reads: the
      // correction/color/transparency presets write blipEffects; Reset
      // Picture discards every picture change (adjustments, border, effects,
      // crop — the split's second item also restores the natural size).
      // Remove Background and Artistic Effects stay greyed (no model for
      // them); Compress re-encodes the selected picture's pixels.
      group("picture-adjust", [
        btn("remove-background", "remove-background", { size: "large" }),
        col([
          grid([
            menu("corrections", "picture-correction", parsedItems(correctionItems())),
            menu("picture-color", "picture-color", parsedItems(pictureColorItems())),
            btn("artistic-effects", "artistic-effects"),
            menu("transparency", "picture-transparency", parsedItems(transparencyItems())),
            btn("compress-pictures", "compress-pictures"),
            btn("change-picture", "change-picture"),
            split("reset-picture", "reset-picture", resetPictureItems()),
          ]),
        ]),
      ]),
      group("picture-styles", [
        // The Quick Styles gallery + Effects/Layout — the styled picture
        // rendering is a later batch; Border is Word's outline panel (palette
        // + Weight / Dashes sub-views committing onto one outline object).
        btn("picture-styles", "picture-styles", { size: "large" }),
        col([
          grid([
            picker("border", "picture-border", "000000", {
              panel: "outline",
              withLabel: true,
            }),
            btn("text-effects", "picture-effects"),
            btn("smartart", "picture-layout"),
          ]),
        ]),
      ]),
      accessibilityGroup(),
      arrangeGroup(),
      sizeGroup("picture-size", true),
    ],
  };
}
