import { opt } from "./shared";

/**
 * Home → Font → Text Effects and Typography: Word's effect gallery as nested
 * submenus (Outline / Shadow / Reflection / Glow / Bevel / 3-D Rotation).
 * Each entry carries the `<kind>:<preset>` token the `text-effects` command
 * applies; `:0` clears that family and `:options` opens the custom dialog.
 * Labels live under `ribbon.opt.te-*` (en + zh).
 */
export const textEffectsItems = (): string =>
  JSON.stringify([
    {
      text: opt("te-outline"),
      children: [
        { text: opt("te-no-outline"), event: "text-effects", value: "outline:0" },
        { text: opt("te-outline-black"), event: "text-effects", value: "outline:1" },
        { text: opt("te-outline-blue"), event: "text-effects", value: "outline:2" },
        { text: opt("te-outline-red"), event: "text-effects", value: "outline:3" },
        { text: opt("te-outline-white"), event: "text-effects", value: "outline:4" },
        { text: opt("te-outline-gold"), event: "text-effects", value: "outline:5" },
        { text: opt("te-more-options"), event: "text-effects", value: "outline:options" },
      ],
    },
    {
      text: opt("te-shadow"),
      children: [
        { text: opt("te-no-shadow"), event: "text-effects", value: "shadow:0" },
        { text: opt("te-shadow-offset-br"), event: "text-effects", value: "shadow:1" },
        { text: opt("te-shadow-offset-tl"), event: "text-effects", value: "shadow:2" },
        { text: opt("te-shadow-offset-b"), event: "text-effects", value: "shadow:3" },
        { text: opt("te-shadow-offset-r"), event: "text-effects", value: "shadow:4" },
        { text: opt("te-more-options"), event: "text-effects", value: "shadow:options" },
      ],
    },
    {
      text: opt("te-reflection"),
      children: [
        { text: opt("te-no-reflection"), event: "text-effects", value: "reflection:0" },
        { text: opt("te-reflection-tight"), event: "text-effects", value: "reflection:1" },
        { text: opt("te-reflection-half"), event: "text-effects", value: "reflection:2" },
        { text: opt("te-reflection-full"), event: "text-effects", value: "reflection:3" },
        { text: opt("te-more-options"), event: "text-effects", value: "reflection:options" },
      ],
    },
    {
      text: opt("te-glow"),
      children: [
        { text: opt("te-no-glow"), event: "text-effects", value: "glow:0" },
        { text: opt("te-glow-blue"), event: "text-effects", value: "glow:1" },
        { text: opt("te-glow-red"), event: "text-effects", value: "glow:2" },
        { text: opt("te-glow-green"), event: "text-effects", value: "glow:3" },
        { text: opt("te-glow-gold"), event: "text-effects", value: "glow:4" },
        { text: opt("te-more-options"), event: "text-effects", value: "glow:options" },
      ],
    },
    {
      text: opt("te-bevel"),
      children: [
        { text: opt("te-no-bevel"), event: "text-effects", value: "bevel:0" },
        { text: opt("te-bevel-circle"), event: "text-effects", value: "bevel:1" },
        { text: opt("te-bevel-relaxed-inset"), event: "text-effects", value: "bevel:2" },
        { text: opt("te-bevel-cross"), event: "text-effects", value: "bevel:3" },
        { text: opt("te-bevel-cool-slant"), event: "text-effects", value: "bevel:4" },
        { text: opt("te-bevel-angle"), event: "text-effects", value: "bevel:5" },
        { text: opt("te-bevel-soft-round"), event: "text-effects", value: "bevel:6" },
        { text: opt("te-more-options"), event: "text-effects", value: "bevel:options" },
      ],
    },
    {
      text: opt("te-rotation"),
      children: [
        { text: opt("te-no-rotation"), event: "text-effects", value: "rotation:0" },
        { text: opt("te-rotation-right-90"), event: "text-effects", value: "rotation:1" },
        { text: opt("te-rotation-left-90"), event: "text-effects", value: "rotation:2" },
        { text: opt("te-rotation-180"), event: "text-effects", value: "rotation:3" },
        { text: opt("te-rotation-oblique-tl"), event: "text-effects", value: "rotation:4" },
        { text: opt("te-rotation-isometric-top"), event: "text-effects", value: "rotation:5" },
        { text: opt("te-rotation-oblique-tr"), event: "text-effects", value: "rotation:6" },
        { text: opt("te-more-options"), event: "text-effects", value: "rotation:options" },
      ],
    },
    { text: opt("te-options"), event: "text-effects", value: "options" },
  ]);
