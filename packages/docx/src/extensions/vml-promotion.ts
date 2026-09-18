import type { ParagraphChild } from "@office-open/docx";

import type { JSONContent } from "../core";
import {
  emuToTwip,
  parseLengthToTwips,
  parseVmlShapeLayout,
  stringifyVmlShapeLayout,
  twipToEmu,
  twipToPt,
} from "./drawing-shape-layout";

type Rec = Record<string, unknown>;

const isRecord = (v: unknown): v is Rec => typeof v === "object" && v !== null && !Array.isArray(v);

const NAMED_COLORS: Record<string, string> = {
  black: "000000",
  white: "FFFFFF",
  red: "FF0000",
  green: "008000",
  blue: "0000FF",
  yellow: "FFFF00",
  orange: "FFA500",
  purple: "800080",
  gray: "808080",
  grey: "808080",
  silver: "C0C0C0",
  maroon: "800000",
  olive: "808000",
  lime: "00FF00",
  aqua: "00FFFF",
  teal: "008080",
  navy: "000080",
  fuchsia: "FF00FF",
};

export function parseColorToHex(val: unknown): string | undefined {
  if (typeof val !== "string") return undefined;
  const s = val.trim().toLowerCase();
  if (!s) return undefined;
  if (NAMED_COLORS[s]) return NAMED_COLORS[s];

  if (s.startsWith("#")) {
    const raw = s.slice(1);
    if (raw.length === 3) {
      return (raw[0]! + raw[0]! + raw[1]! + raw[1]! + raw[2]! + raw[2]!).toUpperCase();
    }
    if (raw.length === 6) return raw.toUpperCase();
  } else if (/^[0-9a-f]{6}$/i.test(s)) {
    return s.toUpperCase();
  }

  const rgbMatch = /^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i.exec(s);
  if (rgbMatch) {
    const r = Math.min(255, Math.max(0, parseInt(rgbMatch[1]!, 10)));
    const g = Math.min(255, Math.max(0, parseInt(rgbMatch[2]!, 10)));
    const b = Math.min(255, Math.max(0, parseInt(rgbMatch[3]!, 10)));
    return ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  }

  return undefined;
}

export function parseWeightToEmu(val: unknown): number {
  if (typeof val === "number") return twipToEmu(val * 20);
  if (typeof val !== "string") return 12700; // 1pt default
  const twips = parseLengthToTwips(val);
  return twips != null ? twipToEmu(twips) : 12700;
}

export type VmlShapeKind = "rect" | "oval" | "line" | "polyline" | "roundrect" | "shape";

const VML_SHAPE_KINDS: readonly VmlShapeKind[] = [
  "rect",
  "oval",
  "line",
  "polyline",
  "roundrect",
  "shape",
];

export interface VmlOrigin {
  kind: VmlShapeKind;
  shapeData: Rec;
  pictWrapper: {
    media?: unknown[];
    vmlFallback?: string;
    mcChoiceRequires?: string;
    w14AnchorId?: string;
    otherChildren?: unknown[];
  };
}

/**
 * Checks if a ParagraphChild `{ pict: PictOptions }` is a non-textbox VML shape
 * that should be promoted to the modern `wpsShape` model.
 */
export function isPromotableVmlPict(pict: Rec): boolean {
  if (!isRecord(pict) || !Array.isArray(pict.children) || pict.children.length === 0) {
    return false;
  }

  // Find if there is a promotable shape child
  let foundShape = false;
  for (const child of pict.children) {
    if (!isRecord(child)) continue;
    for (const kind of VML_SHAPE_KINDS) {
      if (kind in child && isRecord(child[kind])) {
        const s = child[kind] as Rec;
        // Non-textbox: if it carries a textbox or imagedata, it's not a plain shape
        if (s.textbox != null || s.imagedata != null) return false;
        foundShape = true;
      }
    }
  }

  return foundShape;
}

/**
 * Promotes a legacy `<w:pict>` VML shape to a structured `wpsShape` node.
 */
export function promoteVmlPictToWpsShape(pict: Rec): JSONContent | null {
  if (!Array.isArray(pict.children)) return null;

  let targetKind: VmlShapeKind | undefined;
  let targetShapeData: Rec | undefined;
  let targetChildObj: Rec | undefined;

  for (const child of pict.children) {
    if (!isRecord(child)) continue;
    for (const kind of VML_SHAPE_KINDS) {
      if (kind in child && isRecord(child[kind])) {
        const s = child[kind] as Rec;
        if (s.textbox == null && s.imagedata == null) {
          targetKind = kind;
          targetShapeData = { ...s };
          targetChildObj = child;
          break;
        }
      }
    }
    if (targetKind) break;
  }

  if (!targetKind || !targetShapeData) return null;

  const otherChildren = pict.children.filter((c) => c !== targetChildObj);
  const layout = parseVmlShapeLayout(targetShapeData.style as string | Rec | undefined);

  // Width and height in EMU
  const widthTwips = layout?.width ?? 2000; // 100pt default
  const heightTwips = layout?.height ?? 2000;
  const widthEmu = twipToEmu(widthTwips);
  const heightEmu = twipToEmu(heightTwips);

  const transformation: Rec = {
    width: widthEmu,
    height: heightEmu,
    ...(layout?.rotation != null && layout.rotation !== 0 ? { rotation: layout.rotation } : {}),
  };

  // Floating properties
  let floating: Rec | undefined;
  if (layout?.position === "absolute" || layout?.position === "relative") {
    const leftTwips = layout.left ?? layout.marginLeft ?? 0;
    const topTwips = layout.top ?? layout.marginTop ?? 0;
    floating = {
      horizontalPosition: layout.horizontalPosition
        ? { relative: layout.horizontalRelative ?? "page", align: layout.horizontalPosition }
        : { relative: layout.horizontalRelative ?? "page", offset: twipToEmu(leftTwips) },
      verticalPosition: layout.verticalPosition
        ? { relative: layout.verticalRelative ?? "paragraph", align: layout.verticalPosition }
        : { relative: layout.verticalRelative ?? "paragraph", offset: twipToEmu(topTwips) },
      wrap: { type: layout.wrapping ?? "square" },
      ...(layout.zIndex != null ? { zIndex: layout.zIndex } : {}),
    };
  }

  // Geometry mapping
  let presetGeometry: string | undefined;
  let customGeometry: Rec | undefined;

  switch (targetKind) {
    case "rect":
      presetGeometry = "rect";
      break;
    case "oval":
      presetGeometry = "ellipse";
      break;
    case "roundrect":
      presetGeometry = "roundRect";
      break;
    case "line":
      presetGeometry = "line";
      break;
    case "polyline": {
      const pointsStr =
        typeof targetShapeData.points === "string" ? targetShapeData.points.trim() : "";
      if (pointsStr) {
        const ptPairs = pointsStr.split(/\s+/).map((pair) => {
          const [px, py] = pair.split(",").map(Number);
          return { x: px ?? 0, y: py ?? 0 };
        });
        if (ptPairs.length > 0) {
          const commands: Rec[] = [];
          commands.push({
            command: "moveTo",
            point: { x: String(ptPairs[0]!.x), y: String(ptPairs[0]!.y) },
          });
          for (let i = 1; i < ptPairs.length; i++) {
            commands.push({
              command: "lnTo",
              point: { x: String(ptPairs[i]!.x), y: String(ptPairs[i]!.y) },
            });
          }
          customGeometry = {
            pathList: [{ w: widthTwips, h: heightTwips, commands }],
          };
        }
      }
      if (!customGeometry) presetGeometry = "line";
      break;
    }
    case "shape": {
      const pathStr = typeof targetShapeData.path === "string" ? targetShapeData.path.trim() : "";
      if (pathStr) {
        // Parse simple VML path commands e.g. "m 0,0 l 100,100 e" or "m 0,0 l 100,0, 100,100, 0,100 x e"
        const tokens = pathStr.split(/\s+/);
        const commands: Rec[] = [];
        let i = 0;
        while (i < tokens.length) {
          const tok = tokens[i]!.toLowerCase();
          if (tok === "m" && i + 1 < tokens.length) {
            const [x, y] = tokens[++i]!.split(",").map(Number);
            commands.push({ command: "moveTo", point: { x: String(x ?? 0), y: String(y ?? 0) } });
          } else if (tok === "l" && i + 1 < tokens.length) {
            const parts = tokens[++i]!.split(",");
            for (let j = 0; j < parts.length; j += 2) {
              const x = Number(parts[j]);
              const y = Number(parts[j + 1]);
              commands.push({ command: "lnTo", point: { x: String(x ?? 0), y: String(y ?? 0) } });
            }
          } else if (tok === "x" || tok === "close") {
            commands.push({ command: "close" });
          }
          i++;
        }
        if (commands.length > 0) {
          customGeometry = {
            pathList: [{ w: widthTwips, h: heightTwips, commands }],
          };
        }
      }
      if (!customGeometry) presetGeometry = "rect";
      break;
    }
  }

  // Fill styling
  let fill: Rec | undefined;
  if (targetShapeData.filled === false || targetShapeData.filled === "f") {
    fill = { type: "none" };
  } else {
    const rawColor =
      targetShapeData.fillcolor ??
      (isRecord(targetShapeData.fill) ? targetShapeData.fill.color : undefined);
    const hex = parseColorToHex(rawColor);
    fill = hex ? { type: "solid", color: hex } : { type: "solid", color: "4472C4" };
  }

  // Outline styling
  let outline: Rec | undefined;
  if (targetShapeData.stroked === false || targetShapeData.stroked === "f") {
    outline = { type: "none" };
  } else {
    const rawColor =
      targetShapeData.strokecolor ??
      (isRecord(targetShapeData.stroke) ? targetShapeData.stroke.color : undefined);
    const hex = parseColorToHex(rawColor);
    const weight =
      targetShapeData.strokeweight ??
      (isRecord(targetShapeData.stroke) ? targetShapeData.stroke.weight : undefined);
    const width = parseWeightToEmu(weight);
    outline = { color: hex ?? "000000", width };
  }

  const vmlOrigin: VmlOrigin = {
    kind: targetKind,
    shapeData: targetShapeData,
    pictWrapper: {
      media: Array.isArray(pict.media) ? pict.media : undefined,
      vmlFallback: typeof pict.vmlFallback === "string" ? pict.vmlFallback : undefined,
      mcChoiceRequires:
        typeof pict.mcChoiceRequires === "string" ? pict.mcChoiceRequires : undefined,
      w14AnchorId: typeof pict.w14AnchorId === "string" ? pict.w14AnchorId : undefined,
      otherChildren: otherChildren.length > 0 ? otherChildren : undefined,
    },
  };

  const wpsAttrs: Rec = {
    ...(presetGeometry ? { presetGeometry, geometry: presetGeometry } : {}),
    ...(customGeometry ? { customGeometry } : {}),
    transformation,
    ...(floating ? { floating } : {}),
    fill,
    outline,
    vmlOrigin,
  };

  return {
    type: "wpsShape",
    attrs: { wpsShape: wpsAttrs },
    content: [{ type: "paragraph" }],
  };
}

/**
 * Reverts a promoted `wpsShape` node back into its lossless `<w:pict>` ParagraphChild.
 */
export function revertWpsShapeToVmlPict(geometry: Rec): ParagraphChild {
  const origin = geometry.vmlOrigin as VmlOrigin | undefined;
  if (!origin || !origin.kind || !origin.shapeData) {
    return { pict: {} } as unknown as ParagraphChild;
  }

  const { kind, shapeData, pictWrapper } = origin;
  const updatedShape = { ...shapeData };

  // Parse existing style or create fresh layout
  const layout = parseVmlShapeLayout(updatedShape.style as string | Rec | undefined) ?? {};

  // Update layout from transformation if modified
  const tr = isRecord(geometry.transformation) ? geometry.transformation : undefined;
  if (tr) {
    if (typeof tr.width === "number" && tr.width > 0) {
      layout.width = emuToTwip(tr.width);
    }
    if (typeof tr.height === "number" && tr.height > 0) {
      layout.height = emuToTwip(tr.height);
    }
    if (typeof tr.rotation === "number") {
      layout.rotation = tr.rotation;
    }
  }

  // Update floating if present
  const fl = isRecord(geometry.floating) ? geometry.floating : undefined;
  if (fl) {
    layout.position = "absolute";
    const hp = isRecord(fl.horizontalPosition) ? fl.horizontalPosition : undefined;
    if (hp && typeof hp.offset === "number") {
      layout.left = emuToTwip(hp.offset);
    }
    const vp = isRecord(fl.verticalPosition) ? fl.verticalPosition : undefined;
    if (vp && typeof vp.offset === "number") {
      layout.top = emuToTwip(vp.offset);
    }
    const wrap = isRecord(fl.wrap) ? fl.wrap : undefined;
    if (wrap && typeof wrap.type === "string") {
      layout.wrapping = wrap.type as any;
    }
    if (typeof fl.zIndex === "number") {
      layout.zIndex = fl.zIndex;
    }
  }

  updatedShape.style = stringifyVmlShapeLayout(layout);

  // Update fill if modified
  const fill = isRecord(geometry.fill) ? geometry.fill : undefined;
  if (fill) {
    if (fill.type === "none") {
      updatedShape.filled = "f";
    } else if (fill.type === "solid" && typeof fill.color === "string") {
      updatedShape.filled = "t";
      updatedShape.fillcolor = `#${fill.color.replace("#", "")}`;
    }
  }

  // Update outline if modified
  const outline = isRecord(geometry.outline) ? geometry.outline : undefined;
  if (outline) {
    if (outline.type === "none") {
      updatedShape.stroked = "f";
    } else {
      updatedShape.stroked = "t";
      if (typeof outline.color === "string") {
        updatedShape.strokecolor = `#${outline.color.replace("#", "")}`;
      }
      if (typeof outline.width === "number") {
        updatedShape.strokeweight = `${twipToPt(emuToTwip(outline.width))}pt`;
      }
    }
  }

  const childObj = { [kind]: updatedShape };
  const allChildren = [...(pictWrapper.otherChildren ?? []), childObj];

  const pict: Rec = {
    children: allChildren,
    ...(pictWrapper.media ? { media: pictWrapper.media } : {}),
    ...(pictWrapper.vmlFallback ? { vmlFallback: pictWrapper.vmlFallback } : {}),
    ...(pictWrapper.mcChoiceRequires ? { mcChoiceRequires: pictWrapper.mcChoiceRequires } : {}),
    ...(pictWrapper.w14AnchorId ? { w14AnchorId: pictWrapper.w14AnchorId } : {}),
  };

  return { pict } as unknown as ParagraphChild;
}
