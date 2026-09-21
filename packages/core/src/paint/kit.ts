/**
 * PaintKit — graphics abstraction decoupling document layout painting from
 * concrete rendering runtimes.
 *
 * Provides:
 * - `leaferKit`: browser canvas backend using live LeaferJS UI elements.
 * - `nodeKit`: headless scene graph recording vector geometry, transforms,
 *   clipping, and text rows for server-side PDF generation without DOM/Canvas.
 *
 * @module
 */

export interface AffineMatrix {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export const IDENTITY_MATRIX: AffineMatrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

export function composeMatrix(first: AffineMatrix, second: AffineMatrix): AffineMatrix {
  return {
    a: first.a * second.a + first.c * second.b,
    b: first.b * second.a + first.d * second.b,
    c: first.a * second.c + first.c * second.d,
    d: first.b * second.c + first.d * second.d,
    e: first.a * second.e + first.c * second.f + first.e,
    f: first.b * second.e + first.d * second.f + first.f,
  };
}

export function localMatrixOf(x = 0, y = 0, scaleX = 1, scaleY = 1, rotationDeg = 0): AffineMatrix {
  if (rotationDeg === 0) {
    return { a: scaleX, b: 0, c: 0, d: scaleY, e: x, f: y };
  }
  const rad = (rotationDeg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  return {
    a: scaleX * cos,
    b: scaleX * sin,
    c: -scaleY * sin,
    d: scaleY * cos,
    e: x,
    f: y,
  };
}

export interface PaintNodeLike {
  tag: string;
  parent?: PaintGroupLike;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  scaleX?: number;
  scaleY?: number;
  scale?: number;
  rotation?: number;
  opacity?: number;
  visible?: boolean;
  blendMode?: string;
  data?: Record<string, unknown>;
  fill?: unknown;
  stroke?: unknown;
  strokeWidth?: number;
  strokeCap?: string;
  strokeJoin?: string;
  dashPattern?: readonly number[];
  windingRule?: string;
  hittable?: boolean;
  readonly worldTransform: AffineMatrix;
  readonly worldOpacity: number;
  getPathString?(useAbsolute?: boolean, ignoreTransform?: boolean): string;
  descendants?(): PaintNodeLike[];
}

export interface PaintGroupLike extends PaintNodeLike {
  children: PaintNodeLike[];
  add(...children: (PaintNodeLike | readonly PaintNodeLike[])[]): void;
  clear(): void;
}

export interface PaintKit {
  createGroup(props?: Record<string, unknown>): any;
  createBox(props?: Record<string, unknown>): any;
  createRect(props?: Record<string, unknown>): any;
  createLine(props?: Record<string, unknown>): any;
  createText(props?: Record<string, unknown>): any;
  createPath(props?: Record<string, unknown>): any;
  createEllipse(props?: Record<string, unknown>): any;
  createImage(props?: Record<string, unknown>): any;
}

// ─────────────────────────────────────────────────────────────────────────────
// 1. nodeKit — pure Node headless scene runtime (zero DOM / Canvas)
// ─────────────────────────────────────────────────────────────────────────────

export abstract class NodeBaseElement implements PaintNodeLike {
  abstract readonly tag: string;
  parent?: PaintGroupLike;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  scaleX?: number;
  scaleY?: number;
  scale?: number;
  rotation?: number;
  opacity?: number;
  visible = true;
  blendMode?: string;
  data?: Record<string, unknown>;
  fill?: unknown;
  stroke?: unknown;
  strokeWidth?: number;
  strokeCap?: string;
  strokeJoin?: string;
  dashPattern?: readonly number[];
  windingRule?: string;
  hittable = false;

  constructor(props?: Record<string, unknown>) {
    if (props) Object.assign(this, props);
  }

  get type(): string {
    return this.tag;
  }

  get props(): this {
    return this;
  }

  descendants(): PaintNodeLike[] {
    const children = (this as unknown as { children?: PaintNodeLike[] }).children ?? [];
    return children.flatMap((c) => [
      c,
      ...(typeof (c as any).descendants === "function" ? (c as any).descendants() : []),
    ]);
  }

  get worldTransform(): AffineMatrix {
    const s = this.scale ?? 1;
    const sx = (this.scaleX ?? 1) * s;
    const sy = (this.scaleY ?? 1) * s;
    const local = localMatrixOf(this.x ?? 0, this.y ?? 0, sx, sy, this.rotation ?? 0);
    if (!this.parent) return local;
    return composeMatrix(this.parent.worldTransform, local);
  }

  get worldOpacity(): number {
    const parentOpacity = this.parent?.worldOpacity ?? 1;
    const localOpacity = this.opacity ?? 1;
    return parentOpacity * localOpacity;
  }
}

export class NodeGroup extends NodeBaseElement implements PaintGroupLike {
  override readonly tag: string = "Group";
  children: PaintNodeLike[] = [];

  add(...children: (PaintNodeLike | readonly PaintNodeLike[])[]): void {
    for (const child of children) {
      if (Array.isArray(child)) {
        for (const c of child) {
          if (c) {
            c.parent = this;
            this.children.push(c);
          }
        }
      } else if (child) {
        const c = child as PaintNodeLike;
        c.parent = this;
        this.children.push(c);
      }
    }
  }

  clear(): void {
    this.children = [];
  }
}

export class NodeBox extends NodeGroup {
  override readonly tag: string = "Box";
  declare overflow?: string;

  getPathString(): string {
    const w = this.width ?? 0;
    const h = this.height ?? 0;
    return `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`;
  }
}

export class NodeRect extends NodeBaseElement {
  readonly tag = "Rect";
  declare cornerRadius?: number;

  getPathString(): string {
    const w = this.width ?? 0;
    const h = this.height ?? 0;
    const cr = this.cornerRadius;
    if (cr && cr > 0) {
      const r = Math.min(cr, w / 2, h / 2);
      return (
        `M ${r} 0 ` +
        `L ${w - r} 0 Q ${w} 0 ${w} ${r} ` +
        `L ${w} ${h - r} Q ${w} ${h} ${w - r} ${h} ` +
        `L ${r} ${h} Q 0 ${h} 0 ${h - r} ` +
        `L 0 ${r} Q 0 0 ${r} 0 Z`
      );
    }
    return `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`;
  }
}

export class NodeLine extends NodeBaseElement {
  readonly tag = "Line";
  declare points?: readonly number[];

  getPathString(): string {
    const pts = this.points;
    if (!pts || pts.length < 4) return "";
    let d = `M ${pts[0]} ${pts[1]}`;
    for (let i = 2; i < pts.length; i += 2) {
      d += ` L ${pts[i]} ${pts[i + 1]}`;
    }
    return d;
  }
}

export class NodePath extends NodeBaseElement {
  readonly tag = "Path";
  declare path?: string;

  getPathString(): string {
    return this.path ?? "";
  }
}

export class NodeEllipse extends NodeBaseElement {
  readonly tag = "Ellipse";

  getPathString(): string {
    const w = this.width ?? 0;
    const h = this.height ?? 0;
    const rx = w / 2;
    const ry = h / 2;
    const cx = rx;
    const cy = ry;
    return (
      `M ${cx + rx} ${cy} ` +
      `A ${rx} ${ry} 0 1 0 ${cx - rx} ${cy} ` +
      `A ${rx} ${ry} 0 1 0 ${cx + rx} ${cy} Z`
    );
  }
}

export class NodeText extends NodeBaseElement {
  readonly tag = "Text";
  declare text?: string | number;
  declare fontSize?: number;
  declare fontFamily?: string;
  declare fontWeight?: number | string;
  declare italic?: boolean;
  declare letterSpacing?: unknown;
  declare lineHeight?: number;
  declare textAlign?: string;
  declare textWrap?: string;
  declare textDecoration?: string;

  get textDrawData(): { rows: Array<{ x: number; y: number; width: number; text: string }> } {
    const str = typeof this.text === "string" ? this.text : String(this.text ?? "");
    const size = this.fontSize ?? 14;
    return {
      rows: [
        {
          x: 0,
          y: size * 0.8,
          width: this.width ?? 0,
          text: str,
        },
      ],
    };
  }
}

export class NodeImage extends NodeBaseElement {
  readonly tag = "Image";
  declare url?: string;
  declare image?: {
    width?: number;
    height?: number;
    naturalWidth?: number;
    naturalHeight?: number;
  };
}

export const nodeKit: PaintKit = {
  createGroup: (props) => new NodeGroup(props),
  createBox: (props) => new NodeBox(props),
  createRect: (props) => new NodeRect(props),
  createLine: (props) => new NodeLine(props),
  createText: (props) => new NodeText(props),
  createPath: (props) => new NodePath(props),
  createEllipse: (props) => new NodeEllipse(props),
  createImage: (props) => new NodeImage(props),
};

// ─────────────────────────────────────────────────────────────────────────────
// 2. leaferKit — registered browser runtime (zero static Leafer import in Node)
// ─────────────────────────────────────────────────────────────────────────────

export interface LeaferClasses {
  Group: any;
  Box: any;
  Rect: any;
  Line: any;
  Text: any;
  Path: any;
  Ellipse: any;
  Image: any;
  ImageManager?: any;
  Resource?: any;
}

let registeredLeafer: LeaferClasses | null = null;

export function registerLeafer(classes: LeaferClasses): void {
  registeredLeafer = classes;
}

export function isLeaferRegistered(): boolean {
  return registeredLeafer !== null;
}

export function getRegisteredLeafer(): LeaferClasses | null {
  return registeredLeafer;
}

if (
  typeof window !== "undefined" ||
  (typeof globalThis !== "undefined" && (globalThis as any).__vitest_worker__)
) {
  import("leafer-ui")
    .then((mod) => {
      if (!registeredLeafer && mod) {
        const classes = mod.default && (mod.default as any).Line ? mod.default : mod;
        registerLeafer(classes as any);
      }
    })
    .catch(() => {});
}

export const leaferKit: PaintKit = {
  createGroup: (props) =>
    registeredLeafer ? new registeredLeafer.Group(props) : new NodeGroup(props),
  createBox: (props) => (registeredLeafer ? new registeredLeafer.Box(props) : new NodeBox(props)),
  createRect: (props) =>
    registeredLeafer ? new registeredLeafer.Rect(props) : new NodeRect(props),
  createLine: (props) =>
    registeredLeafer ? new registeredLeafer.Line(props) : new NodeLine(props),
  createText: (props) =>
    registeredLeafer ? new registeredLeafer.Text(props) : new NodeText(props),
  createPath: (props) =>
    registeredLeafer ? new registeredLeafer.Path(props) : new NodePath(props),
  createEllipse: (props) =>
    registeredLeafer ? new registeredLeafer.Ellipse(props) : new NodeEllipse(props),
  createImage: (props) =>
    registeredLeafer ? new registeredLeafer.Image(props) : new NodeImage(props),
};

// ─────────────────────────────────────────────────────────────────────────────
// 3. Dynamic Kit Context & Constructor Proxies
// ─────────────────────────────────────────────────────────────────────────────

let activeKit: PaintKit = leaferKit;

export function getActiveKit(): PaintKit {
  return activeKit;
}

export function setActiveKit(kit: PaintKit): void {
  activeKit = kit;
}

export function withKit<T>(kit: PaintKit, fn: () => T): T {
  const prev = activeKit;
  activeKit = kit;
  try {
    return fn();
  } finally {
    activeKit = prev;
  }
}

export const Group: any = class Group {
  constructor(props?: any) {
    return activeKit.createGroup(props);
  }
};
export type Group = any;

export const Box: any = class Box {
  constructor(props?: any) {
    return activeKit.createBox(props);
  }
};
export type Box = any;

export const Rect: any = class Rect {
  constructor(props?: any) {
    return activeKit.createRect(props);
  }
};
export type Rect = any;

export const Line: any = class Line {
  constructor(props?: any) {
    return activeKit.createLine(props);
  }
};
export type Line = any;

export const Text: any = class Text {
  constructor(props?: any) {
    return activeKit.createText(props);
  }
};
export type Text = any;

export const Path: any = class Path {
  constructor(props?: any) {
    return activeKit.createPath(props);
  }
};
export type Path = any;

export const Ellipse: any = class Ellipse {
  constructor(props?: any) {
    return activeKit.createEllipse(props);
  }
};
export type Ellipse = any;

export const Image: any = class Image {
  constructor(props?: any) {
    return activeKit.createImage(props);
  }
};
export type Image = any;

export type IGroup = any;
export type ILeaferImage = any;
