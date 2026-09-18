import type { ParagraphChild } from "@office-open/docx";

import type { JSONContent } from "../core";
import { Node } from "../core";
import type { ParseInlineRule, ResolveContext } from "./types";

/**
 * Lane W6.6: Ink & 3D Model Support.
 *
 * Explicit delegation notice:
 * "3D mesh vertex manipulation and ink vector splitting are delegated to external 3D/ink tools; viewing, layout, sizing, and alt text are editable natively."
 */
export const DELEGATION_NOTICE =
  "3D mesh vertex manipulation and ink vector splitting are delegated to external 3D/ink tools; viewing, layout, sizing, and alt text are editable natively.";

export const MODEL3D_URI = "http://schemas.microsoft.com/office/drawing/2017/model3d";
export const INK_2010_URI = "http://schemas.microsoft.com/office/drawing/2010/ink";
export const INK_MAIN_URI = "http://schemas.microsoft.com/ink/2010/main";
export const INKML_MIME = "application/inkml+xml";

export interface CameraRotation {
  lat?: number;
  lon?: number;
  rev?: number;
}

export interface CameraOptions {
  fov?: number;
  prst?: string;
  rot?: CameraRotation;
}

export interface Model3DAttrs {
  model3d?: Record<string, unknown> | null;
  rawXml?: string | null;
  cx?: number;
  cy?: number;
  width?: number;
  height?: number;
  title?: string;
  descr?: string;
  name?: string;
  id?: number | string;
  rId?: string;
  relType?: string;
  sourceFile?: string;
  camera?: CameraOptions | null;
  rotation?: number;
  floating?: Record<string, unknown> | null;
}

export interface InkAttrs {
  ink?: Record<string, unknown> | null;
  rawXml?: string | null;
  cx?: number;
  cy?: number;
  width?: number;
  height?: number;
  title?: string;
  descr?: string;
  name?: string;
  id?: number | string;
  rId?: string;
  traces?: string[];
  xmlData?: string;
  rotation?: number;
  floating?: Record<string, unknown> | null;
}

// ── XML string utilities ──

export function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function unescapeXml(s: string): string {
  return s
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}

function extractAttr(tag: string, attrName: string): string | undefined {
  const match = new RegExp(`${attrName}="([^"]*)"`, "i").exec(tag);
  return match ? unescapeXml(match[1]) : undefined;
}

// ── Detection Helpers ──

export function is3DModelXml(xml: string): boolean {
  if (!xml || typeof xml !== "string") return false;
  return (
    xml.includes(MODEL3D_URI) ||
    xml.includes("am3d:model3D") ||
    xml.includes("model3D") ||
    xml.includes(".glb") ||
    xml.includes(".3mf") ||
    xml.includes("model/gltf-binary") ||
    xml.includes("model/3mf")
  );
}

export function isInkXml(xml: string): boolean {
  if (!xml || typeof xml !== "string") return false;
  return (
    xml.includes(INK_2010_URI) ||
    xml.includes(INK_MAIN_URI) ||
    xml.includes("msink:ink") ||
    xml.includes("<ink:") ||
    xml.includes("</ink:") ||
    xml.includes(INKML_MIME) ||
    xml.includes("inkml")
  );
}

export function is3DModelChild(child: unknown, ctx?: ResolveContext): boolean {
  if (!child || typeof child !== "object") return false;
  const rec = child as Record<string, unknown>;
  if ("model3d" in rec && rec.model3d != null) return true;
  if ("rawXml" in rec && typeof rec.rawXml === "string" && is3DModelXml(rec.rawXml)) return true;
  if ("contentPart" in rec && rec.contentPart && typeof rec.contentPart === "object") {
    const cp = rec.contentPart as Record<string, unknown>;
    const rId = typeof cp.referenceId === "string" ? cp.referenceId : "";
    if (ctx && rId) {
      const relPath = (ctx as any).resolveRelationship?.(rId) ?? "";
      if (relPath.endsWith(".glb") || relPath.endsWith(".3mf")) return true;
    }
    const nvp = cp.nonVisualProperties as Record<string, unknown> | undefined;
    const name = typeof nvp?.name === "string" ? nvp.name : "";
    if (name.toLowerCase().includes("3d") || name.endsWith(".glb") || name.endsWith(".3mf"))
      return true;
  }
  return false;
}

export function isInkChild(child: unknown, ctx?: ResolveContext): boolean {
  if (!child || typeof child !== "object") return false;
  const rec = child as Record<string, unknown>;
  if ("ink" in rec && rec.ink != null) return true;
  if ("rawXml" in rec && typeof rec.rawXml === "string" && isInkXml(rec.rawXml)) return true;
  if ("contentPart" in rec && rec.contentPart && typeof rec.contentPart === "object") {
    const cp = rec.contentPart as Record<string, unknown>;
    const rId = typeof cp.referenceId === "string" ? cp.referenceId : "";
    if (ctx && rId) {
      const relPath = (ctx as any).resolveRelationship?.(rId) ?? "";
      if (relPath.endsWith(".inkml") || relPath.includes("ink")) return true;
    }
    const nvp = cp.nonVisualProperties as Record<string, unknown> | undefined;
    const name = typeof nvp?.name === "string" ? nvp.name : "";
    if (name.toLowerCase().includes("ink")) return true;
  }
  return false;
}

// ── XML Parsing & Attribute Extraction ──

export function parse3DModelFromXml(xml: string): Model3DAttrs {
  const attrs: Model3DAttrs = {
    rawXml: xml,
    cx: 1905000,
    cy: 1905000,
    width: 200,
    height: 200,
    title: "",
    descr: "",
    name: "3D Model",
    id: 1,
    rotation: 0,
    camera: null,
    floating: null,
  };

  // Extract wp:docPr attributes
  const docPrMatch = /<wp:docPr\b([^>]*)\/?>/i.exec(xml);
  if (docPrMatch) {
    const docPr = docPrMatch[1];
    const title = extractAttr(docPr, "title");
    if (title != null) attrs.title = title;
    const descr = extractAttr(docPr, "descr");
    if (descr != null) attrs.descr = descr;
    const name = extractAttr(docPr, "name");
    if (name != null) attrs.name = name;
    const id = extractAttr(docPr, "id");
    if (id != null) attrs.id = id;
  }

  // Extract wp:extent
  const extentMatch = /<wp:extent\b([^>]*)\/?>/i.exec(xml);
  if (extentMatch) {
    const ext = extentMatch[1];
    const cx = extractAttr(ext, "cx");
    const cy = extractAttr(ext, "cy");
    if (cx && !Number.isNaN(Number(cx))) {
      attrs.cx = Number(cx);
      attrs.width = Math.round(Number(cx) / 9525);
    }
    if (cy && !Number.isNaN(Number(cy))) {
      attrs.cy = Number(cy);
      attrs.height = Math.round(Number(cy) / 9525);
    }
  }

  // Extract relationship ID
  const rIdMatch = /\br:id="([^"]*)"/i.exec(xml) ?? /\br:embed="([^"]*)"/i.exec(xml);
  if (rIdMatch) {
    attrs.rId = rIdMatch[1];
  }

  // Extract camera & rotation
  const cameraMatch = /<am3d:camera\b([^>]*)>([\s\S]*?)<\/am3d:camera>/i.exec(xml);
  if (cameraMatch) {
    const camAttrs = cameraMatch[1];
    const camBody = cameraMatch[2];
    const fov = extractAttr(camAttrs, "fov");
    const prst = extractAttr(camAttrs, "prst");
    const camera: CameraOptions = {};
    if (fov && !Number.isNaN(Number(fov))) camera.fov = Number(fov);
    if (prst) camera.prst = prst;

    const rotMatch = /<am3d:rot\b([^>]*)\/?>/i.exec(camBody);
    if (rotMatch) {
      const rotAttrs = rotMatch[1];
      const lat = extractAttr(rotAttrs, "lat");
      const lon = extractAttr(rotAttrs, "lon");
      const rev = extractAttr(rotAttrs, "rev");
      const rot: CameraRotation = {};
      if (lat && !Number.isNaN(Number(lat))) rot.lat = Number(lat);
      if (lon && !Number.isNaN(Number(lon))) rot.lon = Number(lon);
      if (rev && !Number.isNaN(Number(rev))) {
        rot.rev = Number(rev);
        attrs.rotation = Math.round(Number(rev) / 60000);
      }
      camera.rot = rot;
    }
    attrs.camera = camera;
  }

  // Check if floating (wp:anchor)
  if (xml.includes("<wp:anchor")) {
    attrs.floating = { isFloating: true };
  }

  return attrs;
}

export function parseInkFromXml(xml: string): InkAttrs {
  const attrs: InkAttrs = {
    rawXml: xml,
    cx: 1524000,
    cy: 762000,
    width: 160,
    height: 80,
    title: "",
    descr: "",
    name: "Ink",
    id: 1,
    rotation: 0,
    traces: [],
    floating: null,
  };

  // Extract wp:docPr attributes
  const docPrMatch = /<wp:docPr\b([^>]*)\/?>/i.exec(xml);
  if (docPrMatch) {
    const docPr = docPrMatch[1];
    const title = extractAttr(docPr, "title");
    if (title != null) attrs.title = title;
    const descr = extractAttr(docPr, "descr");
    if (descr != null) attrs.descr = descr;
    const name = extractAttr(docPr, "name");
    if (name != null) attrs.name = name;
    const id = extractAttr(docPr, "id");
    if (id != null) attrs.id = id;
  }

  // Extract wp:extent
  const extentMatch = /<wp:extent\b([^>]*)\/?>/i.exec(xml);
  if (extentMatch) {
    const ext = extentMatch[1];
    const cx = extractAttr(ext, "cx");
    const cy = extractAttr(ext, "cy");
    if (cx && !Number.isNaN(Number(cx))) {
      attrs.cx = Number(cx);
      attrs.width = Math.round(Number(cx) / 9525);
    }
    if (cy && !Number.isNaN(Number(cy))) {
      attrs.cy = Number(cy);
      attrs.height = Math.round(Number(cy) / 9525);
    }
  }

  // Extract relationship ID
  const rIdMatch = /\br:id="([^"]*)"/i.exec(xml) ?? /\br:embed="([^"]*)"/i.exec(xml);
  if (rIdMatch) {
    attrs.rId = rIdMatch[1];
  }

  // Extract traces
  const traceMatches = xml.matchAll(/<msink:trace\b[^>]*>([\s\S]*?)<\/msink:trace>/gi);
  for (const m of traceMatches) {
    if (m[1]) attrs.traces?.push(m[1].trim());
  }

  if (xml.includes("<wp:anchor")) {
    attrs.floating = { isFloating: true };
  }

  return attrs;
}

// ── XML Update Helpers ──

export function updateDocPr(
  xml: string,
  optionsOrTitle?: string | { title?: string; descr?: string; name?: string },
  descr?: string,
  name?: string,
): string {
  let titleVal: string | undefined;
  let descrVal: string | undefined;
  let nameVal: string | undefined;

  if (typeof optionsOrTitle === "object" && optionsOrTitle !== null) {
    titleVal = optionsOrTitle.title;
    descrVal = optionsOrTitle.descr;
    nameVal = optionsOrTitle.name;
  } else {
    titleVal = optionsOrTitle;
    descrVal = descr;
    nameVal = name;
  }

  return xml.replace(/<wp:docPr\b([^>]*?)(\/?)>/i, (_match, body: string, selfClose: string) => {
    let updated = body;

    const setAttr = (k: string, v?: string) => {
      if (v === undefined) return;
      const re = new RegExp(`\\b${k}="[^"]*"`, "i");
      if (re.test(updated)) {
        updated = updated.replace(re, `${k}="${escapeXml(String(v))}"`);
      } else if (v) {
        updated += ` ${k}="${escapeXml(String(v))}"`;
      }
    };

    setAttr("title", titleVal);
    setAttr("descr", descrVal);
    setAttr("name", nameVal);

    return `<wp:docPr${updated}${selfClose}>`;
  });
}

export function updateExtent(xml: string, cx?: number, cy?: number): string {
  if (cx == null && cy == null) return xml;
  let res = xml;
  if (cx != null) {
    res = res.replace(/(<wp:extent\b[^>]*\bcx=")\d+(")/i, `$1${cx}$2`);
    res = res.replace(/(<a:ext\b[^>]*\bcx=")\d+(")/gi, `$1${cx}$2`);
  }
  if (cy != null) {
    res = res.replace(/(<wp:extent\b[^>]*\bcy=")\d+(")/i, `$1${cy}$2`);
    res = res.replace(/(<a:ext\b[^>]*\bcy=")\d+(")/gi, `$1${cy}$2`);
  }
  return res;
}

export function updateCamera(
  xml: string,
  cameraOrRot?: (CameraOptions & CameraRotation) | null,
): string {
  if (!cameraOrRot) return xml;
  let res = xml;
  if (cameraOrRot.fov != null) {
    res = res.replace(/(<am3d:camera\b[^>]*\bfov=")\d+(")/i, `$1${cameraOrRot.fov}$2`);
  }
  const rot =
    cameraOrRot.rot ??
    ("lat" in cameraOrRot || "lon" in cameraOrRot || "rev" in cameraOrRot
      ? cameraOrRot
      : undefined);
  if (rot) {
    const { lat, lon, rev } = rot;
    res = res.replace(/<am3d:rot\b([^>]*?)\/?>/i, (_match, body: string) => {
      let b = body;
      if (lat != null) {
        b = /\blat="[^"]*"/i.test(b)
          ? b.replace(/\blat="[^"]*"/i, `lat="${lat}"`)
          : `${b} lat="${lat}"`;
      }
      if (lon != null) {
        b = /\blon="[^"]*"/i.test(b)
          ? b.replace(/\blon="[^"]*"/i, `lon="${lon}"`)
          : `${b} lon="${lon}"`;
      }
      if (rev != null) {
        b = /\brev="[^"]*"/i.test(b)
          ? b.replace(/\brev="[^"]*"/i, `rev="${rev}"`)
          : `${b} rev="${rev}"`;
      }
      return `<am3d:rot${b}/>`;
    });
  }
  return res;
}

// ── XML Serialization Helpers ──

export function serialize3DModelXml(attrs: Model3DAttrs): string {
  const cx = attrs.cx ?? (attrs.width ? attrs.width * 9525 : 1905000);
  const cy = attrs.cy ?? (attrs.height ? attrs.height * 9525 : 1905000);
  const id = attrs.id ?? 1;
  const name = attrs.name ?? "3D Model";
  const title = attrs.title ?? "";
  const descr = attrs.descr ?? "";
  const rId = attrs.rId ?? "rId1";
  const fov = attrs.camera?.fov ?? 5400000;
  const lat = attrs.camera?.rot?.lat ?? 0;
  const lon = attrs.camera?.rot?.lon ?? 0;
  const rev = attrs.camera?.rot?.rev ?? (attrs.rotation ? attrs.rotation * 60000 : 0);

  return (
    `<w:r>` +
    `<w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="${escapeXml(name)}"${title ? ` title="${escapeXml(title)}"` : ""}${descr ? ` descr="${escapeXml(descr)}"` : ""}/>` +
    `<wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="${MODEL3D_URI}">` +
    `<am3d:model3D xmlns:am3d="${MODEL3D_URI}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rId}">` +
    `<am3d:spPr>` +
    `<a:xfrm>` +
    `<a:off x="0" y="0"/>` +
    `<a:ext cx="${cx}" cy="${cy}"/>` +
    `</a:xfrm>` +
    `</am3d:spPr>` +
    `<am3d:camera fov="${fov}">` +
    `<am3d:rot lat="${lat}" lon="${lon}" rev="${rev}"/>` +
    `</am3d:camera>` +
    `</am3d:model3D>` +
    `</a:graphicData>` +
    `</a:graphic>` +
    `</wp:inline>` +
    `</w:drawing>` +
    `</w:r>`
  );
}

export function serializeInkXml(attrs: InkAttrs): string {
  const cx = attrs.cx ?? (attrs.width ? attrs.width * 9525 : 1524000);
  const cy = attrs.cy ?? (attrs.height ? attrs.height * 9525 : 762000);
  const id = attrs.id ?? 1;
  const name = attrs.name ?? "Ink";
  const title = attrs.title ?? "";
  const descr = attrs.descr ?? "";
  const traces =
    attrs.traces && attrs.traces.length > 0 ? attrs.traces : ["100 200, 150 250, 200 220, 250 280"];
  const tracesXml = traces
    .map((t) => `<msink:trace type="pen">${escapeXml(t)}</msink:trace>`)
    .join("");

  return (
    `<w:r>` +
    `<w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="${escapeXml(name)}"${title ? ` title="${escapeXml(title)}"` : ""}${descr ? ` descr="${escapeXml(descr)}"` : ""}/>` +
    `<wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="${INK_2010_URI}">` +
    `<msink:ink xmlns:msink="${INK_2010_URI}">` +
    `<msink:context xml:id="ctx1"/>` +
    `${tracesXml}` +
    `</msink:ink>` +
    `</a:graphicData>` +
    `</a:graphic>` +
    `</wp:inline>` +
    `</w:drawing>` +
    `</w:r>`
  );
}

// ── Parse & Render Functions ──

export function parse3DModelDocx(
  child: ParagraphChild | string,
  ctx?: ResolveContext,
): JSONContent {
  let attrs: Model3DAttrs;

  if (typeof child === "string") {
    attrs = parse3DModelFromXml(child);
  } else {
    const rec = (child ?? {}) as Record<string, unknown>;
    if (typeof rec.rawXml === "string") {
      attrs = parse3DModelFromXml(rec.rawXml);
    } else if ("contentPart" in rec && rec.contentPart && typeof rec.contentPart === "object") {
      const cp = rec.contentPart as Record<string, unknown>;
      const nvp = (cp.nonVisualProperties ?? {}) as Record<string, unknown>;
      const tr = (cp.transformation ?? {}) as Record<string, unknown>;
      const w = typeof tr.width === "number" ? tr.width : 1905000;
      const h = typeof tr.height === "number" ? tr.height : 1905000;
      attrs = {
        cx: w,
        cy: h,
        width: Math.round(w / 9525),
        height: Math.round(h / 9525),
        id: (nvp.id as number | string) ?? 1,
        name: (nvp.name as string) ?? "3D Model",
        title: (nvp.title as string) ?? "",
        descr: (nvp.description as string) ?? "",
        rId: typeof cp.referenceId === "string" ? cp.referenceId : "",
        relType: "contentPart",
        rotation: typeof tr.rotation === "number" ? tr.rotation : 0,
      };
    } else if ("model3d" in rec && rec.model3d && typeof rec.model3d === "object") {
      attrs = { ...(rec.model3d as Model3DAttrs) };
    } else {
      attrs = {
        cx: 1905000,
        cy: 1905000,
        width: 200,
        height: 200,
        title: "",
        descr: "",
        name: "3D Model",
        id: 1,
      };
    }
  }

  if (attrs.rId && ctx) {
    const resolved = (ctx as any).resolveRelationship?.(attrs.rId);
    if (resolved) attrs.sourceFile = resolved;
  }

  // Ensure structured model3d payload mirrors attrs
  attrs.model3d = { ...attrs };

  return {
    type: "model3d",
    attrs,
  };
}

export function render3DModelDocx(node: JSONContent | Model3DAttrs): ParagraphChild {
  const attrs = ((node && typeof node === "object" && "attrs" in node && node.attrs
    ? node.attrs
    : node) ?? {}) as Model3DAttrs;
  let rawXml = attrs.rawXml;

  if (typeof rawXml === "string" && rawXml.length > 0) {
    // Losslessly preserve original rawXml while reflecting any edits made to dimensions and alt text
    rawXml = updateDocPr(rawXml, attrs.title, attrs.descr, attrs.name);
    rawXml = updateExtent(rawXml, attrs.cx, attrs.cy);
    if (attrs.camera) {
      rawXml = updateCamera(rawXml, attrs.camera);
    }
    // Ensure properly wrapped in <w:r> for paragraph inline emission
    if (rawXml.startsWith("<w:drawing")) {
      rawXml = `<w:r>${rawXml}</w:r>`;
    }
    return { rawXml } as unknown as ParagraphChild;
  }

  // Serialized from scratch
  return { rawXml: serialize3DModelXml(attrs) } as unknown as ParagraphChild;
}

export function parseInkDocx(child: ParagraphChild | string, ctx?: ResolveContext): JSONContent {
  let attrs: InkAttrs;

  if (typeof child === "string") {
    attrs = parseInkFromXml(child);
  } else {
    const rec = (child ?? {}) as Record<string, unknown>;
    if (typeof rec.rawXml === "string") {
      attrs = parseInkFromXml(rec.rawXml);
    } else if ("contentPart" in rec && rec.contentPart && typeof rec.contentPart === "object") {
      const cp = rec.contentPart as Record<string, unknown>;
      const nvp = (cp.nonVisualProperties ?? {}) as Record<string, unknown>;
      const tr = (cp.transformation ?? {}) as Record<string, unknown>;
      const w = typeof tr.width === "number" ? tr.width : 1524000;
      const h = typeof tr.height === "number" ? tr.height : 762000;
      attrs = {
        cx: w,
        cy: h,
        width: Math.round(w / 9525),
        height: Math.round(h / 9525),
        id: (nvp.id as number | string) ?? 1,
        name: (nvp.name as string) ?? "Ink",
        title: (nvp.title as string) ?? "",
        descr: (nvp.description as string) ?? "",
        rId: typeof cp.referenceId === "string" ? cp.referenceId : "",
        rotation: typeof tr.rotation === "number" ? tr.rotation : 0,
      };
    } else if ("ink" in rec && rec.ink && typeof rec.ink === "object") {
      attrs = { ...(rec.ink as InkAttrs) };
    } else {
      attrs = {
        cx: 1524000,
        cy: 762000,
        width: 160,
        height: 80,
        title: "",
        descr: "",
        name: "Ink",
        id: 1,
      };
    }
  }

  if (attrs.rId && ctx) {
    const resolved = (ctx as any).resolveRelationship?.(attrs.rId);
    if (resolved) attrs.title ||= resolved;
  }

  // Ensure structured ink payload mirrors attrs
  attrs.ink = { ...attrs };

  return {
    type: "ink",
    attrs,
  };
}

export function renderInkDocx(node: JSONContent | InkAttrs): ParagraphChild {
  const attrs = ((node && typeof node === "object" && "attrs" in node && node.attrs
    ? node.attrs
    : node) ?? {}) as InkAttrs;
  let rawXml = attrs.rawXml;

  if (typeof rawXml === "string" && rawXml.length > 0) {
    // Losslessly preserve original rawXml while reflecting any edits made to dimensions and alt text
    rawXml = updateDocPr(rawXml, attrs.title, attrs.descr, attrs.name);
    rawXml = updateExtent(rawXml, attrs.cx, attrs.cy);
    if (rawXml.startsWith("<w:drawing")) {
      rawXml = `<w:r>${rawXml}</w:r>`;
    }
    return { rawXml } as unknown as ParagraphChild;
  }

  // Serialized from scratch
  return { rawXml: serializeInkXml(attrs) } as unknown as ParagraphChild;
}

// ── Node Extensions ──

export const Model3D = Node.create({
  name: "model3d",
  inline: true,
  group: "inline",
  atom: true,
  draggable: true,

  addAttributes() {
    return {
      model3d: { default: null },
      rawXml: { default: null },
      cx: { default: 1905000 },
      cy: { default: 1905000 },
      width: { default: 200 },
      height: { default: 200 },
      title: { default: "" },
      descr: { default: "" },
      name: { default: "3D Model" },
      id: { default: 1 },
      rId: { default: null },
      sourceFile: { default: null },
      camera: { default: null },
      rotation: { default: 0 },
      floating: { default: null },
    };
  },

  parseHTML() {
    return [{ tag: "div[data-model3d]" }, { tag: "model3d" }];
  },

  renderDocx(node: JSONContent): Record<string, unknown> | null {
    return render3DModelDocx(node) as Record<string, unknown>;
  },

  parseDocxInline: {
    match: is3DModelChild,
    convert: parse3DModelDocx,
  } as ParseInlineRule,
});

export const Ink = Node.create({
  name: "ink",
  inline: true,
  group: "inline",
  atom: true,
  draggable: true,

  addAttributes() {
    return {
      ink: { default: null },
      rawXml: { default: null },
      cx: { default: 1524000 },
      cy: { default: 762000 },
      width: { default: 160 },
      height: { default: 80 },
      title: { default: "" },
      descr: { default: "" },
      name: { default: "Ink" },
      id: { default: 1 },
      rId: { default: null },
      traces: { default: [] },
      rotation: { default: 0 },
      floating: { default: null },
    };
  },

  parseHTML() {
    return [{ tag: "div[data-ink]" }, { tag: "ink" }];
  },

  renderDocx(node: JSONContent): Record<string, unknown> | null {
    return renderInkDocx(node) as Record<string, unknown>;
  },

  parseDocxInline: {
    match: isInkChild,
    convert: parseInkDocx,
  } as ParseInlineRule,
});

// Aliases
export { Model3D as Drawing3D, Ink as DrawingInk };
