import type { LayoutDrawingMember } from "@docen/layout";
import { Group, Path as LeaferPath, Rect, Text, type IGroup } from "leafer-ui";

/**
 * Lane W6.6: Ink & 3D Model Scene Painter.
 * Renders sleek isometric 3D wireframe / cube with rotation badge, model label,
 * dimensions, and alt text title; and cursive ink stroke path preview with pen-nib glyph.
 */

export function paint3DModelMember(
  tree: IGroup,
  m: Extract<LayoutDrawingMember, { kind: "model3d" }>,
): void {
  const g = new Group();
  tree.add(g);

  // Background container
  g.add(
    new Rect({
      x: m.x,
      y: m.y,
      width: m.width,
      height: m.height,
      fill: "#f8fafc",
      stroke: "#cbd5e1",
      strokeWidth: 1,
      cornerRadius: 6,
    }),
  );

  // Model label badge (top-left)
  const lblWidth = 66;
  const lblHeight = 20;
  g.add(
    new Rect({
      x: m.x + 8,
      y: m.y + 8,
      width: lblWidth,
      height: lblHeight,
      cornerRadius: 4,
      fill: "#dbeafe",
    }),
  );
  g.add(
    new Text({
      x: m.x + 8,
      y: m.y + 11,
      width: lblWidth,
      text: "3D MODEL",
      fill: "#1e40af",
      fontSize: 9,
      fontWeight: "bold",
      textAlign: "center",
    }),
  );

  // Rotation badge (top-right)
  const badgeWidth = 48;
  const badgeHeight = 20;
  const badgeX = m.x + m.width - badgeWidth - 8;
  const badgeY = m.y + 8;
  g.add(
    new Rect({
      x: badgeX,
      y: badgeY,
      width: badgeWidth,
      height: badgeHeight,
      cornerRadius: 10,
      fill: "#eff6ff",
      stroke: "#3b82f6",
      strokeWidth: 1,
    }),
  );
  g.add(
    new Text({
      x: badgeX,
      y: badgeY + 3,
      width: badgeWidth,
      text: "⟳ 360°",
      fill: "#1d4ed8",
      fontSize: 10,
      fontWeight: "bold",
      textAlign: "center",
    }),
  );

  // Sleek isometric 3D wireframe / cube
  const titleText = m.title || m.altText;
  const cx = m.x + m.width / 2;
  const cy = m.y + m.height / 2 - (titleText ? 8 : 0);
  const r = Math.max(14, Math.min(m.width, m.height) * 0.22);
  const dx = r * 0.866; // cos(30°)
  const dy = r * 0.5; // sin(30°)

  // Isometric Top face
  g.add(
    new LeaferPath({
      path: `M ${cx} ${cy - r} L ${cx + dx} ${cy - dy} L ${cx} ${cy} L ${cx - dx} ${cy - dy} Z`,
      fill: "#93c5fd",
      stroke: "#1e3a8a",
      strokeWidth: 1.5,
      strokeJoin: "round",
    }),
  );

  // Isometric Left face
  g.add(
    new LeaferPath({
      path: `M ${cx - dx} ${cy - dy} L ${cx} ${cy} L ${cx} ${cy + r} L ${cx - dx} ${cy + dy} Z`,
      fill: "#60a5fa",
      stroke: "#1e3a8a",
      strokeWidth: 1.5,
      strokeJoin: "round",
    }),
  );

  // Isometric Right face
  g.add(
    new LeaferPath({
      path: `M ${cx} ${cy} L ${cx + dx} ${cy - dy} L ${cx + dx} ${cy + dy} L ${cx} ${cy + r} Z`,
      fill: "#3b82f6",
      stroke: "#1e3a8a",
      strokeWidth: 1.5,
      strokeJoin: "round",
    }),
  );

  // Inner wireframe sub-division gridlines
  g.add(
    new LeaferPath({
      path: `M ${cx} ${cy - r} L ${cx} ${cy + r} M ${cx - dx} ${cy - dy} L ${cx + dx} ${cy + dy} M ${cx + dx} ${cy - dy} L ${cx - dx} ${cy + dy}`,
      stroke: "#ffffff60",
      strokeWidth: 1,
    }),
  );

  // Alt Text Title (centered below cube)
  if (titleText) {
    g.add(
      new Text({
        x: m.x + 8,
        y: cy + r + 6,
        width: m.width - 16,
        text: String(titleText),
        fill: "#0f172a",
        fontSize: 11,
        fontWeight: "bold",
        textAlign: "center",
      }),
    );
  }

  // Dimensions (bottom-left)
  g.add(
    new Text({
      x: m.x + 8,
      y: m.y + m.height - 16,
      text: `${Math.round(m.width)} × ${Math.round(m.height)} px`,
      fill: "#64748b",
      fontSize: 9,
    }),
  );
}

export function paintInkMember(
  tree: IGroup,
  m: Extract<LayoutDrawingMember, { kind: "ink" }>,
): void {
  const g = new Group();
  tree.add(g);

  // Background container
  g.add(
    new Rect({
      x: m.x,
      y: m.y,
      width: m.width,
      height: m.height,
      fill: "#fafaf9",
      stroke: "#e7e5e4",
      strokeWidth: 1,
      cornerRadius: 6,
    }),
  );

  // Ink label badge (top-left)
  const lblWidth = 44;
  const lblHeight = 18;
  g.add(
    new Rect({
      x: m.x + 8,
      y: m.y + 8,
      width: lblWidth,
      height: lblHeight,
      cornerRadius: 4,
      fill: "#f1f5f9",
    }),
  );
  g.add(
    new Text({
      x: m.x + 8,
      y: m.y + 11,
      width: lblWidth,
      text: "INK",
      fill: "#475569",
      fontSize: 9,
      fontWeight: "bold",
      textAlign: "center",
    }),
  );

  // Pen-nib glyph icon (next to badge)
  const nibX = m.x + 58;
  const nibY = m.y + 24;
  g.add(
    new LeaferPath({
      path: `M ${nibX} ${nibY} L ${nibX + 4} ${nibY - 9} L ${nibX + 5} ${nibY - 9} L ${nibX + 5} ${nibY - 13} L ${nibX - 5} ${nibY - 13} L ${nibX - 5} ${nibY - 9} L ${nibX - 4} ${nibY - 9} Z M ${nibX} ${nibY} L ${nibX} ${nibY - 6}`,
      fill: "#e0f2fe",
      stroke: "#0284c7",
      strokeWidth: 1.2,
      strokeJoin: "round",
    }),
  );

  // Stylized cursive ink stroke path preview
  const strokeLeft = m.x + 16;
  const strokeTop = m.y + 30;
  const w = Math.max(40, m.width - 32);
  const h = Math.max(20, m.height - 52);

  // A graceful calligraphic cursive flourish curve scaled to box
  const p0x = strokeLeft;
  const p0y = strokeTop + h * 0.5;
  const p1x = strokeLeft + w * 0.2;
  const p1y = strokeTop;
  const p2x = strokeLeft + w * 0.35;
  const p2y = strokeTop + h;
  const p3x = strokeLeft + w * 0.55;
  const p3y = strokeTop + h * 0.3;
  const p4x = strokeLeft + w * 0.75;
  const p4y = strokeTop + h * 0.8;
  const p5x = strokeLeft + w;
  const p5y = strokeTop + h * 0.2;

  const cursiveD = `M ${p0x} ${p0y} C ${p0x + w * 0.08} ${strokeTop - h * 0.1}, ${p1x} ${p1y}, ${strokeLeft + w * 0.25} ${strokeTop + h * 0.4} C ${strokeLeft + w * 0.3} ${strokeTop + h * 0.7}, ${p2x} ${p2y}, ${p3x} ${p3y} C ${strokeLeft + w * 0.65} ${strokeTop}, ${p4x} ${p4y}, ${p5x} ${p5y}`;

  g.add(
    new LeaferPath({
      path: cursiveD,
      stroke: "#1e3a8a",
      strokeWidth: 2.5,
      strokeCap: "round",
      strokeJoin: "round",
    }),
  );

  // Alt Text Title or caption
  const titleText = m.title || m.altText;
  if (titleText) {
    g.add(
      new Text({
        x: m.x + 8,
        y: m.y + m.height - 16,
        width: m.width - 16,
        text: String(titleText),
        fill: "#334155",
        fontSize: 10,
        fontWeight: "bold",
        textAlign: "center",
      }),
    );
  } else {
    // Dimensions (bottom-left)
    g.add(
      new Text({
        x: m.x + 8,
        y: m.y + m.height - 16,
        text: `${Math.round(m.width)} × ${Math.round(m.height)} px`,
        fill: "#78716c",
        fontSize: 9,
      }),
    );
  }
}
