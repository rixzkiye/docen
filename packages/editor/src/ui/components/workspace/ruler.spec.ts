// @vitest-environment happy-dom
import { docxExtensions } from "@docen/docx";
import { Editor } from "@tiptap/core";
import { afterEach, describe, expect, it } from "vitest";

import { DocenRuler } from "./ruler";

const created: DocenRuler[] = [];
afterEach(() => {
  while (created.length) created.pop()!.remove();
});

async function settle(): Promise<void> {
  for (let i = 0; i < 3; i++) await new Promise((r) => setTimeout(r, 0));
}

async function mountRuler(): Promise<DocenRuler> {
  const ruler = new DocenRuler();
  created.push(ruler);
  document.body.append(ruler);
  await settle();
  return ruler;
}

describe("DocenRuler (<docen-ruler>) (W5.2)", () => {
  it("mounts and computes geometry, zero point, and content width", async () => {
    const ruler = await mountRuler();
    ruler.pageWidthPx = 816; // 8.5" * 96
    ruler.marginLeftPx = 96; // 1"
    ruler.marginRightPx = 96; // 1"
    ruler.scale = 1;

    expect(ruler.zeroXPx).toBe(96);
    expect(ruler.contentWidthPx).toBe(624); // 816 - 96 - 96
  });

  it("renders and toggles units between inches and centimeters", async () => {
    const ruler = await mountRuler();
    ruler.unit = "in";

    let emittedUnit = "";
    ruler.addEventListener("ruler:unit-change", (e: any) => {
      emittedUnit = e.detail.unit;
    });

    ruler.toggleUnit();
    expect(ruler.unit).toBe("cm");
    expect(emittedUnit).toBe("cm");

    ruler.toggleUnit();
    expect(ruler.unit).toBe("in");
    expect(emittedUnit).toBe("in");
  });

  it("computes marker X positions for first-line, hanging, left base, and right indent", async () => {
    const ruler = await mountRuler();
    ruler.marginLeftPx = 96;
    ruler.marginRightPx = 96;
    ruler.pageWidthPx = 816;

    // Default: 0 indents
    expect(ruler.hangingMarkerX).toBe(96);
    expect(ruler.firstLineMarkerX).toBe(96);
    expect(ruler.rightMarkerX).toBe(96 + 624); // 720

    // Set 0.5" left indent (720 twips = 48px), 0.25" first-line indent (360 twips = 24px)
    ruler.setParagraphAttrs({
      left: 720,
      firstLine: 360,
      right: 1440, // 1" right indent = 96px
    });

    expect(ruler.hangingMarkerX).toBe(96 + 48); // 144
    expect(ruler.firstLineMarkerX).toBe(96 + 48 + 24); // 168
    expect(ruler.rightMarkerX).toBe(720 - 96); // 624
  });

  it("handles draggable indent markers with pointer events", async () => {
    const ruler = await mountRuler();
    ruler.marginLeftPx = 96;
    ruler.pageWidthPx = 816;

    let emittedIndent: any = null;
    ruler.addEventListener("ruler:indent-change", (e: any) => {
      emittedIndent = e.detail.indent;
    });

    const firstLineEl = ruler.shadowRoot!.querySelector(".first-line-marker") as HTMLElement;
    expect(firstLineEl).toBeTruthy();
    firstLineEl.setPointerCapture = () => {};

    // Start drag on first line marker
    ruler.onMarkerPointerDown("firstLine", {
      preventDefault: () => {},
      stopPropagation: () => {},
      clientX: 96,
      pointerId: 1,
      target: firstLineEl,
    } as any);

    expect(ruler.dragActiveMarker).toBe("firstLine");

    // Move pointer 48px to the right (= 720 twips)
    window.dispatchEvent(new PointerEvent("pointermove", { clientX: 144 }));
    expect(ruler.firstLineTwips).toBe(720);
    expect(emittedIndent.firstLine).toBe(720);

    // Release pointer
    window.dispatchEvent(new PointerEvent("pointerup"));
    expect(ruler.dragActiveMarker).toBeNull();
  });

  it("handles tab stops: click to add, drag to move, drag off to delete, dblclick to open dialog", async () => {
    const ruler = await mountRuler();
    ruler.marginLeftPx = 96;
    ruler.pageWidthPx = 816;

    let tabAdded: any = null;
    let tabRemoved: any = null;
    let openTabsFired = false;

    ruler.addEventListener("ruler:tabstop-add", (e: any) => {
      tabAdded = e.detail.tabStop;
    });
    ruler.addEventListener("ruler:tabstop-remove", (e: any) => {
      tabRemoved = e.detail.tabStop;
    });
    ruler.addEventListener("ruler:open-tabs", () => {
      openTabsFired = true;
    });

    // Click at 96 + 96 = 192px (1 inch into content = 1440 twips)
    const track = ruler.shadowRoot!.querySelector(".ruler-track") as HTMLElement;
    track.getBoundingClientRect = () => ({
      left: 0,
      top: 10,
      width: 816,
      height: 24,
      right: 816,
      bottom: 34,
      x: 0,
      y: 10,
      toJSON: () => {},
    });

    ruler.onTrackClick({ clientX: 192, currentTarget: track } as any);
    expect(ruler.tabStops).toHaveLength(1);
    expect(ruler.tabStops[0]!.position).toBe(1440);
    expect(tabAdded.position).toBe(1440);

    // Drag tab stop to move
    ruler.onTabPointerDown(0, {
      preventDefault: () => {},
      stopPropagation: () => {},
      clientX: 192,
      pointerId: 2,
      target: { setPointerCapture: () => {} },
    } as any);

    expect(ruler.dragActiveMarker).toBe("tabStop");

    // Move to 288px (2 inches = 2880 twips)
    window.dispatchEvent(new PointerEvent("pointermove", { clientX: 288, clientY: 20 }));
    expect(ruler.tabStops[0]!.position).toBe(2880);

    // Drag off ruler vertically (clientY = 100 > 28px away)
    window.dispatchEvent(new PointerEvent("pointermove", { clientX: 288, clientY: 100 }));
    expect(ruler.dragOffRuler).toBe(true);

    // Release off ruler -> deletes tab stop!
    window.dispatchEvent(new PointerEvent("pointerup"));
    expect(ruler.tabStops).toHaveLength(0);
    expect(tabRemoved).toBeTruthy();

    // Re-add tab stop and double click
    ruler.tabStops = [{ position: 1440, type: "left" }];
    ruler.onTabDblClick(0, { stopPropagation: () => {} } as any);
    expect(openTabsFired).toBe(true);
  });

  it("shows Alt precision tooltip with exact measurement during drag", async () => {
    const ruler = await mountRuler();
    ruler.unit = "in";
    const track = ruler.shadowRoot!.querySelector(".ruler-track") as HTMLElement;
    track.getBoundingClientRect = () => ({
      left: 0,
      top: 0,
      width: 816,
      height: 24,
      right: 816,
      bottom: 24,
      x: 0,
      y: 0,
      toJSON: () => {},
    });

    ruler.onMarkerPointerDown("firstLine", {
      preventDefault: () => {},
      stopPropagation: () => {},
      clientX: 96,
      altKey: true,
      pointerId: 1,
      target: { setPointerCapture: () => {} },
    } as any);

    expect(ruler.showTooltip).toBe(true);

    // Move 144px (= 1.5 in = 2160 twips)
    window.dispatchEvent(new PointerEvent("pointermove", { clientX: 240, altKey: true }));
    expect(ruler.tooltipText).toContain("in");
    expect(ruler.tooltipText).toContain("1.50 in");

    // Switch to cm and check format
    ruler.unit = "cm";
    expect(ruler.formatMeasurement(1440)).toBe("2.54 cm");

    window.dispatchEvent(new PointerEvent("pointerup"));
    expect(ruler.showTooltip).toBe(false);
  });

  it("renders a marker element per tab stop (repeat binding)", async () => {
    // tabStops set before connect: the first render must already emit one
    // marker per stop (a plain array binding renders nothing).
    const ruler = new DocenRuler();
    created.push(ruler);
    ruler.tabStops = [
      { position: 1440, type: "left" },
      { position: 2880, type: "right" },
    ];
    document.body.append(ruler);
    await settle();
    expect(ruler.shadowRoot!.querySelectorAll(".tab-stop-item").length).toBe(2);

    const empty = new DocenRuler();
    created.push(empty);
    document.body.append(empty);
    await settle();
    expect(empty.shadowRoot!.querySelectorAll(".tab-stop-item").length).toBe(0);
  });

  it("synchronizes two-way with ProseMirror active paragraph", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [
          {
            type: "paragraph",
            attrs: {
              indent: { left: 720, right: 360, firstLine: 240 },
              tabStops: [{ position: 1440, type: "left" }],
            },
            content: [{ type: "text", text: "Paragraph with indents" }],
          },
        ],
      },
    });

    const ruler = await mountRuler();
    ruler.bindEditor(editor);

    // Active paragraph attrs are read into ruler
    expect(ruler.leftIndentTwips).toBe(720);
    expect(ruler.rightIndentTwips).toBe(360);
    expect(ruler.firstLineTwips).toBe(240);
    expect(ruler.tabStops).toHaveLength(1);

    // Changing markers on ruler commits to ProseMirror document
    ruler.leftIndentTwips = 1440;
    ruler.firstLineTwips = 720;
    ruler.commitIndentToEditor();

    const paraAttrs = editor.state.doc.child(0).attrs;
    expect(paraAttrs.indent?.left).toBe(1440);
    expect(paraAttrs.indent?.firstLine).toBe(720);
  });

  it("renders the four-level tick hierarchy crisply with signed margin numbers", async () => {
    const ruler = await mountRuler();
    ruler.unit = "in";
    ruler.pageWidthPx = 816;
    ruler.marginLeftPx = 96;
    ruler.marginRightPx = 96;
    ruler.scale = 1;
    ruler.renderTicks();

    const svg = ruler.shadowRoot!.querySelector(".ticks-svg") as SVGSVGElement;
    expect(svg.getAttribute("shape-rendering")).toBe("crispEdges");
    const ticks = Array.from(svg.querySelectorAll("line")).map((line) => ({
      // Lines sit on the half-pixel crisp grid; compare whole pixels.
      x: Math.floor(Number(line.getAttribute("x1"))),
      len: Number(line.getAttribute("y1")) - Number(line.getAttribute("y2")),
    }));
    // Word's hierarchy: major (unit, 7px), half (5px), quarter (3px), eighth (2px).
    expect(new Set(ticks.map((t) => t.len))).toEqual(new Set([7, 5, 3, 2]));
    // Eight subdivisions per inch at 100%: one every 12px.
    const oneInch = ticks.filter((t) => t.x >= 96 && t.x < 192).map((t) => t.x);
    expect(oneInch).toEqual([96, 108, 120, 132, 144, 156, 168, 180]);
    // The margin gutter carries negative unit labels.
    expect(svg.textContent).toContain("-1");
  });

  it("ignores clicks on markers, and double-clicking the ruler opens Tabs without a stray stop", async () => {
    const ruler = await mountRuler();
    ruler.marginLeftPx = 96;
    ruler.pageWidthPx = 816;
    const track = ruler.shadowRoot!.querySelector(".ruler-track") as HTMLElement;
    track.getBoundingClientRect = () => ({
      left: 0,
      top: 0,
      width: 816,
      height: 20,
      right: 816,
      bottom: 20,
      x: 0,
      y: 0,
      toJSON: () => {},
    });
    const marker = ruler.shadowRoot!.querySelector(".first-line-marker") as HTMLElement;

    // Clicking an indent marker must not create a tab stop.
    ruler.onTrackClick({ clientX: 200, currentTarget: track, target: marker } as any);
    expect(ruler.tabStops).toHaveLength(0);

    let opened = 0;
    ruler.addEventListener("ruler:open-tabs", () => {
      opened++;
    });

    // Single click adds a stop; the double-click that follows removes the
    // auto-added stop and opens the Tabs dialog.
    ruler.onTrackClick({ clientX: 240, currentTarget: track, target: track } as any);
    expect(ruler.tabStops).toHaveLength(1);
    ruler.onTrackDblClick({ target: track } as any);
    expect(ruler.tabStops).toHaveLength(0);
    expect(opened).toBe(1);
  });

  it("applies ruler edits to every selected paragraph (Word's selection scope)", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [
          { type: "paragraph", content: [{ type: "text", text: "First paragraph" }] },
          { type: "paragraph", content: [{ type: "text", text: "Second paragraph" }] },
        ],
      },
    });
    const ruler = await mountRuler();
    ruler.bindEditor(editor);
    editor.commands.setTextSelection({ from: 1, to: editor.state.doc.content.size - 1 });

    ruler.leftIndentTwips = 720;
    ruler.commitIndentToEditor();
    expect(editor.state.doc.child(0).attrs.indent?.left).toBe(720);
    expect(editor.state.doc.child(1).attrs.indent?.left).toBe(720);

    ruler.tabStops = [{ position: 1440, type: "left" }];
    ruler.commitTabStopsToEditor();
    expect(editor.state.doc.child(0).attrs.tabStops?.[0]?.position).toBe(1440);
    expect(editor.state.doc.child(1).attrs.tabStops?.[0]?.position).toBe(1440);
  });

  it("commits indent changes live while dragging (Word reflows during the drag)", async () => {
    const editor = new Editor({
      element: null,
      extensions: docxExtensions,
      content: {
        type: "doc",
        content: [{ type: "paragraph", content: [{ type: "text", text: "Live drag" }] }],
      },
    });
    const ruler = await mountRuler();
    ruler.marginLeftPx = 96;
    ruler.pageWidthPx = 816;
    ruler.bindEditor(editor);

    const firstLineEl = ruler.shadowRoot!.querySelector(".first-line-marker") as HTMLElement;
    firstLineEl.setPointerCapture = () => {};
    ruler.onMarkerPointerDown("firstLine", {
      preventDefault: () => {},
      stopPropagation: () => {},
      clientX: 96,
      pointerId: 7,
      target: firstLineEl,
    } as any);

    window.dispatchEvent(new PointerEvent("pointermove", { clientX: 144 }));
    // The live commit rides requestAnimationFrame.
    await new Promise((resolve) => setTimeout(resolve, 40));
    expect(editor.state.doc.child(0).attrs.indent?.firstLine).toBe(720);

    window.dispatchEvent(new PointerEvent("pointerup"));
    expect(ruler.dragActiveMarker).toBeNull();
  });

  it("mirrors ticks, content panel and markers for RTL paragraphs", async () => {
    const ruler = new DocenRuler();
    created.push(ruler);
    ruler.dir = "rtl";
    ruler.pageWidthPx = 816;
    ruler.marginLeftPx = 72; // physical left
    ruler.marginRightPx = 120; // physical right = start margin in RTL
    ruler.scale = 1;
    ruler.setParagraphAttrs({ left: 720, firstLine: 360, right: 1440 });
    document.body.append(ruler);
    await settle();

    expect(ruler.zeroXPx).toBe(696); // 816 - 120
    expect(ruler.contentStartPx).toBe(72); // margins stay physical in RTL
    expect(ruler.hangingMarkerX).toBe(648); // zero - 48
    expect(ruler.firstLineMarkerX).toBe(624); // zero - 72
    expect(ruler.rightMarkerX).toBe(168); // marginLeft + 96
    const contentBg = ruler.shadowRoot!.querySelector(".content-bg") as HTMLElement;
    expect(contentBg.style.left).toBe("72px");
    expect(contentBg.style.width).toBe("624px");
  });
});
