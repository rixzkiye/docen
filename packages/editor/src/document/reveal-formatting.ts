import type { SectionPropertiesOptions } from "@docen/docx";
import type { Editor } from "@docen/docx/core";

import type { FormattingInfo } from "../ui/components/workspace/reveal-formatting-pane";

/** The Reveal Formatting snapshot for a selection — shared by the pane feed
 *  and the compare-to-selection reference capture. Pure editor read: the host
 *  and the tests both call it. */
export function formattingInfoOf(editor: Editor): FormattingInfo {
  const { $from, empty } = editor.state.selection;
  const sampleText = empty
    ? $from.parent.textBetween(
        Math.max(0, $from.parentOffset - 15),
        Math.min($from.parent.content.size, $from.parentOffset + 15),
        " ",
      )
    : editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, " ");

  const marks = $from.marks();
  const markTypes = new Set(marks.map((m) => m.type.name));
  const fontMark = marks.find((m) => m.type.name === "textStyle" || m.attrs?.fontFamily);
  const colorMark = marks.find((m) => m.attrs?.color);

  const para = $from.parent;
  const paraAttrs = (para.attrs ?? {}) as Record<string, any>;

  let activeSecProps: SectionPropertiesOptions | undefined;
  let foundSection = false;
  editor.state.doc.descendants((node, pos) => {
    if (foundSection) return false;
    if (
      node.type.name === "paragraph" &&
      (node.attrs as { sectionProperties?: unknown }).sectionProperties != null
    ) {
      if (pos >= $from.pos) {
        activeSecProps = (node.attrs as { sectionProperties?: SectionPropertiesOptions })
          .sectionProperties;
        foundSection = true;
        return false;
      }
    }
  });
  if (!activeSecProps) {
    activeSecProps = (editor.state.doc.attrs as { sectionProperties?: SectionPropertiesOptions })
      ?.sectionProperties;
  }

  const pageSize =
    activeSecProps?.pageSize && typeof activeSecProps.pageSize === "object"
      ? activeSecProps.pageSize
      : undefined;
  const pageMargin =
    activeSecProps?.pageMargin && typeof activeSecProps.pageMargin === "object"
      ? activeSecProps.pageMargin
      : undefined;

  let orientationStr = "Portrait";
  if (pageSize?.orientation === "landscape") {
    orientationStr = "Landscape";
  }
  let marginsStr = "Normal (1 in)";
  if (pageMargin) {
    const { top, left } = pageMargin;
    if (top != null && left != null) {
      const topIn = (Number(top) / 1440).toFixed(1);
      const leftIn = (Number(left) / 1440).toFixed(1);
      marginsStr = `Top: ${topIn}", Left: ${leftIn}"`;
    }
  }
  let paperSizeStr: string | undefined;
  if (pageSize?.width && pageSize?.height) {
    const wIn = (Number(pageSize.width) / 1440).toFixed(1);
    const hIn = (Number(pageSize.height) / 1440).toFixed(1);
    paperSizeStr = `${wIn}" × ${hIn}"`;
  }

  return {
    sampleText: sampleText.trim() || "Selected text",
    font: {
      family: (fontMark?.attrs?.fontFamily as string) || "Calibri",
      size: (fontMark?.attrs?.fontSize as string) || "11 pt",
      bold: markTypes.has("bold"),
      italic: markTypes.has("italic"),
      underline: markTypes.has("underline"),
      color: (colorMark?.attrs?.color as string) || "Auto",
    },
    paragraph: {
      alignment: (paraAttrs.textAlign as string) || "Left",
      indentLeft: paraAttrs.indentLeft != null ? `${paraAttrs.indentLeft} pt` : "0 pt",
      lineSpacing: paraAttrs.lineSpacing ? String(paraAttrs.lineSpacing) : "1.15",
    },
    section: {
      margins: marginsStr,
      orientation: orientationStr,
      paperSize: paperSizeStr,
    },
  };
}
