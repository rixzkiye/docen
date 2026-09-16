import type { JSONContent } from "../core";

/**
 * Generate plain text from Tiptap JSONContent.
 */
export function generatePlainText(doc: JSONContent | JSONContent[]): string {
  const nodes = Array.isArray(doc) ? doc : (doc.content ?? []);

  const extractText = (node: JSONContent): string => {
    if (node.type === "hardBreak") return "\n";
    if (node.type === "text") return node.text ?? "";

    if (node.type === "paragraph" || node.type === "heading") {
      const inner = (node.content ?? []).map(extractText).join("");
      return inner;
    }

    if (node.type === "listItem") {
      return (node.content ?? []).map(extractText).join("");
    }

    if (node.type === "tableRow") {
      return (node.content ?? []).map(extractText).join("\t");
    }

    if (node.content) {
      return node.content.map(extractText).join("\n");
    }

    return "";
  };

  const blocks: string[] = [];
  for (const node of nodes) {
    if (node.type === "bulletList") {
      const items = (node.content ?? []).map((li) => `• ${extractText(li)}`);
      blocks.push(items.join("\n"));
    } else if (node.type === "orderedList") {
      const items = (node.content ?? []).map((li, idx) => `${idx + 1}. ${extractText(li)}`);
      blocks.push(items.join("\n"));
    } else if (node.type === "table") {
      const rows = (node.content ?? []).map(extractText);
      blocks.push(rows.join("\n"));
    } else {
      blocks.push(extractText(node));
    }
  }

  return blocks.join("\n\n");
}

/**
 * Parse plain text into Tiptap JSONContent.
 */
export function parsePlainText(text: string): JSONContent {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const paragraphs = normalized.split("\n\n");

  const content: JSONContent[] = [];
  for (const para of paragraphs) {
    const lines = para.split("\n");
    const inlines: JSONContent[] = [];
    lines.forEach((line, idx) => {
      if (line) inlines.push({ type: "text", text: line });
      if (idx < lines.length - 1) inlines.push({ type: "hardBreak" });
    });
    content.push({
      type: "paragraph",
      content: inlines,
    });
  }

  if (content.length === 0) {
    content.push({ type: "paragraph", content: [] });
  }

  return { type: "doc", content };
}
