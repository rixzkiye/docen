import { getSchema } from "@tiptap/core";

import type { JSONContent } from "../core";
import { docxExtensions } from "../core";
import { parseHTMLBody } from "../extensions/paste";

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export interface HtmlGenerateOptions {
  fullDocument?: boolean;
  title?: string;
}

/**
 * Generate semantic, styled HTML from Tiptap JSONContent.
 */
export function generateHTML(
  doc: JSONContent | JSONContent[],
  options?: HtmlGenerateOptions,
): string {
  const nodes = Array.isArray(doc) ? doc : (doc.content ?? []);

  const renderInline = (node: JSONContent): string => {
    if (node.type === "hardBreak") return "<br>";
    if (node.type !== "text" || !node.text) return "";

    let html = escapeHtml(node.text);
    if (!node.marks || node.marks.length === 0) return html;

    for (const m of node.marks) {
      switch (m.type) {
        case "bold":
          html = `<strong>${html}</strong>`;
          break;
        case "italic":
          html = `<em>${html}</em>`;
          break;
        case "underline":
          html = `<u>${html}</u>`;
          break;
        case "strike":
          html = `<s>${html}</s>`;
          break;
        case "code":
          html = `<code>${html}</code>`;
          break;
        case "subscript":
          html = `<sub>${html}</sub>`;
          break;
        case "superscript":
          html = `<sup>${html}</sup>`;
          break;
        case "link": {
          const href = escapeHtml(String(m.attrs?.href ?? ""));
          const title = m.attrs?.title ? ` title="${escapeHtml(String(m.attrs.title))}"` : "";
          const target = m.attrs?.target ? ` target="${escapeHtml(String(m.attrs.target))}"` : "";
          html = `<a href="${href}"${target}${title}>${html}</a>`;
          break;
        }
        case "highlight": {
          const color = m.attrs?.color
            ? ` style="background-color: ${escapeHtml(String(m.attrs.color))}"`
            : "";
          html = `<mark${color}>${html}</mark>`;
          break;
        }
        case "textStyle": {
          const styles: string[] = [];
          if (m.attrs?.color) styles.push(`color: ${escapeHtml(String(m.attrs.color))}`);
          if (m.attrs?.fontSize) styles.push(`font-size: ${escapeHtml(String(m.attrs.fontSize))}`);
          if (m.attrs?.fontFamily)
            styles.push(`font-family: ${escapeHtml(String(m.attrs.fontFamily))}`);
          if (styles.length > 0) {
            html = `<span style="${styles.join("; ")}">${html}</span>`;
          }
          break;
        }
      }
    }
    return html;
  };

  const renderBlock = (node: JSONContent): string => {
    if (node.type === "paragraph") {
      const styles: string[] = [];
      if (node.attrs?.textAlign) styles.push(`text-align: ${node.attrs.textAlign}`);
      const styleAttr = styles.length > 0 ? ` style="${styles.join("; ")}"` : "";
      const content = (node.content ?? []).map(renderInline).join("");
      return `<p${styleAttr}>${content || "<br>"}</p>`;
    }

    if (node.type === "heading") {
      const level = Math.min(6, Math.max(1, Number(node.attrs?.level) || 1));
      const styles: string[] = [];
      if (node.attrs?.textAlign) styles.push(`text-align: ${node.attrs.textAlign}`);
      const styleAttr = styles.length > 0 ? ` style="${styles.join("; ")}"` : "";
      const content = (node.content ?? []).map(renderInline).join("");
      return `<h${level}${styleAttr}>${content}</h${level}>`;
    }

    if (node.type === "bulletList") {
      const content = (node.content ?? []).map(renderBlock).join("");
      return `<ul>${content}</ul>`;
    }

    if (node.type === "orderedList") {
      const content = (node.content ?? []).map(renderBlock).join("");
      return `<ol>${content}</ol>`;
    }

    if (node.type === "listItem") {
      const content = (node.content ?? [])
        .map((c) => {
          if (c.type === "paragraph") return (c.content ?? []).map(renderInline).join("");
          return renderBlock(c);
        })
        .join("");
      return `<li>${content}</li>`;
    }

    if (node.type === "table") {
      const content = (node.content ?? []).map(renderBlock).join("");
      return `<table style="border-collapse: collapse; width: 100%;">${content}</table>`;
    }

    if (node.type === "tableRow") {
      const content = (node.content ?? []).map(renderBlock).join("");
      return `<tr>${content}</tr>`;
    }

    if (node.type === "tableCell" || node.type === "tableHeader") {
      const tag = node.type === "tableHeader" ? "th" : "td";
      const attrs: string[] = [];
      if (node.attrs?.colspan && node.attrs.colspan > 1)
        attrs.push(`colspan="${node.attrs.colspan}"`);
      if (node.attrs?.rowspan && node.attrs.rowspan > 1)
        attrs.push(`rowspan="${node.attrs.rowspan}"`);
      const styles: string[] = ["border: 1px solid #d0d0d0", "padding: 4px 8px"];
      if (node.attrs?.background) styles.push(`background-color: ${node.attrs.background}`);
      attrs.push(`style="${styles.join("; ")}"`);
      const content = (node.content ?? [])
        .map((c) => {
          if (c.type === "paragraph") return (c.content ?? []).map(renderInline).join("");
          return renderBlock(c);
        })
        .join("");
      return `<${tag} ${attrs.join(" ")}>${content}</${tag}>`;
    }

    if (node.type === "image") {
      const src = escapeHtml(String(node.attrs?.src ?? ""));
      const alt = node.attrs?.alt ? ` alt="${escapeHtml(String(node.attrs.alt))}"` : "";
      const width = node.attrs?.width ? ` width="${node.attrs.width}"` : "";
      const height = node.attrs?.height ? ` height="${node.attrs.height}"` : "";
      return `<img src="${src}"${alt}${width}${height} />`;
    }

    if (node.type === "horizontalRule") {
      return "<hr />";
    }

    if (node.content) {
      return node.content.map(renderBlock).join("");
    }

    return "";
  };

  const bodyContent = nodes.map(renderBlock).join("\n");

  if (options?.fullDocument) {
    const title = options.title ? escapeHtml(options.title) : "Document";
    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title}</title>
  <style>
    body {
      font-family: Calibri, Arial, sans-serif;
      font-size: 11pt;
      line-height: 1.25;
      margin: 1in;
      color: #000000;
      background-color: #ffffff;
    }
    table {
      border-collapse: collapse;
      width: 100%;
      margin: 1em 0;
    }
    td, th {
      border: 1px solid #d0d0d0;
      padding: 6px 10px;
    }
    h1, h2, h3, h4, h5, h6 {
      margin-top: 1.2em;
      margin-bottom: 0.6em;
    }
    p {
      margin-top: 0;
      margin-bottom: 0.8em;
    }
    ul, ol {
      margin-top: 0;
      margin-bottom: 0.8em;
      padding-left: 2em;
    }
    mark {
      background-color: #ffff00;
    }
  </style>
</head>
<body>
${bodyContent}
</body>
</html>`;
  }

  return bodyContent;
}

/**
 * Parse an HTML string into Tiptap JSONContent.
 */
export async function parseHTML(html: string): Promise<JSONContent> {
  const schema = getSchema(docxExtensions);
  if (typeof document !== "undefined" && typeof DOMParser !== "undefined") {
    const body = new DOMParser().parseFromString(html, "text/html").body;
    return parseHTMLBody(body, schema);
  }

  // Node environment
  const { parseHTML: linkedomParse } = await import("linkedom");
  const { document: doc } = linkedomParse(`<!DOCTYPE html><html><body>${html}</body></html>`);
  return parseHTMLBody(doc.body as unknown as HTMLElement, schema);
}
