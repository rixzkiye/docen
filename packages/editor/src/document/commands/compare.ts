import type { DeduplicateOptions } from "@docen/deduplicate";
import type { JSONContent } from "@docen/docx";

export interface CompareOptions {
  author?: string;
  date?: string;
  dedupOptions?: DeduplicateOptions;
}

export interface DiffToken {
  text: string;
  type: "same" | "ins" | "del";
}

/**
 * Tokenize a text string into words, whitespace, and punctuation for granular diffing.
 */
export function tokenize(text: string): string[] {
  if (!text) return [];
  const tokens: string[] = [];
  const re = /\s+|[\p{L}\p{N}]+|[^\s\p{L}\p{N}]+/gu;
  let match: RegExpExecArray | null;
  while ((match = re.exec(text)) !== null) {
    tokens.push(match[0]);
  }
  return tokens;
}

/**
 * Compute LCS (Longest Common Subsequence) diff between two token arrays.
 */
export function diffTokens(tokensA: string[], tokensB: string[]): DiffToken[] {
  const m = tokensA.length;
  const n = tokensB.length;
  // DP table for LCS lengths
  const dp: number[][] = Array.from({ length: m + 1 }, () =>
    Array.from({ length: n + 1 }, () => 0),
  );

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (tokensA[i - 1] === tokensB[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  // Backtrack to assemble diff tokens
  let i = m;
  let j = n;
  const result: DiffToken[] = [];

  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && tokensA[i - 1] === tokensB[j - 1]) {
      result.push({ text: tokensA[i - 1], type: "same" });
      i--;
      j--;
    } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
      result.push({ text: tokensB[j - 1], type: "ins" });
      j--;
    } else if (i > 0 && (j === 0 || dp[i][j - 1] < dp[i - 1][j])) {
      result.push({ text: tokensA[i - 1], type: "del" });
      i--;
    }
  }

  result.reverse();

  // Merge consecutive tokens of same type
  const merged: DiffToken[] = [];
  for (const item of result) {
    const last = merged[merged.length - 1];
    if (last && last.type === item.type) {
      last.text += item.text;
    } else {
      merged.push({ ...item });
    }
  }

  return merged;
}

function getNodeText(node: JSONContent): string {
  if (!node) return "";
  if (node.type === "text" && node.text) return node.text;
  if (Array.isArray(node.content)) {
    return node.content.map(getNodeText).join("");
  }
  return "";
}

function isTextBlock(block: JSONContent): boolean {
  return block.type === "paragraph" || block.type === "heading";
}

function blockSimilarity(a: JSONContent, b: JSONContent): number {
  if (a.type !== b.type) return 0;
  if (isTextBlock(a)) {
    const textA = getNodeText(a);
    const textB = getNodeText(b);
    if (textA === textB) return 1.0;
    if (!textA && !textB) return 1.0;
    if (!textA || !textB) return 0;
    const tokensA = tokenize(textA.toLowerCase()).filter((t) => /\S/.test(t));
    const tokensB = tokenize(textB.toLowerCase()).filter((t) => /\S/.test(t));
    if (tokensA.length === 0 && tokensB.length === 0) return 1.0;
    if (tokensA.length === 0 || tokensB.length === 0) return 0;
    const setA = new Set(tokensA);
    let common = 0;
    for (const t of tokensB) {
      if (setA.has(t)) common++;
    }
    return (2 * common) / (tokensA.length + tokensB.length);
  }
  return JSON.stringify(a) === JSON.stringify(b) ? 1.0 : 0.8;
}

interface AlignmentStep {
  type: "match" | "del" | "ins";
  orig?: JSONContent;
  rev?: JSONContent;
}

function alignBlocks(origBlocks: JSONContent[], revBlocks: JSONContent[]): AlignmentStep[] {
  const m = origBlocks.length;
  const n = revBlocks.length;
  const dp: number[][] = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const sim = blockSimilarity(origBlocks[i - 1], revBlocks[j - 1]);
      const matchScore = sim >= 0.2 ? sim * 2 : -1;
      dp[i][j] = Math.max(dp[i - 1][j - 1] + matchScore, dp[i - 1][j] - 0.5, dp[i][j - 1] - 0.5);
    }
  }

  let i = m;
  let j = n;
  const steps: AlignmentStep[] = [];

  while (i > 0 || j > 0) {
    if (i > 0 && j > 0) {
      const sim = blockSimilarity(origBlocks[i - 1], revBlocks[j - 1]);
      const matchScore = sim >= 0.2 ? sim * 2 : -1;
      if (dp[i][j] === dp[i - 1][j - 1] + matchScore && sim >= 0.2) {
        steps.push({ type: "match", orig: origBlocks[i - 1], rev: revBlocks[j - 1] });
        i--;
        j--;
        continue;
      }
    }
    if (j > 0 && (i === 0 || dp[i][j] === dp[i][j - 1] - 0.5 || dp[i][j - 1] >= dp[i - 1][j])) {
      steps.push({ type: "ins", rev: revBlocks[j - 1] });
      j--;
    } else if (i > 0) {
      steps.push({ type: "del", orig: origBlocks[i - 1] });
      i--;
    }
  }

  return steps.reverse();
}

/**
 * Compare two documents and return a new JSONContent document showing tracked changes
 * (insertions and deletions with author and date metadata), preserving all block types.
 */
export function compareDocs(
  originalDoc: JSONContent,
  revisedDoc: JSONContent,
  options: CompareOptions = {},
): JSONContent {
  const author = options.author || "Comparison";
  const date = options.date || new Date().toISOString().replace(/\.\d+Z$/, "Z");
  let nextRevisionId = 1;

  const origBlocks = originalDoc.content ?? [];
  const revBlocks = revisedDoc.content ?? [];

  const steps = alignBlocks(origBlocks, revBlocks);
  const outputContent: JSONContent[] = [];

  for (const step of steps) {
    if (step.type === "match" && step.orig && step.rev) {
      if (isTextBlock(step.orig) && isTextBlock(step.rev)) {
        const textA = getNodeText(step.orig);
        const textB = getNodeText(step.rev);
        if (textA === textB) {
          outputContent.push(structuredClone(step.rev));
        } else {
          const tokensA = tokenize(textA);
          const tokensB = tokenize(textB);
          const diffs = diffTokens(tokensA, tokensB);
          const runs: JSONContent[] = [];
          for (const token of diffs) {
            if (token.type === "same") {
              runs.push({ type: "text", text: token.text });
            } else if (token.type === "del") {
              runs.push({
                type: "text",
                text: token.text,
                marks: [
                  {
                    type: "deletion",
                    attrs: { id: nextRevisionId++, author, date },
                  },
                ],
              });
            } else if (token.type === "ins") {
              runs.push({
                type: "text",
                text: token.text,
                marks: [
                  {
                    type: "insertion",
                    attrs: { id: nextRevisionId++, author, date },
                  },
                ],
              });
            }
          }
          outputContent.push({
            ...step.rev,
            content: runs,
          });
        }
      } else {
        // Non-text blocks (table, bulletList, orderedList, etc.) preserved
        outputContent.push(structuredClone(step.rev));
      }
    } else if (step.type === "del" && step.orig) {
      if (isTextBlock(step.orig)) {
        const deletedBlock: JSONContent = {
          ...step.orig,
          content: (step.orig.content ?? []).map((run) => ({
            ...run,
            marks: [
              ...(run.marks ?? []),
              {
                type: "deletion",
                attrs: { id: nextRevisionId++, author, date },
              },
            ],
          })),
        };
        outputContent.push(deletedBlock);
      } else {
        outputContent.push(structuredClone(step.orig));
      }
    } else if (step.type === "ins" && step.rev) {
      if (isTextBlock(step.rev)) {
        const insertedBlock: JSONContent = {
          ...step.rev,
          content: (step.rev.content ?? []).map((run) => ({
            ...run,
            marks: [
              ...(run.marks ?? []),
              {
                type: "insertion",
                attrs: { id: nextRevisionId++, author, date },
              },
            ],
          })),
        };
        outputContent.push(insertedBlock);
      } else {
        outputContent.push(structuredClone(step.rev));
      }
    }
  }

  return {
    type: "doc",
    content: outputContent.length > 0 ? outputContent : structuredClone(revisedDoc.content ?? []),
  };
}

/**
 * Combine multiple reviewers' revisions into a single document base.
 */
export function combineDocs(
  originalDoc: JSONContent,
  revisions: Array<{ doc: JSONContent; author: string }>,
): JSONContent {
  let base = structuredClone(originalDoc);
  for (const rev of revisions) {
    base = compareDocs(base, rev.doc, { author: rev.author });
  }
  return base;
}
