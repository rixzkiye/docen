import {
  compareDocuments,
  type DeduplicateOptions,
  type DocumentComparison,
} from "@docen/deduplicate";
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

/**
 * Compare two documents and return a new JSONContent document showing tracked changes
 * (insertions and deletions with author and date metadata).
 */
export function compareDocs(
  originalDoc: JSONContent,
  revisedDoc: JSONContent,
  options: CompareOptions = {},
): JSONContent {
  const author = options.author || "Comparison";
  const date = options.date || new Date().toISOString().replace(/\.\d+Z$/, "Z");
  let nextRevisionId = 1;

  const originalParagraphs = (originalDoc.content ?? []).filter((n) => n.type === "paragraph");
  const revisedParagraphs = (revisedDoc.content ?? []).filter((n) => n.type === "paragraph");

  const dedupResult: DocumentComparison = compareDocuments(
    originalDoc,
    revisedDoc,
    options.dedupOptions,
  );

  const matchedDoc2Indices = new Set<number>();
  const outputContent: JSONContent[] = [];

  // Iterate over comparisons
  for (let idx = 0; idx < dedupResult.paragraphs.length; idx++) {
    const pComp = dedupResult.paragraphs[idx];
    const origText = pComp.fromDoc1.text;
    let revMatch = pComp.fromDoc2;

    // Fallback: if no match from dedup, look for the most similar unmatched paragraph in revised
    if (!revMatch) {
      let bestIdx = -1;
      let bestScore = 0;
      const tokensA = new Set(tokenize(origText.toLowerCase()).filter((t) => /\w/.test(t)));
      for (let j = 0; j < revisedParagraphs.length; j++) {
        if (matchedDoc2Indices.has(j)) continue;
        const revP = revisedParagraphs[j];
        const textB = (revP.content ?? []).map((c) => c.text ?? "").join("");
        const tokensB = new Set(tokenize(textB.toLowerCase()).filter((t) => /\w/.test(t)));
        let common = 0;
        for (const t of tokensA) {
          if (tokensB.has(t)) common++;
        }
        const score =
          tokensA.size + tokensB.size > 0 ? (2 * common) / (tokensA.size + tokensB.size) : 0;
        if (score > bestScore && score >= 0.2) {
          bestScore = score;
          bestIdx = j;
        }
      }
      if (bestIdx >= 0) {
        const textB = (revisedParagraphs[bestIdx].content ?? []).map((c) => c.text ?? "").join("");
        revMatch = { index: bestIdx, text: textB };
      }
    }

    if (!revMatch) {
      // Entire paragraph deleted from original
      const origPara = originalParagraphs[pComp.fromDoc1.index];
      if (origPara) {
        const deletedPara: JSONContent = {
          ...origPara,
          content: (origPara.content ?? []).map((run) => ({
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
        outputContent.push(deletedPara);
      }
    } else {
      matchedDoc2Indices.add(revMatch.index);
      const revText = revMatch.text;
      const revPara = revisedParagraphs[revMatch.index] ?? { type: "paragraph", content: [] };

      if (origText === revText) {
        // Unmodified paragraph
        outputContent.push(structuredClone(revPara));
      } else {
        // Text inside paragraph modified — compute diff
        const tokensA = tokenize(origText);
        const tokensB = tokenize(revText);
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
          ...revPara,
          content: runs,
        });
      }
    }
  }

  // Handle any paragraphs in revisedDoc that were not matched to any paragraph in originalDoc (pure additions)
  for (let j = 0; j < revisedParagraphs.length; j++) {
    if (!matchedDoc2Indices.has(j)) {
      const addedPara = revisedParagraphs[j];
      outputContent.push({
        ...addedPara,
        content: (addedPara.content ?? []).map((run) => ({
          ...run,
          marks: [
            ...(run.marks ?? []),
            {
              type: "insertion",
              attrs: { id: nextRevisionId++, author, date },
            },
          ],
        })),
      });
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
