import type { Editor } from "@docen/docx/core";
import { TextSelection } from "@tiptap/pm/state";

export const enum ExtendLevel {
  Off = 0,
  Active = 1,
  Word = 2,
  Sentence = 3,
  Paragraph = 4,
  Document = 5,
}

export class ExtendModeManager {
  private _level: ExtendLevel = ExtendLevel.Off;
  private _anchor: number = 0;
  private _history: Array<{ from: number; to: number }> = [];

  get level(): ExtendLevel {
    return this._level;
  }

  get isActive(): boolean {
    return this._level !== ExtendLevel.Off;
  }

  get anchor(): number {
    return this._anchor;
  }

  statusLabel(): string {
    return this.isActive ? "EXT" : "";
  }

  cancel(): void {
    this._level = ExtendLevel.Off;
    this._anchor = 0;
    this._history = [];
  }

  /**
   * Advances extend mode level (F8 in Word).
   * Level 1: Extend mode armed (point anchor).
   * Level 2: Select word.
   * Level 3: Select sentence.
   * Level 4: Select paragraph.
   * Level 5: Select document.
   */
  step(editor: Editor): ExtendLevel {
    const { doc, selection } = editor.state;
    if (this._level === ExtendLevel.Off) {
      this._anchor = selection.from;
      this._level = ExtendLevel.Active;
      this._history = [{ from: selection.from, to: selection.to }];
      return this._level;
    }

    if (this._level === ExtendLevel.Active) {
      // Step to Word
      const range = this.getWordRange(doc, this._anchor);
      this._level = ExtendLevel.Word;
      this.applyRange(editor, range);
      return this._level;
    }

    if (this._level === ExtendLevel.Word) {
      // Step to Sentence
      const range = this.getSentenceRange(doc, this._anchor);
      this._level = ExtendLevel.Sentence;
      this.applyRange(editor, range);
      return this._level;
    }

    if (this._level === ExtendLevel.Sentence) {
      // Step to Paragraph
      const range = this.getParagraphRange(doc, this._anchor);
      this._level = ExtendLevel.Paragraph;
      this.applyRange(editor, range);
      return this._level;
    }

    // Step to Document
    this._level = ExtendLevel.Document;
    const start = TextSelection.near(doc.resolve(0), 1).from;
    const end = TextSelection.near(doc.resolve(doc.content.size), -1).to;
    this.applyRange(editor, { from: start, to: end });
    return this._level;
  }

  /**
   * Shrinks extend mode level (Shift+F8 in Word).
   */
  shrink(editor: Editor): ExtendLevel {
    if (this._level <= ExtendLevel.Active) {
      this.cancel();
      return ExtendLevel.Off;
    }

    if (this._history.length > 1) {
      this._history.pop(); // pop current
      const prev = this._history[this._history.length - 1]!;
      this._level = Math.max(ExtendLevel.Active, (this._level - 1) as ExtendLevel);
      editor.view.dispatch(
        editor.state.tr.setSelection(TextSelection.create(editor.state.doc, prev.from, prev.to)),
      );
      return this._level;
    }

    this.cancel();
    return ExtendLevel.Off;
  }

  private applyRange(editor: Editor, range: { from: number; to: number }): void {
    this._history.push(range);
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, range.from, range.to)),
    );
  }

  private getWordRange(doc: any, pos: number): { from: number; to: number } {
    const $pos = doc.resolve(pos);
    const parent = $pos.parent;
    const offset = $pos.parentOffset;
    const text = parent.textBetween(0, parent.content.size, undefined, " ");
    let start = offset;
    let end = offset;

    // Scan backward to word start
    while (start > 0 && /[\w\d]/.test(text[start - 1] ?? "")) {
      start--;
    }
    // Scan forward to word end
    while (end < text.length && /[\w\d]/.test(text[end] ?? "")) {
      end++;
    }

    const base = pos - offset;
    return { from: base + start, to: base + end };
  }

  private getSentenceRange(doc: any, pos: number): { from: number; to: number } {
    const $pos = doc.resolve(pos);
    const parent = $pos.parent;
    const offset = $pos.parentOffset;
    const text = parent.textBetween(0, parent.content.size, undefined, " ");
    let start = offset;
    let end = offset;

    while (start > 0 && !/[.!?]\s/.test(text.slice(start - 2, start))) {
      start--;
    }
    while (end < text.length && !/[.!?]/.test(text[end] ?? "")) {
      end++;
    }
    if (end < text.length && /[.!?]/.test(text[end] ?? "")) {
      end++;
    }

    const base = pos - offset;
    return { from: base + start, to: base + end };
  }

  private getParagraphRange(doc: any, pos: number): { from: number; to: number } {
    const $pos = doc.resolve(pos);
    const base = pos - $pos.parentOffset;
    return { from: base, to: base + $pos.parent.content.size };
  }
}
