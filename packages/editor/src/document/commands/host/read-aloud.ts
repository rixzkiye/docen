/**
 * Read Aloud Controller (Native Web Speech API)
 *
 * Implements Word's Read Aloud authoring tool: reads the active text
 * aloud paragraph-by-paragraph using the browser's native SpeechSynthesis API.
 */

export interface ReadAloudOptions {
  rate?: number;
  pitch?: number;
  lang?: string;
}

export class ReadAloudController {
  #synth: SpeechSynthesis | null = null;
  #currentUtterance: SpeechSynthesisUtterance | null = null;
  #playing = false;
  #paused = false;
  #rate = 1.0;
  #pitch = 1.0;

  constructor() {
    if (typeof window !== "undefined" && "speechSynthesis" in window) {
      this.#synth = window.speechSynthesis;
    }
  }

  isAvailable(): boolean {
    return this.#synth !== null;
  }

  currentUtterance(): SpeechSynthesisUtterance | null {
    return this.#currentUtterance;
  }

  isPlaying(): boolean {
    return this.#playing;
  }

  isPaused(): boolean {
    return this.#paused;
  }

  rate(): number {
    return this.#rate;
  }

  setRate(rate: number): void {
    this.#rate = Math.max(0.5, Math.min(rate, 2.0));
  }

  speak(text: string, opts?: ReadAloudOptions): void {
    if (!this.#synth || !text.trim()) return;

    this.stop();

    const utterance = new SpeechSynthesisUtterance(text.trim());
    utterance.rate = opts?.rate ?? this.#rate;
    utterance.pitch = opts?.pitch ?? this.#pitch;
    if (opts?.lang) utterance.lang = opts.lang;

    utterance.onend = () => {
      this.#playing = false;
      this.#paused = false;
      this.#currentUtterance = null;
    };

    utterance.onerror = () => {
      this.#playing = false;
      this.#paused = false;
      this.#currentUtterance = null;
    };

    this.#currentUtterance = utterance;
    this.#playing = true;
    this.#paused = false;
    this.#synth.speak(utterance);
  }

  pause(): void {
    if (this.#synth && this.#playing && !this.#paused) {
      this.#synth.pause();
      this.#paused = true;
    }
  }

  resume(): void {
    if (this.#synth && this.#paused) {
      this.#synth.resume();
      this.#paused = false;
    }
  }

  stop(): void {
    if (this.#synth) {
      this.#synth.cancel();
      this.#playing = false;
      this.#paused = false;
      this.#currentUtterance = null;
    }
  }
}
