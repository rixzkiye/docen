/** Handler for one host command — `value` is the ribbon control's value attr. */
export type HostCommandHandler = (value?: string) => boolean;

/**
 * One per-domain slice of the host-command switch (`#onCommand` before the
 * wired-Tiptap fallback).
 *
 * `chrome` lists the events handled without a Tiptap editor (panes, zoom,
 * view mode — UI state the element owns); `editor` lists the events handled
 * only once a document has opened. `run` executes one event of this domain
 * and returns true when it consumed the command; false defers to the wired
 * dispatch (value-gated events such as a valueless `undo` step count).
 */
export interface HostCommandDomain {
  readonly chrome: readonly string[];
  readonly editor: readonly string[];
  run(event: string, value?: string): boolean;
}

/** The two event → handler tables `#onCommand` consults, in dispatch order. */
export interface HostCommandRegistry {
  readonly chrome: ReadonlyMap<string, HostCommandHandler>;
  readonly editor: ReadonlyMap<string, HostCommandHandler>;
}

/** Flatten the domain modules into the registry's lookup tables. */
export function hostRegistry(domains: readonly HostCommandDomain[]): HostCommandRegistry {
  const chrome = new Map<string, HostCommandHandler>();
  const editor = new Map<string, HostCommandHandler>();
  for (const domain of domains) {
    for (const event of domain.chrome) chrome.set(event, (value) => domain.run(event, value));
    for (const event of domain.editor) editor.set(event, (value) => domain.run(event, value));
  }
  return { chrome, editor };
}
