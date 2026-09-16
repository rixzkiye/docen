import {
  FASTElement,
  attr,
  css,
  customElement,
  html,
  observable,
  ref,
} from "@microsoft/fast-element";

import { observeLang, t } from "../../i18n/localize";

/** One comment card's data — the host flattens documentExtras.comments into
 *  this shape (children runs joined to a single text) and nests replies under
 *  their thread root through the commentsExtended paraIdParent chain. */
export interface CommentCard {
  id: number;
  author: string;
  initials: string;
  date: string;
  text: string;
  /** The thread's resolved state (w15:done on the commentsExtended entry). */
  resolved?: boolean;
  /** Replies to this comment (top-level cards have none themselves). */
  replies?: CommentCard[];
}

const styles = css`
  :host {
    display: flex;
    flex-direction: column;
    flex: 1;
    min-height: 0;
    box-sizing: border-box;
    font-size: 12px;
  }
  .toolbar {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    padding: 8px 6px;
    border-bottom: 1px solid var(--docen-color-divider, #e2e2e2);
    background: var(--docen-color-bg, #fff);
  }
  .toolbar-select {
    flex: 1;
    min-width: 85px;
    height: 24px;
    font-size: 11px;
    padding: 2px 4px;
    border: 1px solid var(--docen-color-stroke, #d1d1d1);
    border-radius: 4px;
    background: var(--docen-color-bg, #fff);
    color: var(--docen-color-text-1, #242424);
  }
  .toolbar-search {
    width: 100%;
    height: 24px;
    box-sizing: border-box;
    font-size: 11px;
    padding: 2px 6px;
    border: 1px solid var(--docen-color-stroke, #d1d1d1);
    border-radius: 4px;
    background: var(--docen-color-bg, #fff);
    color: var(--docen-color-text-1, #242424);
  }
  /* Card list — Word's comments pane: one card per comment, a rounded avatar
     with the author's initials, name + timestamp on one row, body under. */
  .list {
    flex: 1;
    min-height: 0;
    overflow: auto;
    padding: 6px;
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .empty {
    color: var(--docen-color-text-2, #616161);
    padding: 14px 8px;
    text-align: center;
  }
  .card {
    display: grid;
    grid-template-columns: 28px 1fr auto;
    grid-template-rows: auto auto;
    column-gap: 8px;
    padding: 8px;
    border-radius: 4px;
    cursor: pointer;
  }
  .card:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  /* The card whose anchored range the caret sits in — Word highlights it
     while the selection is inside the comment's text. */
  .card.active {
    background: var(--docen-color-subtle-selected, #e8f0fb);
  }
  .avatar {
    grid-row: 1 / 3;
    width: 28px;
    height: 28px;
    border-radius: 50%;
    background: var(--docen-color-accent, #0f6cbd);
    color: #fff;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 11px;
    font-weight: 600;
  }
  .meta {
    display: flex;
    align-items: baseline;
    gap: 6px;
    min-width: 0;
  }
  .author {
    font-weight: 600;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .when {
    color: var(--docen-color-text-2, #616161);
    white-space: nowrap;
    font-size: 11px;
  }
  .actions {
    display: none;
    gap: 2px;
    grid-row: 1;
  }
  .card:hover .actions {
    display: inline-flex;
  }
  .act {
    font-size: 11px;
    color: var(--docen-color-text-1, #242424);
  }
  .body {
    grid-column: 2 / 4;
    margin-top: 2px;
    white-space: pre-wrap;
    overflow-wrap: anywhere;
  }
  .mention {
    color: var(--docen-color-accent, #0f6cbd);
    font-weight: 600;
    background: var(--docen-color-subtle-selected, #e8f0fb);
    border-radius: 3px;
    padding: 0 3px;
  }
  .mention-popup {
    position: absolute;
    background: var(--docen-color-bg, #fff);
    border: 1px solid var(--docen-color-divider, #e2e2e2);
    border-radius: 4px;
    box-shadow: var(--shadow4, 0 4px 8px rgba(0, 0, 0, 0.14));
    max-height: 120px;
    overflow-y: auto;
    z-index: 10;
    width: 160px;
  }
  .mention-item {
    padding: 4px 8px;
    font-size: 11px;
    cursor: pointer;
  }
  .mention-item:hover {
    background: var(--docen-color-subtle-background-hover, #f5f5f5);
  }
  /* Inline edit state — the card's body becomes a text area with
     Save / Cancel (Word edits in place, not through a dialog). */
  .edit {
    grid-column: 2 / 4;
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin-top: 2px;
    position: relative;
  }
  .edit fluent-textarea {
    width: 100%;
    box-sizing: border-box;
  }
  .edit .row {
    display: flex;
    gap: 6px;
    justify-content: flex-end;
  }
  /* Reply thread — replies indent under the root's text column and share the
     card layout; a resolved thread dims (Word grays resolved conversations). */
  .thread {
    display: flex;
    flex-direction: column;
    gap: 6px;
  }
  .card.reply {
    margin-left: 36px;
  }
  .card.resolved {
    opacity: 0.55;
  }
  .resolved-badge {
    font-size: 10px;
    line-height: 16px;
    color: var(--docen-color-accent, #0f6cbd);
    border: 1px solid currentColor;
    border-radius: 3px;
    padding: 0 4px;
    white-space: nowrap;
  }
  .replybox {
    grid-column: 2 / 4;
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin-top: 2px;
    position: relative;
  }
  .replybox fluent-textarea {
    width: 100%;
    box-sizing: border-box;
  }
  .replybox .row {
    display: flex;
    gap: 6px;
    justify-content: flex-end;
  }
`;

function renderCommentBody(text: string): DocumentFragment {
  const frag = document.createDocumentFragment();
  const parts = text.split(/(@[a-zA-Z0-9_\u4e00-\u9fa5]+)/g);
  for (const part of parts) {
    if (part.startsWith("@") && part.length > 1) {
      const span = document.createElement("span");
      span.className = "mention";
      span.textContent = part;
      frag.append(span);
    } else if (part) {
      frag.append(document.createTextNode(part));
    }
  }
  return frag;
}

const template = html<DocenCommentsPane>`
  <div class="toolbar" ${ref("toolbarEl")}>
    <select class="toolbar-select" ${ref("statusFilterEl")} @change="${(x) => x.onFilterChange()}">
      <option value="all">${(x) => t("comments.filter.all", x)}</option>
      <option value="active">${(x) => t("comments.filter.active", x)}</option>
      <option value="resolved">${(x) => t("comments.filter.resolved", x)}</option>
    </select>
    <select class="toolbar-select" ${ref("authorFilterEl")} @change="${(x) => x.onFilterChange()}">
      <option value="all">${(x) => t("comments.filter.allAuthors", x)}</option>
    </select>
    <select class="toolbar-select" ${ref("sortEl")} @change="${(x) => x.onFilterChange()}">
      <option value="doc">${(x) => t("comments.sort.doc", x)}</option>
      <option value="newest">${(x) => t("comments.sort.newest", x)}</option>
      <option value="oldest">${(x) => t("comments.sort.oldest", x)}</option>
    </select>
    <input
      type="text"
      class="toolbar-search"
      ${ref("searchEl")}
      placeholder="${(x) => t("comments.search", x)}"
      @input="${(x) => x.onFilterChange()}"
    />
  </div>
  <div class="list" ${ref("listEl")}></div>
`;

/** `<docen-comments-pane comments active-id>` — Word's comments sidebar: one
 *  card per comment (initials avatar, author, timestamp, body) and an inline
 *  edit state per card. `comments` is JSON `CommentCard[]` in the order the
 *  host computed (document order); `active-id` highlights the card whose
 *  anchored range the selection sits in. Interactions emit `comment:select
 *  {id}` (card click — the host scrolls to the range), `comment:update
 *  {id,text}` and `comment:delete {id}`. New comments compose in the floating
 *  box beside the selection (host-owned), not in this pane. */
@customElement({ name: "docen-comments-pane", template, styles })
class DocenCommentsPane extends FASTElement {
  /** JSON CommentCard[] — the document's comments, ordered by the host. */
  @attr comments?: string;
  /** The comment id whose range covers the selection ("" = none). */
  @attr({ attribute: "active-id" }) activeId?: string;

  @observable toolbarEl?: HTMLElement;
  @observable statusFilterEl?: HTMLSelectElement;
  @observable authorFilterEl?: HTMLSelectElement;
  @observable sortEl?: HTMLSelectElement;
  @observable searchEl?: HTMLInputElement;
  @observable listEl?: HTMLElement;

  #unsubscribe?: () => void;

  commentsChanged(): void {
    this.#renderList();
  }

  activeIdChanged(): void {
    this.#highlightActive();
  }

  connectedCallback(): void {
    super.connectedCallback();
    this.#renderList();
    this.#applyI18n();
    this.#unsubscribe = observeLang(() => this.#applyI18n());
  }

  disconnectedCallback(): void {
    this.#unsubscribe?.();
    super.disconnectedCallback();
  }

  onFilterChange(): void {
    this.#renderList();
  }

  #updateAuthorFilterOptions(cards: CommentCard[]): void {
    const select = this.authorFilterEl;
    if (!select) return;
    const authors = new Set<string>();
    for (const c of cards) {
      if (c.author) authors.add(c.author);
      for (const r of c.replies ?? []) {
        if (r.author) authors.add(r.author);
      }
    }
    const currentVal = select.value || "all";
    select.replaceChildren();
    const allOpt = document.createElement("option");
    allOpt.value = "all";
    allOpt.textContent = t("comments.filter.allAuthors", this);
    select.append(allOpt);

    for (const author of [...authors].sort()) {
      const opt = document.createElement("option");
      opt.value = author;
      opt.textContent = author;
      select.append(opt);
    }
    if (authors.has(currentVal)) {
      select.value = currentVal;
    } else {
      select.value = "all";
    }
  }

  #setupMentionAutocomplete(area: HTMLElement & { value?: string }, wrap: HTMLElement): void {
    let popup: HTMLElement | null = null;
    const closePopup = () => {
      popup?.remove();
      popup = null;
    };
    area.addEventListener("input", () => {
      const val = area.value ?? "";
      const lastAt = val.lastIndexOf("@");
      if (lastAt < 0 || (lastAt > 0 && !/\s/.test(val.charAt(lastAt - 1)))) {
        closePopup();
        return;
      }
      const query = val.slice(lastAt + 1).toLowerCase();
      let cards: CommentCard[] = [];
      try {
        cards = this.comments ? (JSON.parse(this.comments) as CommentCard[]) : [];
      } catch {
        cards = [];
      }
      const authors = new Set<string>();
      for (const c of cards) {
        if (c.author) authors.add(c.author);
        for (const r of c.replies ?? []) if (r.author) authors.add(r.author);
      }
      const matches = [...authors].filter((a) => a.toLowerCase().includes(query));
      if (matches.length === 0) {
        closePopup();
        return;
      }
      if (!popup) {
        popup = document.createElement("div");
        popup.className = "mention-popup";
        wrap.append(popup);
      }
      popup.replaceChildren();
      for (const match of matches) {
        const item = document.createElement("div");
        item.className = "mention-item";
        item.textContent = `@${match}`;
        item.addEventListener("mousedown", (e) => {
          e.preventDefault();
          const before = val.slice(0, lastAt);
          area.value = `${before}@${match} `;
          closePopup();
          const input = (area.shadowRoot?.querySelector("textarea") ?? area) as HTMLElement | null;
          input?.focus();
        });
        popup.append(item);
      }
    });
    area.addEventListener("blur", () => {
      setTimeout(closePopup, 200);
    });
  }

  /** One card (root or reply). The edit state swaps the body for a text area
   *  + Save/Cancel; roots grow a Reply action and carry their replies below. */
  #renderCard(comment: CommentCard, frag: HTMLElement, isReply = false): void {
    const card = document.createElement("div");
    card.className = isReply ? "card reply" : "card";
    if (comment.resolved) card.classList.add("resolved");
    card.dataset.id = String(comment.id);
    card.addEventListener("click", (e) => {
      if ((e.target as HTMLElement).closest(".actions, .edit, .replybox")) return;
      this.dispatchEvent(
        new CustomEvent("comment:select", {
          bubbles: true,
          composed: true,
          detail: { id: comment.id },
        }),
      );
    });

    const avatar = document.createElement("div");
    avatar.className = "avatar";
    avatar.textContent = comment.initials || comment.author.slice(0, 2).toUpperCase();
    card.append(avatar);

    const meta = document.createElement("div");
    meta.className = "meta";
    const author = document.createElement("span");
    author.className = "author";
    author.textContent = comment.author;
    const when = document.createElement("span");
    when.className = "when";
    try {
      when.textContent = new Date(comment.date).toLocaleString();
    } catch {
      when.textContent = comment.date;
    }
    meta.append(author, when);
    if (comment.resolved) {
      const badge = document.createElement("span");
      badge.className = "resolved-badge";
      badge.textContent = t("comments.resolved", this);
      meta.append(badge);
    }
    card.append(meta);

    const actions = document.createElement("div");
    actions.className = "actions";
    if (!isReply) {
      const reply = document.createElement("fluent-button");
      reply.setAttribute("appearance", "subtle");
      reply.setAttribute("size", "small");
      reply.className = "act";
      reply.textContent = t("comments.reply", this);
      reply.addEventListener("click", () => this.#beginReply(comment, card));
      const resolve = document.createElement("fluent-button");
      resolve.setAttribute("appearance", "subtle");
      resolve.setAttribute("size", "small");
      resolve.className = "act";
      resolve.textContent = t(comment.resolved ? "comments.reopen" : "comments.resolve", this);
      resolve.addEventListener("click", () => {
        this.dispatchEvent(
          new CustomEvent("comment:resolve", {
            bubbles: true,
            composed: true,
            detail: { id: comment.id, done: !comment.resolved },
          }),
        );
      });
      actions.append(reply, resolve);
    }
    const edit = document.createElement("fluent-button");
    edit.setAttribute("appearance", "subtle");
    edit.setAttribute("size", "small");
    edit.className = "act";
    edit.textContent = t("comments.edit", this);
    edit.addEventListener("click", () => this.#beginEdit(comment, card));
    const del = document.createElement("fluent-button");
    del.setAttribute("appearance", "subtle");
    del.setAttribute("size", "small");
    del.className = "act";
    del.textContent = t("comments.delete", this);
    del.addEventListener("click", () => {
      this.dispatchEvent(
        new CustomEvent("comment:delete", {
          bubbles: true,
          composed: true,
          detail: { id: comment.id },
        }),
      );
    });
    actions.append(edit, del);
    card.append(actions);

    const body = document.createElement("div");
    body.className = "body";
    body.append(renderCommentBody(comment.text));
    card.append(body);

    frag.append(card);
  }

  /** The inline reply box under a thread root (Word's reply entry lives on
   *  the card, not a separate dialog). Posts `comment:reply {parentId,text}`. */
  #beginReply(comment: CommentCard, card: HTMLElement): void {
    if (card.querySelector(".replybox")) return;
    const wrap = document.createElement("div");
    wrap.className = "replybox";
    const area = document.createElement("fluent-textarea") as HTMLTextAreaElement & HTMLElement;
    area.setAttribute("block", "");
    area.setAttribute("rows", "2");
    area.setAttribute("placeholder", t("comments.placeholder", this));
    const row = document.createElement("div");
    row.className = "row";
    const cancel = document.createElement("fluent-button");
    cancel.setAttribute("appearance", "neutral");
    cancel.textContent = t("comments.cancel", this);
    cancel.addEventListener("click", () => wrap.remove());
    const post = document.createElement("fluent-button");
    post.setAttribute("appearance", "accent");
    post.textContent = t("comments.post", this);
    post.addEventListener("click", () => {
      const text = (area.value ?? "").trim();
      if (text)
        this.dispatchEvent(
          new CustomEvent("comment:reply", {
            bubbles: true,
            composed: true,
            detail: { parentId: comment.id, text },
          }),
        );
      wrap.remove();
    });
    area.addEventListener("keydown", (event: KeyboardEvent) => {
      if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        post.click();
      }
      if (event.key === "Escape") {
        event.preventDefault();
        wrap.remove();
      }
    });
    row.append(cancel, post);
    wrap.append(area, row);
    this.#setupMentionAutocomplete(area, wrap);
    card.append(wrap);
    requestAnimationFrame(() => {
      const input = (area.shadowRoot?.querySelector("textarea") ?? area) as HTMLElement | null;
      input?.focus();
    });
  }

  #beginEdit(comment: CommentCard, card: HTMLElement): void {
    if (card.querySelector(".edit")) return;
    card.querySelector(".body")?.remove();
    const wrap = document.createElement("div");
    wrap.className = "edit";
    const area = document.createElement("fluent-textarea") as HTMLTextAreaElement & HTMLElement;
    area.setAttribute("block", "");
    area.setAttribute("rows", "3");
    area.value = comment.text;
    const row = document.createElement("div");
    row.className = "row";
    const cancel = document.createElement("fluent-button");
    cancel.setAttribute("appearance", "neutral");
    cancel.textContent = t("comments.cancel", this);
    cancel.addEventListener("click", () => {
      this.#renderList();
    });
    const save = document.createElement("fluent-button");
    save.setAttribute("appearance", "accent");
    save.textContent = t("comments.save", this);
    save.addEventListener("click", () => {
      const text = (area.value ?? "").trim();
      if (text)
        this.dispatchEvent(
          new CustomEvent("comment:update", {
            bubbles: true,
            composed: true,
            detail: { id: comment.id, text },
          }),
        );
      this.#renderList();
    });
    area.addEventListener("keydown", (event: KeyboardEvent) => {
      if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        save.click();
      }
      if (event.key === "Escape") {
        event.preventDefault();
        this.#renderList();
      }
    });
    row.append(cancel, save);
    wrap.append(area, row);
    this.#setupMentionAutocomplete(area, wrap);
    card.append(wrap);
    requestAnimationFrame(() => {
      const input = (area.shadowRoot?.querySelector("textarea") ?? area) as HTMLElement | null;
      input?.focus();
    });
  }

  #renderList(): void {
    const list = this.listEl;
    if (!list) return;
    list.replaceChildren();
    let cards: CommentCard[] = [];
    try {
      cards = this.comments ? (JSON.parse(this.comments) as CommentCard[]) : [];
    } catch {
      cards = [];
    }

    this.#updateAuthorFilterOptions(cards);

    const status = this.statusFilterEl?.value ?? "all";
    const authorFilter = this.authorFilterEl?.value ?? "all";
    const search = (this.searchEl?.value ?? "").trim().toLowerCase();
    const sort = this.sortEl?.value ?? "doc";

    let filtered = cards.filter((card) => {
      if (status === "active" && card.resolved) return false;
      if (status === "resolved" && !card.resolved) return false;

      if (authorFilter !== "all") {
        const matchesAuthor =
          card.author === authorFilter ||
          (card.replies ?? []).some((r) => r.author === authorFilter);
        if (!matchesAuthor) return false;
      }

      if (search) {
        const textMatch =
          card.text.toLowerCase().includes(search) ||
          card.author.toLowerCase().includes(search) ||
          (card.replies ?? []).some(
            (r) => r.text.toLowerCase().includes(search) || r.author.toLowerCase().includes(search),
          );
        if (!textMatch) return false;
      }

      return true;
    });

    if (sort === "newest") {
      filtered = [...filtered].sort((a, b) => b.date.localeCompare(a.date));
    } else if (sort === "oldest") {
      filtered = [...filtered].sort((a, b) => a.date.localeCompare(b.date));
    }

    if (filtered.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty";
      empty.textContent = t("comments.empty", this);
      list.append(empty);
      return;
    }
    const frag = document.createDocumentFragment();
    for (const card of filtered) {
      const thread = document.createElement("div");
      thread.className = "thread";
      this.#renderCard(card, thread);
      for (const reply of card.replies ?? []) this.#renderCard(reply, thread, true);
      frag.append(thread);
    }
    list.append(frag);
    this.#highlightActive();
  }

  /** Paint the active card and bring it into the pane's viewport — the scroll
   *  stays inside .list (manual scrollTop), never the page behind the pane. */
  #highlightActive(): void {
    const list = this.listEl;
    if (!list) return;
    const id = this.activeId;
    for (const el of list.querySelectorAll<HTMLElement>(".card")) {
      const on = id != null && id !== "" && el.dataset.id === id;
      el.classList.toggle("active", on);
      if (on) list.scrollTop = el.offsetTop - list.clientHeight / 2 + el.clientHeight / 2;
    }
  }

  #applyI18n(): void {
    this.#renderList();
  }
}

export default DocenCommentsPane;
