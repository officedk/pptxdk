import type { DeclarationChild, DeclarationNode } from "#/declaration.ts";
import {
  insertDeclarationChild,
  removeDeclarationChild,
  replaceDeclarationProps,
} from "#/document-internals.ts";
import { snapshotValue } from "#/input-values.ts";
import type { TextStyle } from "#/theme.ts";
import type {
  DeckProps,
  LayoutStyle,
  SectionProps,
  SlideProps,
} from "#/types.ts";

export type DocumentProps = Readonly<Record<string, unknown>>;
export type DocumentStyle = Readonly<Record<string, unknown>>;

export type DocumentUpdate<Props extends object> = {
  [Key in keyof Omit<Props, "children" | "style">]?:
    | Omit<Props, "children" | "style">[Key]
    | undefined;
};

export type StyleUpdate<Style extends object> = {
  [Key in keyof Style]?: Style[Key] | undefined;
};

export type RichTextItem = string | CommonNode;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function declarationChild(item: RichTextItem): DeclarationChild {
  if (typeof item === "string") {
    return Object.freeze({ type: "#text", text: item });
  }
  if (item instanceof CommonNode) return item;
  throw new TypeError(
    "Rich text content must be a string or an editable document node",
  );
}

function requirePatchObject(value: unknown, label: string): void {
  if (!isRecord(value)) {
    throw new TypeError(`${label} must be an object`);
  }
}

function normalizeProps(
  props: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
  const copy: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(props)) {
    if (key === "children" || value === undefined) continue;
    Object.defineProperty(copy, key, {
      configurable: false,
      enumerable: true,
      value: snapshotValue(value),
      writable: false,
    });
  }
  return Object.freeze(copy);
}

/**
 * Common editable document node shared by command authoring and the React handoff.
 *
 * The node exposes a read-only declaration view to the formal build. Author mutations go through
 * `set`, `style`, and structure operations so parent ownership, value isolation, and update timing
 * stay in one place. Declaration getters return fresh isolated values; live nodes keep identity.
 */
export class CommonNode<
  Kind extends string = string,
  Props extends object = DocumentProps,
  Style extends object = DocumentStyle,
> implements DeclarationNode {
  readonly type: string;
  readonly kind: Kind;
  /** Page-affiliated declarations are reached below hidden canvas parents by the formal build. */
  readonly pageAffiliated: boolean;
  #props: Readonly<Record<string, unknown>>;
  #children: DeclarationChild[] = [];
  #childrenView: readonly DeclarationChild[] = Object.freeze([]);
  #parent: CommonNode | undefined;
  #owner: CommonNode | undefined;

  constructor(
    kind: Kind,
    type: string,
    props: Props = {} as Props,
    options: { readonly pageAffiliated?: boolean } = {},
  ) {
    this.kind = kind;
    this.type = type;
    this.pageAffiliated = options.pageAffiliated === true;
    this.#props = normalizeProps(props as Readonly<Record<string, unknown>>);
    if (kind === "deck") this.#owner = this;
  }

  /** Isolated declaration values for the formal build; mutating the returned value cannot edit the node. */
  get props(): Readonly<Record<string, unknown>> {
    return snapshotValue(this.#props);
  }

  /** Author id declared on this node, ignoring non-string values rejected by the formal build. */
  get id(): string | undefined {
    const value = this.#props["id"];
    return typeof value === "string" && value.trim() !== "" ? value : undefined;
  }

  /** Current parent container, or undefined while detached. */
  get parent(): CommonNode | undefined {
    return this.#parent;
  }

  /** Scoped author query rooted at this node; it searches children and stops at the next named scope. */
  get query(): AuthorQuery<this> {
    return new AuthorQuery(this);
  }

  /** Read-only child declaration view; node entries keep their editable identity. */
  get children(): readonly DeclarationChild[] {
    return this.#childrenView;
  }

  /**
   * Adapter-only complete declaration replacement for one React commit.
   *
   * React commits all current JSX props, so absent fields must disappear instead of surviving from a
   * previous commit as they would under the author-facing single-level update. Values still pass
   * through the same field-level input snapshot as construction.
   */
  [replaceDeclarationProps](props: Readonly<Record<string, unknown>>): void {
    this.#props = normalizeProps(props);
  }

  /**
   * Adapter-only insertion or same-document reordering during React reconciliation.
   *
   * React may reposition an attached host instance with `insertBefore`; author `append` remains
   * strict and the adapter reaches this operation explicitly. Text entries are accepted so rich-text
   * containers can be reconciled in place without rebuilding their content.
   */
  [insertDeclarationChild](
    child: DeclarationChild,
    before?: DeclarationChild,
  ): void {
    if (before === child) return;
    if (child instanceof CommonNode) {
      if (child === this) {
        throw new Error("Cannot attach a node inside its own descendant tree");
      }
      if (child.#parent === this) {
        const index = this.#children.indexOf(child);
        if (index !== -1) this.#children.splice(index, 1);
      } else {
        this.assertCanInsert(child, true);
        const previousParent = child.#parent;
        if (previousParent !== undefined) previousParent.#detach(child);
        const owner = this.#owner;
        if (owner !== undefined) child.#assignOwner(owner);
        child.#parent = this;
      }
    } else {
      const index = this.#children.indexOf(child);
      if (index !== -1) this.#children.splice(index, 1);
    }
    const index = before === undefined ? -1 : this.#children.indexOf(before);
    if (index === -1) this.#children.push(child);
    else this.#children.splice(index, 0, child);
    this.#refreshChildren();
  }

  /** Adapter-only removal during React reconciliation; author removal remains `remove()`. */
  [removeDeclarationChild](child: DeclarationChild): void {
    this.#detach(child);
  }

  /** Applies an ordinary property update; `style` and `children` use their dedicated operations. */
  set(patch: DocumentUpdate<Props>): this {
    const record = patch as Readonly<Record<string, unknown>>;
    requirePatchObject(record, "set() patch");
    if ("style" in record || "children" in record) {
      throw new TypeError(
        "set() cannot update style or children; use style() and structural operations",
      );
    }
    const next = { ...this.#props };
    for (const [key, value] of Object.entries(record)) {
      if (value === undefined) delete next[key];
      else next[key] = snapshotValue(value);
    }
    this.#props = Object.freeze(next);
    return this;
  }

  /** Applies a single-level local style update; nested values replace and undefined removes a local declaration. */
  style(patch: StyleUpdate<Style>): this {
    const record = patch as Readonly<Record<string, unknown>>;
    requirePatchObject(record, "style() patch");
    const current = isRecord(this.#props["style"]) ? this.#props["style"] : {};
    const next = { ...current };
    for (const [key, value] of Object.entries(record)) {
      if (value === undefined) delete next[key];
      else next[key] = snapshotValue(value);
    }
    this.#props = Object.freeze({ ...this.#props, style: Object.freeze(next) });
    return this;
  }

  /** Detaches this node while retaining its document ownership, mirroring explicit removal semantics. */
  remove(): this {
    const parent = this.#parent;
    if (parent !== undefined) parent.#detach(this);
    return this;
  }

  /** Moves a node in the same document or attaches a detached node, preserving identity. */
  moveTo(parent: CommonNode, before?: CommonNode): this {
    if (!(parent instanceof CommonContainer)) {
      throw new TypeError("moveTo() target must be an editable container");
    }
    if (before !== undefined && before.#parent !== parent) {
      throw new Error(
        "before must be a node that is a direct child of the target container",
      );
    }
    if (before === this) return this;
    parent.assertCanInsert(this, true);
    const previousParent = this.#parent;
    if (previousParent !== undefined) previousParent.#detach(this);
    parent.insertNode(this, before);
    return this;
  }

  #refreshChildren(): void {
    this.#childrenView = Object.freeze([...this.#children]);
  }

  #detach(child: DeclarationChild): void {
    const index = this.#children.indexOf(child);
    if (index === -1) return;
    this.#children.splice(index, 1);
    if (child instanceof CommonNode) child.#parent = undefined;
    this.#refreshChildren();
  }

  #ownerConflict(owner: CommonNode): boolean {
    if (this.#owner !== undefined && this.#owner !== owner) return true;
    return this.#children.some(
      (child) => child instanceof CommonNode && child.#ownerConflict(owner),
    );
  }

  #assignOwner(owner: CommonNode): void {
    if (this.#owner === undefined) this.#owner = owner;
    for (const child of this.#children) {
      if (child instanceof CommonNode) child.#assignOwner(owner);
    }
  }

  protected assertCanInsert(child: CommonNode, allowAttached = false): void {
    if (child === this) {
      throw new Error("Cannot attach a node inside its own descendant tree");
    }
    for (
      let ancestor = this.#parent;
      ancestor !== undefined;
      ancestor = ancestor.#parent
    ) {
      if (ancestor === child) {
        throw new Error("Cannot attach a node inside its own descendant tree");
      }
    }
    if (!allowAttached && child.#parent !== undefined) {
      throw new Error(
        "Node is already attached; use moveTo() for an explicit same-document move",
      );
    }
    const owner = this.#owner;
    if (owner !== undefined && child.#ownerConflict(owner)) {
      throw new Error("Cannot attach a node owned by another document");
    }
  }

  protected insertNode(child: CommonNode, before?: CommonNode): void {
    this.assertCanInsert(child);
    if (before !== undefined && before.#parent !== this) {
      throw new Error(
        "before must be a node that is a direct child of the target container",
      );
    }
    const owner = this.#owner;
    if (owner !== undefined) child.#assignOwner(owner);
    child.#parent = this;
    const index =
      before === undefined
        ? this.#children.length
        : this.#children.indexOf(before);
    this.#children.splice(index, 0, child);
    this.#refreshChildren();
  }

  protected insertRichTextItems(
    items: readonly RichTextItem[],
    replace: boolean,
  ): void {
    const next = items.map((item) => declarationChild(item));
    const seen = new Set<DeclarationChild>();
    for (const child of next) {
      if (seen.has(child)) {
        throw new Error("The same node cannot be attached twice");
      }
      seen.add(child);
      if (child instanceof CommonNode) this.assertCanInsert(child);
    }
    const owner = this.#owner;
    if (replace) {
      for (const child of this.#children) {
        if (child instanceof CommonNode) child.#parent = undefined;
      }
      this.#children = [];
    }
    for (const child of next) {
      if (child instanceof CommonNode) {
        child.#parent = this;
        if (owner !== undefined) child.#assignOwner(owner);
      }
      this.#children.push(child);
    }
    this.#refreshChildren();
  }
}

/** Editable structural container: deck, slide, section, box, and annotation-only containers. */
export class CommonContainer<
  Kind extends string = string,
  Props extends object = DocumentProps,
  Style extends object = DocumentStyle,
> extends CommonNode<Kind, Props, Style> {
  append(...children: readonly CommonNode[]): this {
    const seen = new Set<CommonNode>();
    for (const child of children) {
      if (!(child instanceof CommonNode)) {
        throw new TypeError("append() accepts editable document nodes only");
      }
      if (seen.has(child)) {
        throw new Error("The same node cannot be attached twice");
      }
      seen.add(child);
      this.assertCanInsert(child);
    }
    for (const child of children) this.insertNode(child);
    return this;
  }
}

/** Rich-text container whose children may also be literal text. */
export class RichTextContainer<
  Kind extends string = string,
  Props extends object = DocumentProps,
  Style extends object = DocumentStyle,
> extends CommonContainer<Kind, Props, Style> {
  override append(...items: readonly RichTextItem[]): this {
    this.insertRichTextItems(items, false);
    return this;
  }

  /** Replaces all rich-text content atomically after validating every new item. */
  content(...items: readonly RichTextItem[]): this {
    this.insertRichTextItems(items, true);
    return this;
  }
}

/** Public type-narrowing predicate used by scoped author queries. */
export type AuthorNodeGuard<Node extends CommonNode> = (
  node: CommonNode,
) => node is Node;

/** Explicit failures raised by scoped author queries instead of returning unchecked nodes. */
export class AuthorQueryError extends Error {
  readonly code: "invalid-id" | "missing-id" | "ambiguous-id" | "type-mismatch";

  constructor(code: AuthorQueryError["code"], message: string) {
    super(message);
    this.code = code;
    this.name = "AuthorQueryError";
  }
}

function assertAuthorId(id: string): void {
  if (typeof id !== "string" || id.trim() === "") {
    throw new AuthorQueryError(
      "invalid-id",
      "Author id must be a non-empty string",
    );
  }
}

function authorIdOf(node: CommonNode): string | undefined {
  const value = node.props["id"];
  return typeof value === "string" && value.trim() !== "" ? value : undefined;
}

function authorPath(node: CommonNode): string {
  const segments: string[] = [];
  for (
    let current: CommonNode | undefined = node;
    current !== undefined;
    current = current.parent
  ) {
    if (current.parent === undefined && current.kind === "deck") {
      segments.unshift("deck");
      continue;
    }
    const id = authorIdOf(current);
    segments.unshift(
      id === undefined ? current.kind : `${current.kind}[${id}]`,
    );
  }
  return segments.join("/");
}

/**
 * Author-facing locator over the live editable document.
 *
 * A query is scoped to one node. Lookups traverse anonymous layout nodes but stop at the next node
 * declaring an author id, so callers descend explicitly through named scopes. Pages always declare
 * their author id and act as the root author scope. Returned `.node` values are the same editable
 * instances the document holds; they are never build observations.
 */
export class AuthorQuery<Node extends CommonNode = CommonNode> {
  readonly #node: Node;

  constructor(node: Node) {
    this.#node = node;
  }

  /** Editable node at this query scope. */
  get node(): Node {
    return this.#node;
  }

  /** Readable path derived from the current parent chain and declared author ids. */
  get path(): string {
    return authorPath(this.#node);
  }

  /** Returns the single editable node named `id` directly in this scope. */
  getNode(id: string): CommonNode {
    assertAuthorId(id);
    return this.#one(id, "getNode").node;
  }

  /** Returns a scoped query handle for the single node named `id`, enabling stepwise lookups. */
  get(id: string): AuthorQuery {
    assertAuthorId(id);
    return this.#one(id, "get");
  }

  /** Lists every node named `id` in this scope; duplicates remain visible to callers before build validation. */
  getAll(id: string): readonly AuthorQuery[] {
    assertAuthorId(id);
    return Object.freeze(
      this.#matches(id).map((match) => new AuthorQuery(match)),
    );
  }

  /**
   * Returns the node named `id` after checking the caller's type predicate.
   *
   * Missing, ambiguous, and type-mismatched results throw `AuthorQueryError` with distinct codes so
   * callers never need an unchecked assertion to reach an editable capability node.
   */
  getAs<Selected extends CommonNode>(
    id: string,
    guard: AuthorNodeGuard<Selected>,
  ): Selected {
    const match = this.#one(id, "getAs").node;
    if (!guard(match)) {
      throw new AuthorQueryError(
        "type-mismatch",
        `Author id ${JSON.stringify(id)} does not match the expected node type`,
      );
    }
    return match;
  }

  #one(id: string, operation: string): AuthorQuery {
    assertAuthorId(id);
    const matches = this.#matches(id);
    if (matches.length === 0) {
      throw new AuthorQueryError(
        "missing-id",
        `No author node found for id ${JSON.stringify(id)} in ${this.path}`,
      );
    }
    if (matches.length > 1) {
      throw new AuthorQueryError(
        "ambiguous-id",
        `Author id ${JSON.stringify(id)} is duplicated in ${this.path}`,
      );
    }
    const [match] = matches;
    if (match === undefined) {
      throw new AuthorQueryError(
        "missing-id",
        `${operation} could not resolve author id ${JSON.stringify(id)}`,
      );
    }
    return new AuthorQuery(match);
  }

  #matches(id: string): readonly CommonNode[] {
    const matches: CommonNode[] = [];
    const visit = (parent: CommonNode): void => {
      for (const child of parent.children) {
        if (!(child instanceof CommonNode)) continue;
        const childId = authorIdOf(child);
        if (childId === id) matches.push(child);
        else if (childId === undefined) visit(child);
      }
    };
    visit(this.#node);
    return matches;
  }
}

/** Creates the editable Deck node used by both React host expansion and command authoring. */
export function createDeckNode<Props extends object = DeckProps>(
  props: Props,
): CommonContainer<"deck", Props, TextStyle> {
  return new CommonContainer("deck", "deck", props);
}

/** Creates the editable Slide node used by both React host expansion and command authoring. */
export function createSlideNode<Props extends object = SlideProps>(
  props: Props,
): CommonContainer<"slide", Props, TextStyle & Omit<LayoutStyle, "zIndex">> {
  return new CommonContainer("slide", "slide", props);
}

/** Creates the editable Section node used by both React host expansion and command authoring. */
export function createSectionNode<Props extends object = SectionProps>(
  props: Props,
): CommonContainer<"section", Props> {
  return new CommonContainer("section", "section", props);
}
