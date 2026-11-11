import type { DeclarationChild, DeclarationNode } from "#/declaration.ts";
import { createFormalBuild } from "#/formal-build.ts";
import { snapshotValue } from "#/input-values.ts";
import type { BuildOptions, BuildResult, CapabilityModule } from "#/types.ts";

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
  #props: Readonly<Record<string, unknown>>;
  #children: DeclarationChild[] = [];
  #childrenView: readonly DeclarationChild[] = Object.freeze([]);
  #parent: CommonNode | undefined;
  #owner: CommonNode | undefined;

  constructor(kind: Kind, type: string, props: Props = {} as Props) {
    this.kind = kind;
    this.type = type;
    this.#props = normalizeProps(props as Readonly<Record<string, unknown>>);
    if (kind === "deck") this.#owner = this;
  }

  /** Isolated declaration values for the formal build; mutating the returned value cannot edit the node. */
  get props(): Readonly<Record<string, unknown>> {
    return snapshotValue(this.#props);
  }

  /** Read-only child declaration view; node entries keep their editable identity. */
  get children(): readonly DeclarationChild[] {
    return this.#childrenView;
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
    const current = isRecord(this.#props["style"])
      ? (this.#props["style"] as Record<string, unknown>)
      : {};
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

export function createDocumentBuild(modules: readonly CapabilityModule[]) {
  const formalBuild = createFormalBuild(modules);
  return async function buildDocument(
    document: CommonNode,
    options: BuildOptions = {},
  ): Promise<BuildResult> {
    // Consume the current declaration synchronously before this async function returns its Promise.
    // Caller edits after the call therefore affect only later builds.
    return formalBuild({ children: [document] }, options);
  };
}
