/**
 * Internal host-adapter protocol for React reconciliation and JSX host element binding.
 *
 * Solely for the React host adapter (`@pptxdk/react`) and JSX host elements in
 * `@pptxdk/capabilities`; this subpath is not an author API and must never be re-exported from
 * `.` or `./document`.
 */
import { validateAuthorIdScopes } from "#/author-ids.ts";
import type {
  DeclarationChild,
  DeclarationNode,
  DeclarationText,
} from "#/declaration.ts";
import {
  insertDeclarationChild as insertChildSymbol,
  removeDeclarationChild as removeChildSymbol,
  replaceDeclarationProps as replacePropsSymbol,
  setDeclarationText as setTextSymbol,
} from "#/document-internals.ts";
import type { CommonNode } from "#/document.ts";

/**
 * Internal React host prop carrying the editable-node creator for one declaration type.
 *
 * A JSX host element knows which common node class its capability declares. The adapter removes
 * this value before the declaration reaches the formal build, so it is not author content and the
 * adapter does not need a separate tag-to-node-class table.
 */
export const hostNodeCreator = "__pptxdkNodeCreator";

/**
 * Creates the editable node shared by a command factory and the matching JSX host element.
 *
 * Both authoring entries use the same creator, keeping the node identity returned by expansion and
 * factory construction aligned with the operations their public types expose.
 */
export type DocumentNodeCreator<
  Props extends object = Readonly<Record<string, unknown>>,
  Node extends CommonNode = CommonNode,
> = (props: Props) => Node;

/** A literal text instance reconciled inside a rich-text common node. */
export class HostTextNode implements DeclarationText {
  #text: string;

  constructor(text: string) {
    this.#text = text;
  }

  get type(): "#text" {
    return "#text";
  }

  get text(): string {
    return this.#text;
  }

  [setTextSymbol](text: string): void {
    this.#text = text;
  }
}

export function replaceDeclarationProps(
  node: CommonNode,
  props: Readonly<Record<string, unknown>>,
): void {
  node[replacePropsSymbol](props);
}

export function insertDeclarationChild(
  parent: CommonNode,
  child: DeclarationChild,
  before?: DeclarationChild,
): void {
  parent[insertChildSymbol](child, before);
}

export function removeDeclarationChild(
  parent: CommonNode,
  child: DeclarationChild,
): void {
  parent[removeChildSymbol](child);
}

export function setDeclarationText(node: HostTextNode, text: string): void {
  node[setTextSymbol](text);
}

export { validateAuthorIdScopes };
export type { DeclarationChild, DeclarationNode, DeclarationText };
