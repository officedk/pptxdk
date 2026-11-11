import type { DeclarationChild, DeclarationText } from "#/declaration.ts";
import {
  insertDeclarationChild,
  removeDeclarationChild,
  replaceDeclarationProps,
  setDeclarationText,
} from "#/document-internals.ts";
import { CommonContainer, CommonNode, RichTextContainer } from "#/document.ts";

/** Renderer-side root during a single React expansion; it does not own a second document tree. */
export type HostContainer = { children: DeclarationChild[] };

/** A literal text instance reconciled inside a rich-text common node. */
export class ReactTextNode implements DeclarationText {
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

  [setDeclarationText](text: string): void {
    this.#text = text;
  }
}

/**
 * Structure-only tag mapping for the shared node classes.
 *
 * It selects which common node type a React host tag becomes so public editing types match the
 * declaration; it does not decide capability support or property rules. Unknown or trimmed
 * capabilities are still diagnosed by the formal build.
 */
const richTextDeclarations = new Set([
  "rp:text",
  "rp:paragraph",
  "rp:span",
  "rp:cell",
]);

const structuralDeclarations = new Set([
  "rp:deck",
  "rp:slide",
  "rp:section",
  "rp:box",
  "rp:group",
  "rp:table",
  "rp:row",
  "rp:motion-scope",
  "rp:motion-step",
]);

/**
 * Nodes whose React tree has been handed to a caller. Renderer teardown and late React work must
 * not edit the completed document, so adapter mutations targeting a handed-off entry are ignored.
 * Author operations on the returned nodes are unaffected.
 */
const handedOff = new WeakSet<object>();

function withoutChildren(
  props: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
  if (!("children" in props)) return props;
  const { children: _children, ...rest } = props;
  return rest;
}

function declarationKind(type: string): string {
  return type.startsWith("rp:") ? type.slice(3) : type;
}

export function createHostInstance(
  type: string,
  props: Readonly<Record<string, unknown>>,
): CommonNode {
  const kind = declarationKind(type);
  const declaration = withoutChildren(props);
  if (richTextDeclarations.has(type)) {
    return new RichTextContainer(kind, type, declaration);
  }
  if (structuralDeclarations.has(type)) {
    return new CommonContainer(kind, type, declaration);
  }
  return new CommonNode(kind, type, declaration);
}

export function createHostText(text: string): ReactTextNode {
  return new ReactTextNode(text);
}

export function updateHostText(text: ReactTextNode, next: string): void {
  if (handedOff.has(text)) return;
  text[setDeclarationText](next);
}

export function updateHostInstance(
  instance: CommonNode,
  props: Readonly<Record<string, unknown>>,
): void {
  if (handedOff.has(instance)) return;
  // React commits all current JSX props at once, unlike author updates that patch only provided
  // fields. Replace the declaration so fields absent from this commit do not survive.
  instance[replaceDeclarationProps](withoutChildren(props));
}

export function attachHostChild(
  parent: CommonNode,
  child: DeclarationChild,
  before?: DeclarationChild,
): void {
  if (handedOff.has(parent) || handedOff.has(child)) return;
  parent[insertDeclarationChild](child, before);
}

export function detachHostChild(
  parent: CommonNode,
  child: DeclarationChild,
): void {
  if (handedOff.has(parent) || handedOff.has(child)) return;
  parent[removeDeclarationChild](child);
}

export function insertRootChild(
  container: HostContainer,
  child: DeclarationChild,
  before?: DeclarationChild,
): void {
  const current = container.children.indexOf(child);
  if (current !== -1) container.children.splice(current, 1);
  const index = before === undefined ? -1 : container.children.indexOf(before);
  if (index === -1) container.children.push(child);
  else container.children.splice(index, 0, child);
}

export function removeRootChild(
  container: HostContainer,
  child: DeclarationChild,
): void {
  const index = container.children.indexOf(child);
  if (index !== -1) container.children.splice(index, 1);
}

/**
 * Closes the React ownership window on the completed candidate before renderer teardown. The
 * candidate shares its nodes with the returned document; teardown must only detach the renderer
 * root, never clear this content.
 */
export function lockHandoffTree(root: CommonNode): void {
  handedOff.add(root);
  for (const child of root.children) {
    handedOff.add(child);
    if (child instanceof CommonNode) lockHandoffTree(child);
  }
}
