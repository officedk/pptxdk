import { CommonNode } from "@pptxdk/core/document";
import {
  hostNodeCreator,
  HostTextNode,
  insertDeclarationChild,
  removeDeclarationChild,
  replaceDeclarationProps,
  setDeclarationText,
  type DeclarationChild,
  type DocumentNodeCreator,
} from "@pptxdk/core/host";

/** Renderer-side root during a single React expansion; it does not own a second document tree. */
export type HostContainer = { children: DeclarationChild[] };

/**
 * Nodes whose React tree has been handed to a caller. Renderer teardown and late React work must
 * not edit the completed document, so adapter mutations targeting a handed-off entry are ignored.
 * Author operations on the returned nodes are unaffected.
 */
const handedOff = new WeakSet<object>();

function withoutHostMetadata(
  props: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
  const {
    children: _children,
    [hostNodeCreator]: _create,
    ...declaration
  } = props;
  return declaration;
}

function isDocumentNodeCreator(value: unknown): value is DocumentNodeCreator {
  return typeof value === "function";
}

export function createHostInstance(
  type: string,
  props: Readonly<Record<string, unknown>>,
): CommonNode {
  const create = props[hostNodeCreator];
  const declaration = withoutHostMetadata(props);
  if (isDocumentNodeCreator(create)) return create(declaration);
  return new CommonNode(type, type, declaration);
}

export function createHostText(text: string): HostTextNode {
  return new HostTextNode(text);
}

export function updateHostText(text: HostTextNode, next: string): void {
  if (handedOff.has(text)) return;
  setDeclarationText(text, next);
}

export function updateHostInstance(
  instance: CommonNode,
  props: Readonly<Record<string, unknown>>,
): void {
  if (handedOff.has(instance)) return;
  // React commits all current JSX props at once, unlike author updates that patch only provided
  // fields. Replace the declaration so fields absent from this commit do not survive.
  replaceDeclarationProps(instance, withoutHostMetadata(props));
}

export function attachHostChild(
  parent: CommonNode,
  child: DeclarationChild,
  before?: DeclarationChild,
): void {
  if (handedOff.has(parent) || handedOff.has(child)) return;
  insertDeclarationChild(parent, child, before);
}

export function detachHostChild(
  parent: CommonNode,
  child: DeclarationChild,
): void {
  if (handedOff.has(parent) || handedOff.has(child)) return;
  removeDeclarationChild(parent, child);
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
