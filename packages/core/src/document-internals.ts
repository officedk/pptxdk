import type { DeclarationNode } from "#/declaration.ts";

/**
 * Operation keys shared by the editable document and the React host adapter.
 *
 * The adapter must insert, reorder, and replace complete JSX declarations without exposing those
 * reconciliation operations as author API. These symbols are internal to the core package and are
 * not re-exported by the public document entry.
 */
export const replaceDeclarationProps = Symbol("pptxdk.replaceDeclarationProps");
export const insertDeclarationChild = Symbol("pptxdk.insertDeclarationChild");
export const removeDeclarationChild = Symbol("pptxdk.removeDeclarationChild");
export const setDeclarationText = Symbol("pptxdk.setDeclarationText");

const internalPropsByNode = new WeakMap<
  object,
  Readonly<Record<string, unknown>>
>();

/** Records a `CommonNode`'s internal frozen props snapshot without exposing a symbol method on the class. */
export function registerInternalDeclarationProps(
  node: object,
  props: Readonly<Record<string, unknown>>,
): void {
  internalPropsByNode.set(node, props);
}

/**
 * Reads a declaration node's internal props view for read-only pipeline stages without triggering
 * `CommonNode.props`'s defensive `cloneExternalValue` copy, preserving internal `Uint8Array`
 * identity so `bytesSha256` memoizes across `resolve`, `emit`, and `codegen`.
 */
export function readInternalDeclarationProps(
  node: DeclarationNode,
): Readonly<Record<string, unknown>> {
  return internalPropsByNode.get(node) ?? node.props;
}
