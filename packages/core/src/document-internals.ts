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
/** Internal read-only accessor symbol returning a CommonNode's stored props without re-cloning. */
export const readDeclarationProps = Symbol("pptxdk.readDeclarationProps");

/**
 * Reads a declaration node's internal props view for read-only pipeline stages without triggering
 * `CommonNode.props`'s defensive `snapshotValue` clone, preserving internal `Uint8Array` identity
 * so `bytesSha256` memoizes across `resolve`, `emit`, and `codegen`.
 */
export function readInternalDeclarationProps(
  node: DeclarationNode,
): Readonly<Record<string, unknown>> {
  const reader = (
    node as DeclarationNode & {
      readonly [readDeclarationProps]?: () => Readonly<Record<string, unknown>>;
    }
  )[readDeclarationProps];
  return reader !== undefined ? reader.call(node) : node.props;
}
