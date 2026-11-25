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
