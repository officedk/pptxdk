/**
 * Plain authoring declaration tree consumed by the formal build.
 *
 * React components and command authoring produce editable common nodes that expose this structure
 * once; the formal build reads it without executing React lifecycle or depending on React at
 * runtime. Build-time parsing and layout remain derived structures of the current declaration, not
 * another long-lived authoring model.
 */
export type DeclarationText = {
  readonly type: "#text";
  readonly text: string;
};

export type DeclarationNode = {
  readonly type: string;
  readonly props: Readonly<Record<string, unknown>>;
  readonly children: readonly DeclarationChild[];
};

export type DeclarationChild = DeclarationNode | DeclarationText;

export type DeclarationContainer = {
  readonly children: readonly DeclarationChild[];
};

export function isDeclarationNode(child: unknown): child is DeclarationNode {
  return (
    typeof child === "object" &&
    child !== null &&
    "props" in child &&
    "children" in child
  );
}

/** Author-visible identity used in diagnostic labels; the scoped author id is the only identity. */
export function declarationIdentity(
  props: Readonly<Record<string, unknown>>,
): unknown {
  return props["id"];
}

/** Diagnostic label for a child declaration, matching the existing build path vocabulary. */
export function componentLabel(node: DeclarationNode, index: number): string {
  const type = node.type.startsWith("rp:") ? node.type.slice(3) : node.type;
  const identity = declarationIdentity(node.props);
  const label =
    typeof identity === "string" && identity.trim() !== ""
      ? identity
      : String(index + 1);
  return `${type}[${label}]`;
}
