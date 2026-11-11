import {
  componentLabel,
  isDeclarationNode,
  type DeclarationNode,
} from "#/declaration.ts";
import { BuildError } from "#/diagnostics.ts";

/**
 * Reads a declared author id, rejecting present but invalid values.
 *
 * Anonymous nodes omit the id. Present values must be non-empty strings; the formal build reports
 * them once through the same path as other declaration errors.
 */
export function readAuthorId(value: unknown, path: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new BuildError([
      {
        severity: "error",
        path,
        code: "invalid-author-id",
        message: "id must be a non-empty string",
      },
    ]);
  }
  return value;
}

/**
 * Validates author ids on one page using the same scopes the author query traverses.
 *
 * A node's own id belongs to its parent scope; its descendants belong to a new scope only when the
 * node declares an id. Anonymous layout nodes therefore stay transparent, while each page is the
 * root author scope.
 */
export function validateAuthorIdScopes(
  node: DeclarationNode,
  path: string,
  scope: Set<string>,
): void {
  node.children.forEach((child, index) => {
    if (!isDeclarationNode(child)) return;
    const childPath = `${path}/${componentLabel(child, index)}`;
    const id = readAuthorId(child.props["id"], childPath);
    if (id !== undefined) {
      if (scope.has(id)) {
        throw new BuildError([
          {
            severity: "error",
            path: childPath,
            code: "duplicate-author-id",
            message: `Duplicate author id ${JSON.stringify(id)} in the same scope`,
          },
        ]);
      }
      scope.add(id);
    }
    validateAuthorIdScopes(
      child,
      childPath,
      id === undefined ? scope : new Set(),
    );
  });
}
