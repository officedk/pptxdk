import { BuildError } from "#/diagnostics.ts";
import { CommonNode, type CommonContainer } from "#/document.ts";

function duplicateAuthorId(path: string, id: string): never {
  throw new BuildError([
    {
      severity: "error",
      path,
      code: "duplicate-author-id",
      message: `Duplicate author id ${JSON.stringify(id)} in the same scope`,
    },
  ]);
}

function childPath(
  parentPath: string,
  child: CommonNode,
  index: number,
): string {
  const id = child.id;
  return `${parentPath}/${child.kind}[${id ?? index + 1}]`;
}

/**
 * Validates ids inside one appended subtree against a scope that already contains the ids of the
 * current target structure. A named node starts a fresh scope for its descendants, matching author
 * query traversal and formal build validation.
 */
function validateSubtree(
  node: CommonNode,
  scope: Set<string>,
  path: string,
): void {
  const id = node.id;
  if (id !== undefined) {
    if (scope.has(id)) duplicateAuthorId(`${path}/${node.kind}[${id}]`, id);
    scope.add(id);
  }
  const descendantScope = id === undefined ? scope : new Set<string>();
  node.children.forEach((child, index) => {
    if (child instanceof CommonNode) {
      validateSubtree(child, descendantScope, childPath(path, child, index));
    }
  });
}

function collectAnonymousIds(parent: CommonNode, ids: Set<string>): void {
  for (const child of parent.children) {
    if (!(child instanceof CommonNode)) continue;
    const id = child.id;
    if (id !== undefined) ids.add(id);
    else if (child.kind !== "slide") collectAnonymousIds(child, ids);
  }
}

function scopeOwner(target: CommonNode): CommonNode {
  for (
    let current: CommonNode | undefined = target;
    current !== undefined;
    current = current.parent
  ) {
    if (
      current.id !== undefined ||
      current.kind === "slide" ||
      current.kind === "deck"
    ) {
      return current;
    }
  }
  return target;
}

/** Visits the slides the formal deck model sees directly or through one Section level. */
function deckSlides(deck: CommonNode): readonly CommonNode[] {
  const slides: CommonNode[] = [];
  for (const child of deck.children) {
    if (!(child instanceof CommonNode)) continue;
    if (child.kind === "slide") slides.push(child);
    else if (child.kind === "section") {
      for (const member of child.children) {
        if (member instanceof CommonNode && member.kind === "slide") {
          slides.push(member);
        }
      }
    }
  }
  return slides;
}

function validateDeckAppend(
  deck: CommonNode,
  nodes: readonly CommonNode[],
  basePath: string,
): void {
  const slideIds = new Set<string>();
  for (const slide of deckSlides(deck)) {
    const id = slide.id;
    if (id !== undefined) slideIds.add(id);
  }
  const visit = (node: CommonNode, path: string): void => {
    if (node.kind === "slide") {
      const id = node.id;
      if (id !== undefined) {
        if (slideIds.has(id))
          duplicateAuthorId(`${path}/${node.kind}[${id}]`, id);
        slideIds.add(id);
      }
      node.children.forEach((child, index) => {
        if (child instanceof CommonNode) {
          validateSubtree(
            child,
            new Set<string>(),
            childPath(path, child, index),
          );
        }
      });
    } else if (node.kind === "section") {
      node.children.forEach((child, index) => {
        if (child instanceof CommonNode) {
          visit(child, childPath(path, child, index));
        }
      });
    }
  };
  for (const node of nodes) {
    visit(node, basePath);
  }
}

/**
 * Preflights the ids introduced by one append against the target's current structure.
 *
 * The target can be edited while the fragment expands, so the check runs only when the caller is
 * about to attach. It rejects conflicts caused by this fragment without re-validating unrelated
 * existing content; the formal build remains responsible for all other diagnostics.
 */
export function assertAppendableAuthorIds(
  target: CommonContainer,
  nodes: readonly CommonNode[],
): void {
  if (nodes.length === 0) return;
  const owner = scopeOwner(target);
  const basePath = target.query.path;
  if (owner.kind === "deck") {
    validateDeckAppend(owner, nodes, basePath);
    return;
  }
  const currentIds = new Set<string>();
  for (const child of owner.children) {
    if (!(child instanceof CommonNode)) continue;
    if (child.id !== undefined) currentIds.add(child.id);
    else if (child.kind !== "slide") collectAnonymousIds(child, currentIds);
  }
  for (const node of nodes) {
    validateSubtree(node, currentIds, basePath);
  }
}
