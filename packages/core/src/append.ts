import { validateAuthorIdScopes } from "#/author-ids.ts";
import type { DeclarationNode } from "#/declaration.ts";
import { CommonNode, type CommonContainer } from "#/document.ts";

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

/** Collects the slide ids the formal deck model sees directly or through one Section level. */
function deckSlideIds(deck: CommonNode): Set<string> {
  const ids = new Set<string>();
  const add = (slide: CommonNode): void => {
    if (slide.id !== undefined) ids.add(slide.id);
  };
  for (const child of deck.children) {
    if (!(child instanceof CommonNode)) continue;
    if (child.kind === "slide") add(child);
    else if (child.kind === "section") {
      for (const member of child.children) {
        if (member instanceof CommonNode && member.kind === "slide") {
          add(member);
        }
      }
    }
  }
  return ids;
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
  const currentIds =
    owner.kind === "deck" ? deckSlideIds(owner) : new Set<string>();
  if (owner.kind !== "deck") {
    for (const child of owner.children) {
      if (!(child instanceof CommonNode)) continue;
      if (child.id !== undefined) currentIds.add(child.id);
      else if (child.kind !== "slide") collectAnonymousIds(child, currentIds);
    }
  }
  // Reuse the formal validator for the appended roots; the current scope set supplies the
  // existing target ids, so the check still rejects only conflicts introduced by this append.
  const appended: DeclarationNode = {
    type: owner.type,
    props: {},
    children: nodes,
  };
  validateAuthorIdScopes(appended, target.query.path, currentIds);
}
