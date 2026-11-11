import type { DeclarationChild, DeclarationContainer } from "#/declaration.ts";

export type HostText = {
  readonly type: "#text";
  text: string;
  hidden: boolean;
};

export class HostNode {
  readonly type: string;
  props: Readonly<Record<string, unknown>>;
  readonly children: HostChild[] = [];
  fiber: unknown;
  hidden = false;

  constructor(
    type: string,
    props: Readonly<Record<string, unknown>>,
    fiber: unknown,
  ) {
    this.type = type;
    this.props = props;
    this.fiber = fiber;
  }
}

export type HostChild = HostNode | HostText;
export type HostContainer = { readonly children: HostChild[] };

/**
 * Copies the completed host tree into plain declarations before React cleanup.
 *
 * Unmounting deletes host nodes from their parents, so the formal build must not read the live tree
 * after cleanup. The snapshot shares already-owned author values and does not create another
 * long-lived content model.
 */
export function declarationFromHost(
  container: HostContainer,
): DeclarationContainer {
  return { children: container.children.map(declarationChild) };
}

function declarationChild(child: HostChild): DeclarationChild {
  if (child instanceof HostNode) {
    return {
      type: child.type,
      props: child.props,
      children: child.children.map(declarationChild),
      ...(child.hidden ? { hidden: true } : {}),
    };
  }
  return {
    type: "#text",
    text: child.text,
    ...(child.hidden ? { hidden: true } : {}),
  };
}

export function withoutChildren(
  props: Readonly<Record<string, unknown>>,
): Readonly<Record<string, unknown>> {
  if (!("children" in props)) return props;
  const { children: _children, ...rest } = props;
  return rest;
}
