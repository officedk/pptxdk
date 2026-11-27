import { BuildError, toBuildDiagnostic } from "@pptxdk/core";
import { CommonContainer, CommonNode } from "@pptxdk/core/document";
import { createElement, useRef, type ReactNode } from "react";

import { Mount } from "#/mount.ts";

/**
 * Result returned by a container vanilla factory to bridge both an outer node and an inner slot
 * container into React via {@link fromVanilla}.
 */
export type VanillaSlotResult<
  Node extends CommonNode = CommonNode,
  Slot extends CommonContainer = CommonContainer,
> = {
  /** Outer editable node mounted into the enclosing React container. */
  readonly node: Node;
  /** Descendant container (or `node` itself) where React reconciles JSX `children`. */
  readonly slot: Slot;
};

type VanillaOutput = CommonNode | VanillaSlotResult;

function isVanillaSlotResult(value: unknown): value is VanillaSlotResult {
  if (typeof value !== "object" || value === null) return false;
  if (value instanceof CommonNode) return false;
  if (!("node" in value) || !("slot" in value)) return false;
  return (
    value.node instanceof CommonNode && value.slot instanceof CommonContainer
  );
}

function hasNonEmptyChildren(children: unknown): boolean {
  if (
    children === undefined ||
    children === null ||
    typeof children === "boolean"
  ) {
    return false;
  }
  if (Array.isArray(children)) {
    return children.some(hasNonEmptyChildren);
  }
  return true;
}

function shallowEqualProps(
  a: Readonly<Record<string, unknown>>,
  b: Readonly<Record<string, unknown>>,
): boolean {
  const keysA = Object.keys(a);
  const keysB = Object.keys(b);
  if (keysA.length !== keysB.length) return false;
  for (const key of keysA) {
    if (!Object.prototype.hasOwnProperty.call(b, key)) return false;
    if (!Object.is(a[key], b[key])) return false;
  }
  return true;
}

function bridgeError(message: string): BuildError {
  return new BuildError([
    toBuildDiagnostic("mount", "invalid-structure", message),
  ]);
}

/**
 * Wraps a framework-free `@officedk/pptxdk/vanilla` factory function as a first-class React
 * component backed by {@link Mount}.
 *
 * Factories returning a single {@link CommonNode} produce closed components that statically reject
 * `children`, while factories returning `{ node, slot }` produce container components whose JSX
 * `children` are reconciled into `slot`. The factory output is memoized across re-renders and
 * `<Suspense>` retries while non-`children` props remain shallow-equal.
 */
export function fromVanilla<
  Node extends CommonNode = CommonNode,
  Slot extends CommonContainer = CommonContainer,
>(
  factory: () => VanillaSlotResult<Node, Slot>,
): (props: { readonly children?: ReactNode }) => ReactNode;
export function fromVanilla<
  P extends object,
  Node extends CommonNode = CommonNode,
  Slot extends CommonContainer = CommonContainer,
>(
  factory: (props: P) => VanillaSlotResult<Node, Slot>,
): (props: P & { readonly children?: ReactNode }) => ReactNode;
export function fromVanilla<Node extends CommonNode = CommonNode>(
  factory: () => Node,
): (props: { readonly children?: never }) => ReactNode;
export function fromVanilla<
  P extends object,
  Node extends CommonNode = CommonNode,
>(
  factory: (props: P) => Node,
): (props: P & { readonly children?: never }) => ReactNode;
export function fromVanilla(
  factory: (props?: never) => VanillaOutput,
): (props: Record<string, unknown>) => ReactNode {
  function VanillaComponent(props: Record<string, unknown>): ReactNode {
    const { children, ...vanillaProps } = props;
    const cacheRef = useRef<{
      readonly props: Readonly<Record<string, unknown>>;
      readonly output: VanillaOutput;
    } | null>(null);

    let output: VanillaOutput;
    if (
      cacheRef.current !== null &&
      shallowEqualProps(cacheRef.current.props, vanillaProps)
    ) {
      output = cacheRef.current.output;
    } else {
      output = factory(vanillaProps as never);
      cacheRef.current = { props: vanillaProps, output };
    }

    if (isVanillaSlotResult(output)) {
      return createElement(
        Mount,
        { node: output.node, slot: output.slot },
        children as ReactNode,
      );
    }

    if (output instanceof CommonNode) {
      if (hasNonEmptyChildren(children)) {
        throw bridgeError(
          "Closed vanilla component returned a CommonNode and cannot accept children",
        );
      }
      return createElement(Mount, { node: output });
    }

    throw bridgeError(
      "fromVanilla factory must return a CommonNode or a { node, slot } container result",
    );
  }

  VanillaComponent.displayName =
    factory.name !== "" ? `Vanilla(${factory.name})` : "VanillaComponent";
  return VanillaComponent;
}
