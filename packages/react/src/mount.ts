import type { CommonContainer, CommonNode } from "@pptxdk/core/document";
import { createElement, type ReactNode } from "react";

export const hostMountMarker = "__pptxdkMount";
export const hostMountPlaceholder = "__pptxdkMountPlaceholder";
export const mountHostType = "__pptxdk-mount";

/**
 * Props for {@link Mount}, adopting pre-constructed editable nodes into a React tree without
 * introducing a wrapper node.
 */
export type MountProps =
  | {
      /**
       * Pre-constructed editable node, read-only node list, or empty value to attach directly into
       * the enclosing container.
       */
      readonly node: CommonNode | readonly CommonNode[] | null | undefined;
      readonly slot?: undefined;
      readonly children?: undefined;
    }
  | {
      /** Single pre-constructed editable root node attached to the enclosing container. */
      readonly node: CommonNode;
      /**
       * Target container inside `node` (or `node` itself) where JSX `children` are reconciled
       * after any pre-existing imperative children.
       */
      readonly slot: CommonContainer;
      /** JSX children reconciled inside `slot`. */
      readonly children?: ReactNode;
    };

/**
 * Adopts pre-constructed {@link CommonNode} instances directly into the React tree without
 * creating a wrapper node.
 *
 * Adopted nodes retain their own props, identity, and imperative subtree, participate in layout,
 * motion, and theme scopes, and enforce strict single-parent ownership. When `slot` is provided,
 * JSX `children` are reconciled inside `slot` after any imperative preset children.
 */
export function Mount(props: MountProps): ReactNode {
  return createElement(
    mountHostType,
    {
      [hostMountMarker]: true,
      node: props.node,
      slot: props.slot,
    },
    props.children,
  );
}

/**
 * Internal helper used by `<Slide placeholders={{ ... }} />` so object-node placeholder bindings
 * are carried into the commit phase instead of mutating external nodes during React render.
 */
export function createPlaceholderMount(
  placeholder: string,
  node: CommonNode,
): ReactNode {
  return createElement(mountHostType, {
    key: `__ph_${placeholder}`,
    [hostMountMarker]: true,
    [hostMountPlaceholder]: placeholder,
    node,
    slot: undefined,
  });
}
