import { createElement, type ReactNode } from "react";

import type { DeckProps, SlideProps } from "#/types.ts";

function hostElement<P extends { readonly children?: ReactNode }>(
  type: string,
) {
  return function HostElement(props: P): ReactNode {
    return createElement(type, props, props.children);
  };
}

export const Deck = hostElement<DeckProps>("rp:deck");
export const Slide = hostElement<SlideProps>("rp:slide");
