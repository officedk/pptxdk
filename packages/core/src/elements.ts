import { createElement, type ReactNode } from "react";

import {
  createDeckNode,
  createSectionNode,
  createSlideNode,
  hostNodeCreator,
  type DocumentNodeCreator,
} from "#/document.ts";
import type { DeckProps, SectionProps, SlideProps } from "#/types.ts";

function hostElement<P extends { readonly children?: ReactNode }>(
  type: string,
  create: DocumentNodeCreator<P>,
) {
  return function HostElement(props: P): ReactNode {
    return createElement(
      type,
      { ...props, [hostNodeCreator]: create },
      props.children,
    );
  };
}

export const Deck = hostElement<DeckProps>("rp:deck", createDeckNode);
export const Slide = hostElement<SlideProps>("rp:slide", createSlideNode);
export const Section = hostElement<SectionProps>(
  "rp:section",
  createSectionNode,
);
