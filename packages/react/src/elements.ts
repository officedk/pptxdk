import type {
  DeckProps as CoreDeckProps,
  SectionProps as CoreSectionProps,
  SlideProps as CoreSlideProps,
} from "@react-pptx/core";
import {
  createDeckNode,
  createSectionNode,
  createSlideNode,
} from "@react-pptx/core/document";
import {
  hostNodeCreator,
  type DocumentNodeCreator,
} from "@react-pptx/core/host";
import { createElement, type ReactNode } from "react";

export type ElementProps = {
  readonly children?: ReactNode;
};

export type DeckProps = CoreDeckProps & ElementProps;
export type SlideProps = CoreSlideProps & ElementProps;
export type SectionProps = CoreSectionProps & ElementProps;

function hostElement<P extends ElementProps>(
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
