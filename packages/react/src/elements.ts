import type {
  DeckProps as CoreDeckProps,
  SectionProps as CoreSectionProps,
  SlideLayoutProps as CoreSlideLayoutProps,
  SlideMasterProps as CoreSlideMasterProps,
  SlideProps as CoreSlideProps,
} from "@pptxdk/core";
import {
  createDeckNode,
  createSectionNode,
  createSlideLayoutNode,
  createSlideMasterNode,
  createSlideNode,
} from "@pptxdk/core/document";
import { hostNodeCreator, type DocumentNodeCreator } from "@pptxdk/core/host";
import { createElement, type ReactNode } from "react";

export type ElementProps = {
  readonly children?: ReactNode;
};

export type DeckProps = CoreDeckProps & ElementProps;
export type SlideMasterProps = CoreSlideMasterProps & ElementProps;
export type SlideLayoutProps = CoreSlideLayoutProps & ElementProps;
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

export const Deck = hostElement<DeckProps>("deck", createDeckNode);
export const SlideMaster = hostElement<SlideMasterProps>(
  "slide-master",
  createSlideMasterNode,
);
export const SlideLayout = hostElement<SlideLayoutProps>(
  "slide-layout",
  createSlideLayoutNode,
);
export const Slide = hostElement<SlideProps>("slide", createSlideNode);
export const Section = hostElement<SectionProps>("section", createSectionNode);
