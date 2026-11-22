import type {
  SlideLayoutAuthorProps,
  SlideLayoutSource,
  SlideMasterAuthorProps,
} from "@pptxdk/capabilities/document";
import type { Transition } from "@pptxdk/capabilities/transition";
import {
  Slide as ReactSlide,
  SlideLayout as ReactSlideLayout,
  SlideMaster as ReactSlideMaster,
  type ElementProps,
  type SlideProps as ReactSlideProps,
} from "@pptxdk/react";
import type { ReactNode } from "react";

export type SlideProps = Omit<ReactSlideProps, "transition" | "layout"> & {
  readonly layout?: SlideLayoutSource | undefined;
  readonly transition?: Transition;
};

export type SlideMasterProps = SlideMasterAuthorProps & ElementProps;
export type SlideLayoutProps = SlideLayoutAuthorProps & ElementProps;

export const Slide = ReactSlide as (props: SlideProps) => ReactNode;
export const SlideMaster = ReactSlideMaster as (
  props: SlideMasterProps,
) => ReactNode;
export const SlideLayout = ReactSlideLayout as (
  props: SlideLayoutProps,
) => ReactNode;
