import type { SlideLayoutSource } from "@pptxdk/capabilities/document";
import type { Transition } from "@pptxdk/capabilities/transition";
import {
  Slide as ReactSlide,
  type SlideProps as ReactSlideProps,
} from "@pptxdk/react";
import type { ReactNode } from "react";

export type SlideProps = Omit<ReactSlideProps, "transition" | "layout"> & {
  readonly layout?: SlideLayoutSource | undefined;
  readonly transition?: Transition;
};

export const Slide = ReactSlide as (props: SlideProps) => ReactNode;
