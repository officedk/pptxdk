import type { Transition } from "@react-pptx/capabilities/transition";
import {
  Slide as ReactSlide,
  type SlideProps as ReactSlideProps,
} from "@react-pptx/react";
import type { ReactNode } from "react";

export type SlideProps = Omit<ReactSlideProps, "transition"> & {
  readonly transition?: Transition;
};

export const Slide = ReactSlide as (props: SlideProps) => ReactNode;
