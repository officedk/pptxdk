import type { Transition } from "@react-pptx/capabilities/transition";
import { Slide as CoreSlide } from "@react-pptx/core";
import type { SlideProps as CoreSlideProps } from "@react-pptx/core";
import type { ReactNode } from "react";

export type SlideProps = Omit<CoreSlideProps, "transition"> & {
  readonly transition?: Transition;
};

export const Slide = CoreSlide as (props: SlideProps) => ReactNode;
