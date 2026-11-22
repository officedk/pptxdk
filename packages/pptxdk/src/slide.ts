import type {
  InferPlaceholderMap,
  JsxPlaceholders,
  SlideLayoutSource,
} from "@pptxdk/capabilities/document";
import type { Transition } from "@pptxdk/capabilities/transition";
import {
  Slide as ReactSlide,
  type SlideProps as ReactSlideProps,
} from "@pptxdk/react";
import type { ReactNode } from "react";

type BaseSlideProps = Omit<
  ReactSlideProps,
  "transition" | "layout" | "placeholders"
> & {
  readonly transition?: Transition;
};

export type BoundSlideProps<L extends SlideLayoutSource = SlideLayoutSource> =
  BaseSlideProps & {
    readonly layout: L;
    readonly placeholders?: JsxPlaceholders<InferPlaceholderMap<L>>;
  };

export type FreeSlideProps = BaseSlideProps & {
  readonly layout?: undefined;
  readonly placeholders?: never;
};

export type SlideProps<
  L extends SlideLayoutSource | undefined = SlideLayoutSource | undefined,
> = [L] extends [undefined]
  ? FreeSlideProps
  : [L] extends [SlideLayoutSource]
    ? BoundSlideProps<L>
    : BoundSlideProps | FreeSlideProps;

export function Slide<
  const L extends SlideLayoutSource | undefined = undefined,
>(props: SlideProps<L>): ReactNode {
  return ReactSlide(props);
}
