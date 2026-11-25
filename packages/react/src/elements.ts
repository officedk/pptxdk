import type {
  DeckProps as CoreDeckProps,
  SectionProps as CoreSectionProps,
  SlideLayoutProps as CoreSlideLayoutProps,
  SlideMasterProps as CoreSlideMasterProps,
  SlideProps as CoreSlideProps,
} from "@pptxdk/core";
import {
  RichTextContainer,
  createDeckNode,
  createSectionNode,
  createSlideLayoutNode,
  createSlideMasterNode,
  createSlideNode,
} from "@pptxdk/core/document";
import { hostNodeCreator, type DocumentNodeCreator } from "@pptxdk/core/host";
import {
  cloneElement,
  createElement,
  isValidElement,
  type ReactNode,
} from "react";

export type ElementProps = {
  readonly children?: ReactNode;
};

export type DeckProps = CoreDeckProps & ElementProps;
export type SlideMasterProps = CoreSlideMasterProps & ElementProps;
export type SlideLayoutProps = CoreSlideLayoutProps & ElementProps;
export type SlideProps = CoreSlideProps &
  ElementProps & {
    readonly placeholders?: Readonly<Record<string, unknown>> | undefined;
  };
export type SectionProps = CoreSectionProps & ElementProps;

const inlineTextMarker = Symbol.for("@pptxdk/text-inline-element");

function createPlaceholderTextNode(
  props: Readonly<Record<string, unknown>>,
): RichTextContainer {
  return new RichTextContainer("text", "text", props);
}

function isInlineTextElement(value: unknown): boolean {
  if (!isValidElement(value)) return false;
  if (typeof value.type === "function") {
    return inlineTextMarker in value.type;
  }
  return (
    value.type === "paragraph" ||
    value.type === "span" ||
    value.type === "slide-number"
  );
}

function expandPlaceholderEntry(key: string, value: unknown): ReactNode {
  if (value === undefined) return null;
  if (
    typeof value === "string" ||
    Array.isArray(value) ||
    isInlineTextElement(value)
  ) {
    return createElement(
      "text",
      {
        key: `__ph_${key}`,
        placeholder: key,
        [hostNodeCreator]: createPlaceholderTextNode,
      },
      value as ReactNode,
    );
  }
  if (isValidElement(value)) {
    const existingProps = value.props as Record<string, unknown>;
    return cloneElement(value, {
      ...(value.key === null ? { key: `__ph_${key}` } : {}),
      placeholder: key,
    } as typeof existingProps);
  }
  return null;
}

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

export function Slide(props: SlideProps): ReactNode {
  const { placeholders, children, ...rest } = props;
  if (placeholders === undefined) {
    return createElement(
      "slide",
      { ...rest, [hostNodeCreator]: createSlideNode },
      children,
    );
  }
  const entries = Object.entries(placeholders).filter(
    ([, value]) => value !== undefined,
  );
  const keepEmpty = props.layout === undefined && entries.length === 0;
  const placeholderChildren = entries.map(([key, value]) =>
    expandPlaceholderEntry(key, value),
  );
  return createElement(
    "slide",
    {
      ...(keepEmpty ? { ...rest, placeholders } : rest),
      [hostNodeCreator]: createSlideNode,
    },
    ...placeholderChildren,
    children,
  );
}

export const Section = hostElement<SectionProps>("section", createSectionNode);
