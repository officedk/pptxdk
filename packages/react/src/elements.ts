import type {
  DeckProps as CoreDeckProps,
  SectionProps as CoreSectionProps,
  SlideLayoutProps as CoreSlideLayoutProps,
  SlideMasterProps as CoreSlideMasterProps,
  SlideProps as CoreSlideProps,
} from "@pptxdk/core";
import {
  CommonNode,
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

import { Mount } from "#/mount.ts";

export type ElementProps = {
  readonly children?: ReactNode;
};

/** Props for the root `<Deck>` component; canvas `width` and `height` are in CSS pixels (`1280 × 720` default). */
export type DeckProps = CoreDeckProps & ElementProps;
export type SlideMasterProps = CoreSlideMasterProps & ElementProps;
export type SlideLayoutProps = CoreSlideLayoutProps & ElementProps;
export type SlideProps = CoreSlideProps &
  ElementProps & {
    readonly placeholders?: Readonly<Record<string, unknown>> | undefined;
  };
/** Props for `<Section>`, grouping contiguous child `<Slide>` elements into a native `p14:sectionLst` section. */
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

function isInlineTextNode(node: CommonNode): boolean {
  return (
    node.type === "paragraph" ||
    node.type === "span" ||
    node.type === "slide-number"
  );
}

function expandPlaceholderEntry(
  key: string,
  value: unknown,
  seenObjectNodes: Set<CommonNode>,
): ReactNode {
  if (value === undefined) return null;
  if (value instanceof CommonNode) {
    if (isInlineTextNode(value)) {
      return createElement(
        "text",
        {
          key: `__ph_${key}`,
          placeholder: key,
          [hostNodeCreator]: createPlaceholderTextNode,
        },
        createElement(Mount, { node: value }),
      );
    }
    if (value.parent === undefined && !seenObjectNodes.has(value)) {
      seenObjectNodes.add(value);
      value.set({ placeholder: key });
    }
    return createElement(Mount, { key: `__ph_${key}`, node: value });
  }
  if (Array.isArray(value)) {
    const mapped = value.map((item, index) =>
      item instanceof CommonNode
        ? createElement(Mount, { key: `__ph_${key}_${index}`, node: item })
        : (item as ReactNode),
    );
    return createElement(
      "text",
      {
        key: `__ph_${key}`,
        placeholder: key,
        [hostNodeCreator]: createPlaceholderTextNode,
      },
      ...mapped,
    );
  }
  if (typeof value === "string" || isInlineTextElement(value)) {
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
  const seenObjectNodes = new Set<CommonNode>();
  const placeholderChildren = entries.map(([key, value]) =>
    expandPlaceholderEntry(key, value, seenObjectNodes),
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
